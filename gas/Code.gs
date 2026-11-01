/**
 * LPS Batik – API sinkron (Google Apps Script + Google Sheets).
 *
 * Pasang: lihat lps-pwa/gas/README.md
 *  1. Buat Google Sheet baru → Extensions → Apps Script → tempel file ini.
 *  2. Jalankan fungsi setup() sekali (izinkan akses). Kode akses tim muncul di log.
 *  3. Deploy → New deployment → Web app → Execute as: Me, Who has access: Anyone.
 *  4. Salin URL /exec ke aplikasi (⚙ → Sinkron tim).
 *
 * Protokol (POST body JSON, Content-Type text/plain):
 *   { key, user, since, tasks:[...], daily:[...], masterVer, master? }
 *   → { ok, rev, tasks:[...], daily:[...], applied, skipped, masterVer, master? }
 * master = jadwal & referensi dari Excel (meta, pekerjaan, areas, …; lihat tools/push_master.py),
 * disimpan di tab Master agar aplikasi yang di-hosting publik tidak membawa data proyek.
 * Server mengirim master hanya bila masterVer milik perangkat berbeda.
 * Record membawa _u (waktu ubah ISO dari perangkat), _by (nama), _del (true = dihapus).
 * Konflik: _u terbaru menang. Setiap tulis diberi _rev (nomor urut server) untuk tarik inkremental.
 */

var KAT = ['eng', 'mat', 'mp', 'oth'];
var KFIELD = ['st', 'issue', 'detail', 'pic', 'target'];
var MASTER_SHEET = 'Master', CHUNK = 40000; // batas isi satu sel Sheets 50.000 karakter
var TABLES = {
  tasks: {
    sheet: 'Tasks', key: 'kode',
    cols: ['kode', 'minggu', 'disiplin', 'pekerjaan', 'area', 'uraian', 'lokasi', 'pic', 'mulai', 'selesai', 'disepakati', 'alasan']
      .concat(KAT.reduce(function (a, k) { return a.concat(KFIELD.map(function (f) { return 'k.' + k + '.' + f; })); }, []))
      .concat(['catatanKendala', 'penyebabTerlambat', 'inLPS', '_u', '_by', '_del', '_rev'])
  },
  daily: {
    sheet: 'DailyLog', key: 'id',
    cols: ['id', 'tgl', 'kode', 'target', 'tercapai', 'progress', 'penyebab', 'catatan', '_u', '_by', '_del', '_rev']
  }
};

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(TABLES).forEach(function (t) {
    var T = TABLES[t];
    var sh = ss.getSheetByName(T.sheet) || ss.insertSheet(T.sheet);
    sh.getRange(1, 1, sh.getMaxRows(), T.cols.length).setNumberFormat('@'); // teks polos: tanggal tidak diubah Sheets
    sh.getRange(1, 1, 1, T.cols.length).setValues([T.cols]).setFontWeight('bold');
    sh.setFrozenRows(1);
  });
  var ms = ss.getSheetByName(MASTER_SHEET) || ss.insertSheet(MASTER_SHEET);
  ms.getRange(1, 1, ms.getMaxRows(), 1).setNumberFormat('@');
  var log = ss.getSheetByName('Log') || ss.insertSheet('Log');
  if (log.getLastRow() === 0) log.appendRow(['Waktu', 'User', 'Tabel', 'Kunci', 'Aksi']);
  log.setFrozenRows(1);
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('TEAM_KEY')) props.setProperty('TEAM_KEY', Utilities.getUuid().slice(0, 8).toUpperCase());
  if (!props.getProperty('REV')) props.setProperty('REV', '0');
  Logger.log('Kode akses tim: ' + props.getProperty('TEAM_KEY'));
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (!p.key) return json_({ ok: true, app: 'lps-batik-sync' });
  return json_(handle_({ key: p.key, since: +p.since || 0 }));
}

function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'JSON tidak valid' }); }
  return json_(handle_(body));
}

function handle_(b) {
  var props = PropertiesService.getScriptProperties();
  if (!b.key || b.key !== props.getProperty('TEAM_KEY')) return { ok: false, error: 'Kode akses salah' };
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var rev = +props.getProperty('REV') || 0;
    var res = { ok: true, applied: 0, skipped: 0 };
    var logRows = [];
    Object.keys(TABLES).forEach(function (t) {
      var T = TABLES[t], sh = ss.getSheetByName(T.sheet);
      if (!sh) throw new Error('Sheet ' + T.sheet + ' belum ada – jalankan setup()');
      var n = T.cols.length, last = sh.getLastRow();
      var rows = last > 1 ? sh.getRange(2, 1, last - 1, n).getValues() : [];
      var iKey = T.cols.indexOf(T.key), iU = T.cols.indexOf('_u'), iRev = T.cols.indexOf('_rev');
      var idx = {};
      rows.forEach(function (r, i) { idx[String(r[iKey])] = i; });
      var changed = false;
      (b[t] || []).forEach(function (rec) {
        var k = String(rec[T.key] || '');
        if (!k) return;
        var i = idx[k];
        if (i !== undefined && String(rows[i][iU]) >= String(rec._u || '')) { res.skipped++; return; } // server lebih baru
        var row = flatten_(rec, T.cols);
        row[iRev] = ++rev;
        if (i === undefined) { idx[k] = rows.length; rows.push(row); } else rows[i] = row;
        changed = true; res.applied++;
        logRows.push([new Date(), rec._by || b.user || '', T.sheet, k, rec._del ? 'hapus' : (i === undefined ? 'baru' : 'ubah')]);
      });
      if (changed) sh.getRange(2, 1, rows.length, n).setNumberFormat('@').setValues(rows);
      var since = +b.since || 0;
      res[t] = rows.filter(function (r) { return +r[iRev] > since; }).map(function (r) { return unflatten_(r, T.cols); });
    });
    props.setProperty('REV', String(rev));
    if (b.master && typeof b.master === 'object') {
      writeMaster_(ss, b.master);
      props.setProperty('MASTER_VER', String(b.master.exportedAt || new Date().toISOString()));
      logRows.push([new Date(), b.user || '', MASTER_SHEET, String(b.master.exportedAt || ''), 'jadwal']);
    }
    res.masterVer = props.getProperty('MASTER_VER') || '';
    if (res.masterVer && String(b.masterVer || '') !== res.masterVer) res.master = readMaster_(ss);
    if (logRows.length) {
      var ms = ss.getSheetByName(MASTER_SHEET) || ss.insertSheet(MASTER_SHEET);
  ms.getRange(1, 1, ms.getMaxRows(), 1).setNumberFormat('@');
  var log = ss.getSheetByName('Log');
      log.getRange(log.getLastRow() + 1, 1, logRows.length, 5).setValues(logRows);
    }
    res.rev = rev;
    return res;
  } catch (err) {
    return { ok: false, error: String(err && err.message || err) };
  } finally {
    lock.releaseLock();
  }
}

function writeMaster_(ss, m) {
  var sh = ss.getSheetByName(MASTER_SHEET);
  if (!sh) throw new Error('Sheet ' + MASTER_SHEET + ' belum ada – jalankan setup()');
  var s = JSON.stringify(m), parts = [];
  for (var i = 0; i < s.length; i += CHUNK) parts.push([s.slice(i, i + CHUNK)]);
  var last = sh.getLastRow();
  if (last) sh.getRange(1, 1, last, 1).setValues(Array.from({ length: last }, function () { return ['']; }));
  sh.getRange(1, 1, parts.length, 1).setNumberFormat('@').setValues(parts);
}

function readMaster_(ss) {
  var sh = ss.getSheetByName(MASTER_SHEET), last = sh ? sh.getLastRow() : 0;
  if (!last) return null;
  return JSON.parse(sh.getRange(1, 1, last, 1).getValues().map(function (r) { return r[0]; }).join(''));
}

function flatten_(rec, cols) {
  return cols.map(function (c) {
    var v = c.split('.').reduce(function (o, k) { return o == null ? undefined : o[k]; }, rec);
    if (v === null || v === undefined) return '';
    if (typeof v === 'boolean') return v ? 'TRUE' : '';
    return String(v);
  });
}

function unflatten_(row, cols) {
  var o = {};
  cols.forEach(function (c, i) {
    var ks = c.split('.'), cur = o;
    for (var j = 0; j < ks.length - 1; j++) cur = cur[ks[j]] = cur[ks[j]] || {};
    cur[ks[ks.length - 1]] = row[i] === '' ? null : String(row[i]);
  });
  return o;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
