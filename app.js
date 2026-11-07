'use strict';
/* LPS Batik – PWA monitoring Last Planner System.
   Data = seed.json (hasil tools/export_excel.py) + perubahan lapangan di localStorage.
   Versi yang di-hosting publik tidak membawa seed.json: jadwal (master) diambil dari Google Sheet tim setelah kode akses diisi.
   Peta Excel → field ada di PETA-EXCEL-KE-APLIKASI.md */

const KEY = 'lps-batik-v1';
const KAT = { eng: 'Engineering / HSE', mat: 'Material & Alat', mp: 'Man Power', oth: 'Other' };
const KK = Object.keys(KAT);
const BLN = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];
const VIEWS = ['dash', 'jadwal', 'komit', 'ready', 'harian', 'issue', 'sd'];

const $ = s => document.querySelector(s);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const opt = (v, cur, label) => `<option value="${esc(v)}"${String(v) === String(cur ?? '') ? ' selected' : ''}>${esc(label ?? v)}</option>`;
const pct = (x, d = 0) => x == null || isNaN(x) ? '–' : (x * 100).toFixed(d) + '%';

let D = null;
const UI = { view: 'dash', week: '', jSearch: '', jDis: '', jLook: '2', jFrom: null, harianTgl: null, focus: null, fPrio: '', sd: 'mp', mpTgl: null, mpAll: false, dokF: 'open', dokScope: '', dokJenis: '', dokSub: '', dokQ: '', dokTutup: [] };

/* ---------- tanggal ---------- */
const todayISO = () => { const d = new Date(); return new Date(d - d.getTimezoneOffset() * 6e4).toISOString().slice(0, 10); };
const tgl = () => D.settings.tglMode === 'auto' ? todayISO() : (D.settings.tglFixed || D.meta.tglUpdate || todayISO());
const fmt = iso => { if (!iso) return '–'; const [y, m, d] = iso.split('-'); return `${d}-${BLN[+m - 1]}-${y.slice(2)}`; };
const days = (a, b) => Math.round((Date.parse(b) - Date.parse(a)) / 864e5);
const addDays = (iso, n) => new Date(Date.parse(iso) + n * 864e5).toISOString().slice(0, 10);
/* = XLOOKUP(tgl, MASTER!X9:BU9, MASTER!X8:BU8, , -1): minggu terakhir yang start ≤ tanggal */
function weekOf(iso) {
  if (!iso) return null;
  let w = null;
  for (const x of D.meta.weeks) if (x.start <= iso) w = x.w;
  return w ?? D.meta.weeks[0].w;
}
const weekObj = w => D.meta.weeks.find(x => x.w === +w);

/* ---------- logika LPS (setara rumus Excel) ---------- */
/* Daily Log = rumus tblDaily. Bobot & % tercapai dalam persen (Excel: pecahan).
   Kontribusi = bobot × % tercapai; progress kumulatif = MIN(100, Σ kontribusi tanggal sebelumnya + kontribusi hari ini) */
const nOr = v => v === null || v === undefined || v === '' || isNaN(v) ? null : +v;
const r6 = v => Math.round(v * 1e6) / 1e6;
const kontrib = x => nOr(x.bobot) == null || nOr(x.capaian) == null ? 0 : r6(x.bobot * x.capaian / 100);
const progKemarin = x => r6(D.daily.reduce((s, y) => y.kode === x.kode && y.tgl < x.tgl ? s + kontrib(y) : s, 0));
// Beberapa janji untuk task yang sama di hari yang sama: janji sebelumnya (urut id, sama di semua perangkat) ikut dihitung
const samaHari = x => D.daily.filter(y => y !== x && y.kode === x.kode && y.tgl === x.tgl && String(y.id) < String(x.id));
const progSebelum = x => r6(progKemarin(x) + samaHari(x).reduce((s, y) => s + kontrib(y), 0));
const sisaTask = x => r6(100 - progKemarin(x) - samaHari(x).reduce((s, y) => s + (nOr(y.bobot) || 0), 0)); // untuk cek bobot
const progKum = x => Math.min(100, r6(progSebelum(x) + kontrib(x)));
/* sisa bobot yang belum direncanakan = 100 − progress − bobot janji yang belum dilaporkan */
const belumLapor = kode => r6(D.daily.reduce((s, y) => y.kode === kode && nOr(y.capaian) == null ? s + (nOr(y.bobot) || 0) : s, 0));
const sisaRencana = kode => Math.max(0, r6(100 - (prog(kode) || 0) - belumLapor(kode)));
/* Isian sore = actual progress hari ini (% dari task). Disimpan sebagai capaian = actual ÷ bobot (sama dgn "% Tercapai Harian" Excel) */
const aktual = x => nOr(x.bobot) == null || nOr(x.capaian) == null ? null : r6(x.bobot * x.capaian / 100);
const tercapai = x => nOr(x.capaian) == null ? '' : x.capaian >= 100 ? 'Ya' : 'Tidak';
const pc = v => v == null ? '–' : (Math.round(v * 10) / 10) + '%';
/* = MAXIFS(tblDaily[Progress Kumulatif], tblDaily[Kode Task], kode) */
function prog(kode) {
  let m = null;
  for (const x of D.daily) if (x.kode === kode) m = Math.max(m ?? 0, progKum(x));
  return m;
}
const stK = t => KK.some(k => t.k[k].st === 'Pending') ? 'PENDING' : 'READY';
function stP(t) {
  if (stK(t) !== 'READY') return '';
  if ((prog(t.kode) || 0) >= 100) return 'FINISH';
  if (t.selesai && tgl() > t.selesai) return 'LATE / DELAY';
  return 'ON PROGRESS';
}
const inW = w => t => w === '' || +t.minggu === +w;
const inWeek = t => inW(UI.week)(t);
const planTasks = () => D.tasks.filter(t => t.disepakati === 'Ya' && inWeek(t));
function stJadwal(p) {
  const d = tgl();
  if (!p.mulai) return 'Tanpa jadwal';
  if (p.mulai > d) return 'Belum mulai';
  if (p.selesai < d) return 'Selesai (jadwal)';
  return 'Berjalan';
}
function issueList(w = UI.week) {
  const out = [];
  for (const t of D.tasks) {
    if (t.disepakati !== 'Ya' || !inW(w)(t)) continue;
    for (const k of KK) {
      const c = t.k[k];
      if (c.st !== 'Pending') continue;
      const sisa = c.target ? days(tgl(), c.target) : null;
      const st = sisa == null ? 'Belum ada target' : sisa < 0 ? '⚠ LEWAT TARGET' : sisa === 0 ? 'Hari ini' : 'On track';
      out.push({ t, k, c, sisa, st });
    }
  }
  return out.sort((a, b) => (a.c.target || '9999') < (b.c.target || '9999') ? -1 : 1);
}
function ppcDay(date) {
  const rows = D.daily.filter(x => x.tgl === date);
  const ya = rows.filter(x => tercapai(x) === 'Ya').length, lap = rows.filter(x => tercapai(x));
  return { n: rows.length, ya, tidak: rows.filter(x => tercapai(x) === 'Tidak').length, ppc: rows.length ? ya / rows.length : null,
    rata: lap.length ? lap.reduce((s, x) => s + Math.min(100, +x.capaian), 0) / lap.length / 100 : null }; // rata-rata % tercapai
}
/* PPC janji = sheet "PPC Harian" bag. 2 (PPC MINGGUAN): tercapai ÷ dijanjikan dari Daily Log.
   Janji bertanggal setelah tanggal update belum dihitung (direncanakan, belum waktunya). */
function ppcJanji(w) {
  const rows = D.daily.filter(x => x.tgl <= tgl() && (w === '' || weekOf(x.tgl) === +w));
  const ya = rows.filter(x => tercapai(x) === 'Ya').length;
  return { n: rows.length, ya, ppc: rows.length ? ya / rows.length : null };
}
/* = Board 3: task FINISH ÷ task Final (READY) */
function ppcWeek(w) {
  const ts = D.tasks.filter(t => t.disepakati === 'Ya' && (w === '' || +t.minggu === +w) && stK(t) === 'READY');
  const fin = ts.filter(t => stP(t) === 'FINISH').length;
  return { final: ts.length, fin, ppc: ts.length ? fin / ts.length : null };
}
const nextKode = () => 'T-' + String(D.tasks.reduce((m, t) => Math.max(m, +(/(\d+)$/.exec(t.kode) || [0, 0])[1]), 0) + 1).padStart(3, '0');
const label = t => `${t.uraian}${t.pic ? ' (' + t.pic + ')' : ''}`;
const stChip = s => s ? `<span class="chip ${{ READY: 'ok', PENDING: 'warn', FINISH: 'info', 'LATE / DELAY': 'bad', 'ON PROGRESS': 'prog', 'Berjalan': 'prog', 'Belum mulai': '', 'Selesai (jadwal)': 'info', 'On track': 'ok', 'Hari ini': 'warn', '⚠ LEWAT TARGET': 'bad', 'Belum ada target': 'warn', Ya: 'ok', Tidak: 'bad' }[s] ?? ''}">${esc(s)}</span>` : '';

/* ---------- penyimpanan ---------- */
function blankK() { return Object.fromEntries(KK.map(k => [k, { st: 'OK', issue: '', detail: '', pic: '', target: null }])); }
async function fetchSeed() {
  const r = await fetch('data/seed.json', { cache: 'no-cache' });
  if (!r.ok) throw new Error('seed.json tidak ditemukan');
  return r.json();
}
const MASTER_F = ['meta', 'pekerjaan', 'areas', 'areaMap', 'issues', 'penyebab', 'disiplin', 'source', 'exportedAt', 'mpPlan', 'dokSrc', 'timeline'];
const blank = () => ({ meta: { project: '', tglUpdate: null, targetPPC: 0.8, weeks: [], scurve: [] }, pekerjaan: [], areas: [], areaMap: {},
  issues: Object.fromEntries(KK.map(k => [k, []])), penyebab: [], disiplin: ['CSA', 'MEP', 'ARS', 'PRELIM'], tasks: [], daily: [], progress: [], mpPlan: [], manpower: [], dokumen: [], timeline: [], tlEdit: [] });
const hasMaster = () => D.meta.weeks.length > 0;
function applyMaster(m) { // jadwal & referensi diganti; task, daily log, target PPC lokal dipertahankan
  const ppc = hasMaster() ? D.meta.targetPPC : m.meta.targetPPC;
  MASTER_F.forEach(f => { if (m[f] !== undefined) D[f] = m[f]; });
  D.meta = { ...m.meta, targetPPC: ppc ?? 0.8 };
}
const TXT_T = ['disiplin', 'pekerjaan', 'area', 'uraian', 'lokasi', 'pic', 'disepakati', 'alasan', 'catatanKendala', 'penyebabTerlambat', 'prio', 'koord', 'followup'];
/* ---------- prioritas & koordinasi disiplin (task urgent yang perlu follow up detail) ----------
   prio: '' (Normal) | 'Tinggi' | 'Urgent' · koord: disiplin lain yang terlibat, dipisah koma · followup: catatan */
const PRIO_T = ['prio', 'koord', 'followup'];
const prioRank = t => ({ Urgent: 3, Tinggi: 2 })[t.prio] || 1;
const koordList = t => String(t.koord || '').split(',').map(s => s.trim()).filter(Boolean);
const perluFU = t => prioRank(t) >= 2 || koordList(t).length > 0;
const prioChips = t => (t.prio === 'Urgent' ? '<span class="chip bad">URGENT</span>' : t.prio === 'Tinggi' ? '<span class="chip warn">Prioritas tinggi</span>' : '')
  + (koordList(t).length ? `<span class="chip info" title="Koordinasi disiplin lain">↔ ${esc(koordList(t).join(', '))}</span>` : '');
const prioSel = t => `<select class="prio-sel" aria-label="Prioritas" data-f="task|${esc(t.kode)}|prio">${opt('', t.prio, 'Prioritas: normal')}${opt('Tinggi', t.prio, 'Prioritas: tinggi')}${opt('Urgent', t.prio, 'Prioritas: URGENT')}</select>`;
const fuNote = t => t.followup ? `<div class="fu">↳ ${esc(t.followup)}</div>` : '';
const FPRIO = { '': 'Semua', urgent: 'Urgent', tinggi: 'Urgent + tinggi', koord: 'Koordinasi disiplin', fu: 'Semua perlu follow up' };
const prioOK = t => { const f = UI.fPrio; return !f || !t ? !f : f === 'urgent' ? t.prio === 'Urgent' : f === 'tinggi' ? prioRank(t) >= 2 : f === 'koord' ? koordList(t).length > 0 : perluFU(t); };
const byPrio = (a, b) => prioRank(b) - prioRank(a);
const prioBar = () => `<div class="prio-f" role="group" aria-label="Filter prioritas">${Object.entries(FPRIO).map(([k, l]) => `<button type="button" data-act="fprio" data-v="${k}" class="${UI.fPrio === k ? 'on' : ''}">${esc(l)}</button>`).join('')}</div>`;
function normTask(t) {
  t.k = Object.assign(blankK(), t.k || {});
  KK.forEach(k => { const c = t.k[k]; c.st = c.st || 'OK'; ['issue', 'detail', 'pic'].forEach(f => { c[f] = c[f] ?? ''; }); c.target = c.target || null; });
  TXT_T.forEach(f => { t[f] = t[f] ?? ''; });
  t.minggu = +t.minggu || 0; t.mulai = t.mulai || null; t.selesai = t.selesai || null;
  t.inLPS = t.inLPS === true || t.inLPS === 'TRUE';
  if (t.bl == null || t.bl === '') delete t.bl; else t.bl = String(t.bl); // id aktivitas Timeline Schedule
  return t;
}
/* actual progress mingguan (input manual): aktual = kumulatif proyek s/d minggu w (%) */
function normProg(x) { x.w = +x.w; x.aktual = nOr(x.aktual); x.catatan = x.catatan ?? ''; return x; }
const aktualW = w => D.progress.find(p => p.w === +w)?.aktual ?? null;
const aktualTerakhir = w => D.progress.filter(p => p.aktual != null && p.w <= w).sort((a, b) => b.w - a.w)[0] || null;
const dev = v => `<b style="color:var(--${v >= 0 ? 'ok' : 'bad'})">${v >= 0 ? '+' : ''}${(Math.round(v * 100) / 100).toFixed(2)}%</b>`;
function aktualForm() {
  const cw = weekOf(tgl()), S = D.meta.scurve.filter(s => s.w <= cw + 1);
  return `<h3>Actual progress mingguan</h3>
  <p class="muted" style="font-size:13px;margin-top:-4px">Isi <b>actual kumulatif proyek (%)</b> tiap minggu dari laporan progress mingguan. Kosongkan bila belum ada.</p>
  <div class="scroll"><table class="tbl"><tr><th>Minggu</th><th class="n">Rencana kum.</th><th class="n">Actual kum. %</th><th class="n">Deviasi</th></tr>
  ${S.map(s => { const a = aktualW(s.w); return `<tr${s.w === cw ? ' style="background:var(--warn-bg)"' : ''}><td>W${s.w}<br><small class="muted">${esc(weekObj(s.w)?.label || '')}</small></td><td class="n">${s.cum.toFixed(2)}%</td>
    <td class="n"><input type="number" name="a_${s.w}" min="0" max="100" step="any" inputmode="decimal" value="${a ?? ''}" style="width:6.5em;text-align:right"></td><td class="n">${a == null ? '–' : dev(a - s.cum)}</td></tr>`; }).reverse().join('')}
  </table></div>
  <div class="dlg-act"><button value="cancel" class="btn" formnovalidate>Batal</button><button value="save" class="btn pri">Simpan</button></div>`;
}
function simpanAktual(fd) {
  let n = 0;
  for (const s of D.meta.scurve) {
    if (!fd.has('a_' + s.w)) continue;
    const v = nOr(fd.get('a_' + s.w)), r = D.progress.find(p => p.w === s.w);
    if (v == null) { if (r) { D.progress = D.progress.filter(p => p !== r); n++; } }
    else if (!r) { D.progress.push({ w: s.w, aktual: Math.min(100, v), catatan: '' }); n++; }
    else if (r.aktual !== v) { r.aktual = Math.min(100, v); n++; }
  }
  const a = D.progress.filter(p => p.aktual != null).sort((x, y) => x.w - y.w);
  const turun = a.find((p, i) => i && p.aktual < a[i - 1].aktual);
  save(); render();
  toast(n ? `Actual progress disimpan (${n} minggu)${turun ? ` · ⚠ W${turun.w} lebih kecil dari minggu sebelumnya` : ''}` : 'Tidak ada perubahan');
}
function normDaily(x) {
  ['target', 'penyebab', 'catatan'].forEach(f => { x[f] = x[f] ?? ''; });
  x.bobot = nOr(x.bobot); x.capaian = nOr(x.capaian);
  delete x.tercapai; delete x.progress; // model lama: kini dihitung dari bobot & % tercapai
  return x;
}
function normalize(s) {
  s.settings = s.settings || { tglMode: 'auto', tglFixed: null };
  s.sync = Object.assign({ url: '', key: '', user: '', rev: 0, masterVer: '', h: {}, dirty: {}, tomb: {}, lastOk: null, error: '' }, s.sync || {});
  s.tasks.forEach(normTask);
  s.progress = (s.progress || []).map(normProg);
  s.mpPlan = s.mpPlan || []; s.manpower = (s.manpower || []).map(normMp); s.dokumen = (s.dokumen || []).map(normDok);
  // v22 ke bawah mengabaikan master.timeline tapi sudah mencatat versi master → tarik ulang master sekali
  if (!Array.isArray(s.timeline)) { s.timeline = []; s.sync.masterVer = ''; }
  s.tlEdit = (s.tlEdit || []).map(normTlE);
  (s.daily || []).forEach(x => { // migrasi model lama tidak dihitung sebagai perubahan (agar baris yang dihapus tim tidak terkirim balik)
    const id = 'daily:' + x.id, sama = s.sync.h[id] === recHash(x);
    normDaily(x);
    if (sama) s.sync.h[id] = recHash(x);
  });
  return s;
}
function persist() {
  try { localStorage.setItem(KEY, JSON.stringify(D)); } catch (e) { toast('Gagal menyimpan: ' + e.message); }
}
function save() { TLc = null; stamp(); persist(); scheduleSync(); }

/* ---------- sinkron tim (Google Sheets via Apps Script, lihat gas/Code.gs) ----------
   Setiap record (task / daily log) diberi _u (waktu ubah) & _by (nama). Perubahan dideteksi dengan
   membandingkan hash isi record, jadi semua aksi cukup memanggil save(). Konflik: _u terbaru menang. */
const TBL = { tasks: 'kode', daily: 'id', progress: 'w', manpower: 'id', dokumen: 'id', tlEdit: 'id' };
const canon = v => Array.isArray(v) ? '[' + v.map(canon).join(',') + ']'
  : v && typeof v === 'object' ? '{' + Object.keys(v).sort().map(k => JSON.stringify(k) + ':' + canon(v[k])).join(',') + '}' : JSON.stringify(v ?? null);
const recHash = r => { const { _u, _by, ...rest } = r; return canon(rest); };
const syncOn = () => !!(D.sync.url && D.sync.key);
let syncing = false, syncTimer = null;
const newId = () => 'd' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6); // unik antar perangkat

function stamp() {
  const S = D.sync, now = new Date().toISOString(), seen = {};
  for (const [t, key] of Object.entries(TBL)) for (const r of D[t]) {
    const id = t + ':' + r[key], h = recHash(r);
    seen[id] = 1;
    if (S.h[id] !== h) { S.h[id] = h; r._u = now; r._by = S.user; S.dirty[id] = 1; delete S.tomb[id]; }
  }
  for (const id of Object.keys(S.h)) if (!seen[id]) { delete S.h[id]; S.tomb[id] = { _u: now, _by: S.user }; S.dirty[id] = 1; }
}
function localRec(id) {
  const [t, k] = [id.slice(0, id.indexOf(':')), id.slice(id.indexOf(':') + 1)];
  if (D.sync.tomb[id]) return { t, r: { [TBL[t]]: k, ...D.sync.tomb[id], _del: true } };
  const r = D[t].find(x => String(x[TBL[t]]) === k);
  return r ? { t, r } : null;
}
let oldGas = false; // Code.gs di Google belum versi terbaru (kolom prioritas belum ada)
function applyRemote(res, force) {
  const S = D.sync;
  let n = 0;
  if (Object.keys(TBL).some(t => !(t in res))) oldGas = true; // server belum punya tabel baru (Progress/Manpower/Dokumen)
  for (const [t, key] of Object.entries(TBL)) for (const r of res[t] || []) {
    const id = t + ':' + r[key], del = r._del === true || r._del === 'TRUE';
    delete r._del; delete r._rev;
    if (!force && S.dirty[id]) { const l = localRec(id); if (l && (l.r._u || '') > (r._u || '')) continue; } // lokal lebih baru → dikirim nanti
    const arr = D[t], i = arr.findIndex(x => String(x[key]) === String(r[key]));
    if (t === 'tasks' && !del && !('bl' in r) && i >= 0 && arr[i].bl) r.bl = arr[i].bl;
    if (t === 'tasks' && !del && !('prio' in r)) { PRIO_T.forEach(f => { r[f] = i >= 0 ? arr[i][f] : ''; }); oldGas = true; }
    if (del) { if (i >= 0) { arr.splice(i, 1); n++; } delete S.h[id]; }
    else {
      const rec = ({ tasks: normTask, daily: normDaily, progress: normProg, manpower: normMp, dokumen: normDok, tlEdit: normTlE })[t](r), h = recHash(rec);
      if (S.h[id] !== h) n++;
      if (i >= 0) arr[i] = rec; else arr.push(rec);
      S.h[id] = h;
    }
    delete S.tomb[id]; delete S.dirty[id];
  }
  S.rev = Math.max(S.rev || 0, +res.rev || 0);
  return n;
}
async function post(payload) {
  // body string → Content-Type text/plain: "simple request", tanpa preflight CORS (wajib untuk Apps Script)
  let r;
  try { r = await fetch(D.sync.url, { method: 'POST', body: JSON.stringify(payload), redirect: 'follow' }); }
  catch (e) { throw new Error(navigator.onLine ? 'server sinkron tidak terjangkau – cek URL Web App / DNS / VPN' : 'offline'); }
  if (!r.ok) throw new Error('HTTP ' + r.status);
  const j = await r.json();
  if (!j.ok) throw new Error(j.error || 'server menolak');
  return j;
}
async function syncNow(manual) {
  const S = D.sync;
  if (!syncOn()) { if (manual) toast('Atur sinkron tim dulu di ⚙'); return; }
  if (syncing) return;
  if (!navigator.onLine) { renderSync(); if (manual) toast('Offline – perubahan disimpan & dikirim saat online'); return; }
  syncing = true; renderSync();
  let changed = 0;
  try {
    const base = { key: S.key, user: S.user };
    const master = res => { if (res.master) { applyMaster(res.master); changed++; } if (res.masterVer) S.masterVer = res.masterVer; return res; };
    if (!S.rev) changed += applyRemote(master(await post({ ...base, since: 0, masterVer: S.masterVer })), true); // pertama kali: data tim menang
    const sent = {}, body = { ...base, since: S.rev, masterVer: S.masterVer, ...Object.fromEntries(Object.keys(TBL).map(t => [t, []])) };
    for (const id of Object.keys(S.dirty)) { const l = localRec(id); if (l) { body[l.t].push(l.r); sent[id] = l.r._u; } else delete S.dirty[id]; }
    changed += applyRemote(master(await post(body)), false);
    for (const id in sent) { const l = localRec(id); if (!l || l.r._u === sent[id]) delete S.dirty[id]; }
    S.lastOk = new Date().toISOString(); S.error = '';
    if (oldGas && !syncNow.warned) { syncNow.warned = true; toast('Prioritas & actual progress belum ikut tersinkron: perbarui Code.gs di Google (lihat gas/README.md)'); }
    if (manual) toast(changed ? `Sinkron selesai · ${changed} perubahan dari tim` : 'Sinkron selesai');
  } catch (e) {
    S.error = e.message;
    if (manual) toast('Sinkron gagal: ' + e.message);
  } finally {
    syncing = false; persist();
    const a = document.activeElement;
    // jangan render ulang saat orang sedang mengetik; handler change akan render dengan data terbaru
    if ((changed || !hasMaster()) && !(a && $('#view').contains(a) && /INPUT|TEXTAREA|SELECT/.test(a.tagName))) render(); else renderSync();
  }
}
function scheduleSync() { if (!syncOn()) return; clearTimeout(syncTimer); syncTimer = setTimeout(syncNow, 2500); renderSync(); }
function renderSync() {
  const el = $('#sync'); if (!el || !D) return;
  const S = D.sync, nd = Object.keys(S.dirty).length;
  const [cls, tip] = !syncOn() ? ['off', 'Sinkron tim belum diatur'] : syncing ? ['busy', 'Menyinkron…']
    : !navigator.onLine ? ['offline', `Offline · ${nd} perubahan menunggu`] : S.error ? ['err', 'Gagal: ' + S.error]
    : nd ? ['wait', `${nd} perubahan belum terkirim`] : ['ok', 'Tersinkron ' + (S.lastOk ? new Date(S.lastOk).toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' }) : '')];
  el.className = 'sync ' + cls; el.title = tip; el.setAttribute('aria-label', tip);
}
async function load() {
  try { D = JSON.parse(localStorage.getItem(KEY)); } catch (e) { D = null; }
  if (!D) D = normalize(await fetchSeed().catch(() => blank())); // hosting publik: tanpa seed.json
  else normalize(D);
  stamp(); persist();
}

/* ---------- render umum ---------- */
function toast(msg) { const el = $('#toast'); el.textContent = msg; el.classList.add('show'); clearTimeout(toast.t); toast.t = setTimeout(() => el.classList.remove('show'), 2200); }

function renderHeader() {
  const w = weekOf(tgl()), wo = weekObj(w);
  $('#hdr-sub').textContent = `Update ${fmt(tgl())} · Minggu ${w}${wo ? ' (' + wo.label + ')' : ''}`;
  $('#f-week').innerHTML = opt('', UI.week, 'Semua') + D.meta.weeks.map(x => opt(x.w, UI.week, `W${x.w} · ${x.label}`)).join('');
  const nOver = issueList().filter(x => x.sisa != null && x.sisa < 0).length;
  const b = $('#issue-badge'); b.hidden = !nOver; b.textContent = nOver;
  document.querySelectorAll('#tabs a').forEach(a => a.classList.toggle('on', a.dataset.v === UI.view));
  renderSync();
}

function vSetup() { // belum ada jadwal: aplikasi publik sebelum terhubung ke Sheet tim
  const S = D.sync;
  return `<div class="card" style="max-width:520px;margin:24px auto">
    <h2 style="margin-top:0">Hubungkan ke data tim</h2>
    <p>Data proyek disimpan di Google Sheet tim. Minta <b>URL Web App</b> dan <b>kode akses tim</b> ke admin proyek, lalu isi di pengaturan.</p>
    ${syncOn() ? `<p class="muted">${syncing ? 'Mengambil data tim…' : S.error ? 'Gagal: ' + esc(S.error) : 'Jadwal belum diunggah ke Sheet tim. Admin perlu menjalankan tools/push_master.py.'}</p>` : ''}
    <button class="btn pri" data-act="settings">Isi kode akses</button>
    ${syncOn() ? ' <button class="btn" data-act="sync-now">Coba lagi</button>' : ''}
  </div>`;
}

function render() {
  TLc = null;
  document.querySelector('.wk').hidden = !hasMaster();
  if (!hasMaster()) {
    $('#hdr-sub').textContent = 'Belum terhubung'; $('#f-week').innerHTML = ''; $('#issue-badge').hidden = true;
    $('#view').innerHTML = vSetup(); renderSync(); return;
  }
  renderHeader();
  $('#view').innerHTML = { dash: vDash, jadwal: vJadwal, komit: vKomit, ready: vReady, harian: vHarian, issue: vIssue, sd: vSD }[UI.view]();
  if (UI.focus) {
    const el = document.getElementById(UI.focus);
    if (el) { el.classList.add('hl'); el.scrollIntoView({ block: 'center' }); }
    UI.focus = null;
  }
}

/* ---------- DASHBOARD (Board 1, Board 3, PPC, kurva-S) ---------- */
function vDash() {
  const ts = planTasks();
  const by = s => ts.filter(t => stP(t) === s).length;
  const ready = ts.filter(t => stK(t) === 'READY').length;
  const pend = ts.length - ready;
  const pw = ppcWeek(UI.week), pj = ppcJanji(UI.week), pd = ppcDay(tgl()), tgt = D.meta.targetPPC;
  const cw = weekOf(tgl()), A = aktualTerakhir(UI.week === '' ? cw : +UI.week), planA = A && D.meta.scurve.find(s => s.w === A.w);
  const iss = issueList(), over = iss.filter(x => x.sisa != null && x.sisa < 0);
  const disc = [...new Set(D.tasks.map(t => t.disiplin).filter(Boolean))];
  const komitAll = D.tasks.filter(inWeek).length;
  const fu = D.tasks.filter(t => inWeek(t) && perluFU(t) && stP(t) !== 'FINISH').sort((a, b) => byPrio(a, b) || a.kode.localeCompare(b.kode));

  return `
  <div class="grid kpis">
    <div class="kpi"><b>${ts.length}</b><span>Task Work Plan (disepakati)</span></div>
    <div class="kpi ok"><b>${ready}</b><span>Task Final · READY</span></div>
    <div class="kpi warn"><b>${pend}</b><span>Pending (make ready)</span></div>
    <div class="kpi"><b>${by('ON PROGRESS')}</b><span>On Progress</span></div>
    <div class="kpi bad"><b>${by('LATE / DELAY')}</b><span>Late / Delay</span></div>
    <div class="kpi info"><b>${by('FINISH')}</b><span>Finish</span></div>
    <div class="kpi ${pj.ppc == null ? '' : pj.ppc >= tgt ? 'ok' : 'bad'}"><b>${pct(pj.ppc)}</b><span>PPC ${UI.week === '' ? 'total' : 'minggu ' + UI.week} · ${pj.ya}/${pj.n} janji tercapai (target ${pct(tgt)})</span></div>
    <div class="kpi ${pw.ppc == null ? '' : pw.ppc >= tgt ? 'ok' : 'warn'}"><b>${pct(pw.ppc)}</b><span>Task Finish ÷ Final · ${pw.fin}/${pw.final}</span></div>
    <div class="kpi ${!A ? '' : A.aktual >= (planA?.cum ?? 0) ? 'ok' : 'bad'}"><b>${A ? pc(A.aktual) : '–'}</b><span>${A ? `Actual s/d W${A.w} · rencana ${pc(planA?.cum)} · deviasi ${dev(A.aktual - (planA?.cum ?? 0))}` : 'Actual progress belum diisi'}</span></div>
    <div class="kpi ${pd.ppc == null ? '' : pd.ppc >= tgt ? 'ok' : 'bad'}"><b>${pct(pd.ppc)}</b><span>PPC hari ini · ${pd.ya}/${pd.n} janji</span></div>
    <div class="kpi ${fu.some(t => t.prio === 'Urgent') ? 'bad' : fu.length ? 'warn' : ''}"><b>${fu.length}</b><span>Perlu follow up (urgent / tinggi / koordinasi)</span></div>
    ${D.mpPlan.length ? (() => { const W = mpWeek(cw), a = mpHari(tgl()), ada = D.mpPlan.some(i => mpAkt(tgl(), i.id) != null); return `<div class="kpi ${!ada ? '' : a >= W.plan ? 'ok' : 'bad'}"><b>${ada ? a : '–'}<small class="muted" style="font-size:13px"> / ${W.plan}</small></b><span>Manpower hari ini / rencana W${cw}</span></div>`; })() : ''}
    ${TL().length ? (() => { const w = tlWeek(), k = tlKosong(w).length; return `<a class="kpi ${k ? 'warn' : 'ok'}" href="#komit" data-act="tl-goto" data-w="${w}"><b>${k}<small class="muted" style="font-size:13px"> / ${tlAkt(w).length}</small></b><span>Aktivitas timeline W${w} belum ada task komitmen →</span></a>`; })() : ''}
    ${(() => { const A = allDok(); return A.length ? `<div class="kpi ${A.some(lewatDok) ? 'bad' : 'warn'}"><b>${A.filter(x => x.status !== 'Approved').length}</b><span>Dokumen belum approve · ${A.filter(lewatDok).length} lewat target</span></div>` : ''; })()}
  </div>

  ${msCard()}
  ${fu.length ? `<div class="card">
    <h3>Task prioritas & koordinasi disiplin</h3>
    ${fu.slice(0, 8).map(t => `<div style="padding:6px 0;border-bottom:1px solid var(--line)">
      <div style="display:flex;justify-content:space-between;gap:8px;flex-wrap:wrap"><span><span class="code">${esc(t.kode)}</span> ${esc(t.uraian)} <small class="muted">${esc(t.disiplin)} · PIC ${esc(t.pic || '–')}</small></span>
        <span class="chips" style="display:flex;gap:4px;flex-wrap:wrap">${prioChips(t)}${t.disepakati === 'Ya' ? stChip(stK(t)) + stChip(stP(t)) : '<span class="chip">belum disepakati</span>'}</span></div>
      ${fuNote(t)}
      ${t.disepakati === 'Ya' ? `<a href="#ready" data-act="goto-ready" data-kode="${esc(t.kode)}" style="font-size:13px">Follow up →</a>` : `<a href="#komit" data-act="goto-komit" data-kode="${esc(t.kode)}" style="font-size:13px">Bahas di komitmen →</a>`}
    </div>`).join('')}
    ${fu.length > 8 ? `<p><a href="#ready" data-act="fprio" data-v="fu">Lihat semua (${fu.length}) →</a></p>` : ''}
  </div>` : ''}

  <div class="card">
    <div class="vh"><h3 style="margin:0">Kurva-S rencana vs actual</h3><button class="btn sm" data-act="aktual">Input actual mingguan</button></div>
    ${sCurve()}
  </div>

  <div class="two">
    <div class="card">
      <h3>Rekap per disiplin</h3>
      <div class="scroll"><table class="tbl">
        <tr><th></th>${disc.map(d => `<th class="n">${esc(d)}</th>`).join('')}<th class="n">Total</th></tr>
        ${[['Task komitmen', t => inWeek(t)], ['Disepakati', t => t.disepakati === 'Ya' && inWeek(t)],
           ['Ready', t => t.disepakati === 'Ya' && inWeek(t) && stK(t) === 'READY'],
           ['Pending', t => t.disepakati === 'Ya' && inWeek(t) && stK(t) === 'PENDING'],
           ['On Progress', t => t.disepakati === 'Ya' && inWeek(t) && stP(t) === 'ON PROGRESS'],
           ['Late / Delay', t => t.disepakati === 'Ya' && inWeek(t) && stP(t) === 'LATE / DELAY'],
           ['Finish', t => t.disepakati === 'Ya' && inWeek(t) && stP(t) === 'FINISH']]
          .map(([n, f]) => `<tr><td>${n}</td>${disc.map(d => `<td class="n">${D.tasks.filter(t => t.disiplin === d && f(t)).length}</td>`).join('')}<td class="n"><b>${D.tasks.filter(f).length}</b></td></tr>`).join('')}
        <tr><td>% Komitmen</td>${disc.map(d => { const a = D.tasks.filter(t => t.disiplin === d && inWeek(t)); return `<td class="n">${pct(a.length ? a.filter(t => t.disepakati === 'Ya').length / a.length : null)}</td>`; }).join('')}<td class="n"><b>${pct(komitAll ? ts.length / komitAll : null)}</b></td></tr>
      </table></div>
    </div>
    <div class="card">
      <h3>Kendala perlu follow up ${over.length ? `<span class="chip bad">${over.length} lewat target</span>` : ''}</h3>
      ${iss.length ? iss.slice(0, 6).map(x => `
        <div style="padding:6px 0;border-bottom:1px solid var(--line)">
          <div class="hd" style="display:flex;justify-content:space-between;gap:8px"><span><span class="code">${esc(x.t.kode)}</span> ${esc(x.c.issue || KAT[x.k])}</span>${stChip(x.st)}</div>
          <small class="muted">${esc(KAT[x.k])} · PIC ${esc(x.c.pic || '–')} · target ${fmt(x.c.target)}</small>
        </div>`).join('') + `<p><a href="#issue">Lihat semua (${iss.length}) →</a></p>` : '<p class="muted">Tidak ada kendala pending. 👍</p>'}
    </div>
  </div>`;
}

function sCurve() {
  const S = D.meta.scurve, W = 640, H = 230, L = 34, R = 10, T = 12, B = 28;
  const n = S.length, x = i => L + (W - L - R) * i / (n - 1), y = v => T + (H - T - B) * (1 - v / 100);
  const maxPlan = Math.max(...S.map(s => s.plan), 1);
  const bw = (W - L - R) / n * 0.6;
  const cw = weekOf(tgl()), ci = S.findIndex(s => s.w === cw);
  const now = ci >= 0 ? S[ci] : null;
  let g = '';
  for (const v of [0, 25, 50, 75, 100]) g += `<line x1="${L}" x2="${W - R}" y1="${y(v)}" y2="${y(v)}" stroke="var(--grid)" stroke-width="1"/><text x="${L - 6}" y="${y(v) + 4}" text-anchor="end">${v}%</text>`;
  S.forEach((s, i) => {
    const h = (H - T - B) * 0.45 * s.plan / maxPlan;
    g += `<rect x="${x(i) - bw / 2}" y="${H - B - h}" width="${bw}" height="${h}" fill="var(--brand-2)" opacity=".55" stroke="none"><title>W${s.w}: rencana ${s.plan.toFixed(2)}%</title></rect>`;
    if (i % 3 === 0) g += `<text x="${x(i)}" y="${H - 10}" text-anchor="middle">W${s.w}</text>`;
  });
  g += `<polyline points="${S.map((s, i) => x(i) + ',' + y(s.cum)).join(' ')}" stroke="var(--plan)" stroke-width="2.5"/>`;
  const act = D.progress.filter(p => p.aktual != null).sort((a, b) => a.w - b.w).map(p => ({ ...p, i: S.findIndex(s => s.w === p.w) })).filter(p => p.i >= 0);
  if (act.length) {
    g += `<polyline points="${act.map(p => x(p.i) + ',' + y(p.aktual)).join(' ')}" stroke="var(--ok)" stroke-width="2.5"/>`;
    act.forEach(p => { g += `<circle cx="${x(p.i)}" cy="${y(p.aktual)}" r="3.5" fill="var(--ok)" stroke="var(--surface)" stroke-width="1.5"><title>W${p.w}: actual ${pc(p.aktual)} · rencana ${pc(S[p.i].cum)}</title></circle>`; });
  }
  S.forEach((s, i) => { g += `<circle cx="${x(i)}" cy="${y(s.cum)}" r="6" fill="transparent" stroke="none"><title>W${s.w}: kumulatif ${s.cum.toFixed(2)}%</title></circle>`; });
  if (now) g += `<line x1="${x(ci)}" x2="${x(ci)}" y1="${T}" y2="${H - B}" stroke="var(--bad)" stroke-dasharray="4 3" stroke-width="1.5"/>
      <circle cx="${x(ci)}" cy="${y(now.cum)}" r="4.5" fill="var(--bad)" stroke="var(--surface)" stroke-width="2"/>
      <text x="${Math.min(x(ci) + 6, W - 120)}" y="${Math.max(y(now.cum) - 8, T + 10)}" style="fill:var(--text);font-weight:600">Rencana s/d W${now.w}: ${now.cum.toFixed(2)}%</text>`;
  return `<svg class="chart" viewBox="0 0 ${W} ${H}" role="img" aria-label="Kurva S rencana">${g}</svg>
    <div class="legend"><span><i style="background:var(--plan)"></i>Kumulatif rencana</span><span><i style="background:var(--ok)"></i>Kumulatif actual</span><span><i style="background:var(--brand-2);opacity:.55"></i>Rencana per minggu</span><span><i style="background:var(--bad)"></i>Minggu berjalan</span></div>`;
}

/* ---------- JADWAL (Daftar Pekerjaan + Lookahead) ---------- */
/* pekerjaan aktif di rentang minggu (lookahead) – dipakai tampilan Jadwal & PDF Jadwal */
function jadwalList(from, look, dis, q = '') {
  const wF = weekObj(from), wT = weekObj(Math.min(+from + (+look || 0) - 1, D.meta.weeks.at(-1).w));
  q = q.toLowerCase();
  const list = D.pekerjaan.filter(p =>
    (!dis || p.disiplin === dis) &&
    (!q || (p.nama + ' ' + p.area + ' ' + p.bill).toLowerCase().includes(q)) &&
    (look === 'all' || (p.mulai && wF && wT && p.mulai <= wT.end && p.selesai >= wF.start)));
  return { wF, wT, list };
}
function vJadwal() {
  const cw = weekOf(tgl());
  const from = UI.jFrom ?? cw;
  const look = UI.jLook;
  const { wF, wT, list } = jadwalList(from, look, UI.jDis, UI.jSearch);
  return `
  <div class="vh"><h2>Jadwal & Lookahead</h2><button class="btn sm" data-act="pdf-jadwal">PDF jadwal</button></div>
  <div class="bar">
    <input class="grow" type="search" placeholder="Cari pekerjaan / area…" value="${esc(UI.jSearch)}" data-ui="jSearch">
    <select data-ui="jDis">${opt('', UI.jDis, 'Semua disiplin')}${D.disiplin.map(d => opt(d, UI.jDis)).join('')}</select>
    <select data-ui="jLook">${[['1', '1 minggu'], ['2', '2 minggu'], ['3', '3 minggu'], ['4', '4 minggu'], ['6', '6 minggu'], ['all', 'Semua jadwal']].map(([v, l]) => opt(v, look, 'Lookahead ' + l)).join('')}</select>
    ${look === 'all' ? '' : `<select data-ui="jFrom">${D.meta.weeks.map(x => opt(x.w, from, `mulai W${x.w}${x.w === cw ? ' (ini)' : ''}`)).join('')}</select>`}
  </div>
  <p class="muted" style="margin-top:-4px">${list.length} pekerjaan${look === 'all' ? '' : ` aktif ${fmt(wF?.start)} – ${fmt(wT?.end)}`}</p>
  ${list.map(p => {
    const ts = D.tasks.filter(t => t.pekerjaan === p.nama);
    const ok = ts.filter(t => t.disepakati === 'Ya'), fin = ok.filter(t => stP(t) === 'FINISH').length;
    return `<div class="item">
      <div class="hd"><div><span class="code">${esc(p.kode)}</span> <span class="ttl">${esc(p.nama.replace(/^\S+\s/, ''))}</span>
        <div class="sub">${esc(p.bill)}${p.area ? ' · ' + esc(p.area) : ''}</div></div>
        <div class="chips"><span class="chip">${esc(p.disiplin)}</span>${stChip(stJadwal(p))}</div></div>
      <div class="meta"><span>${fmt(p.mulai)} → ${fmt(p.selesai)}</span><span>W${weekOf(p.mulai) ?? '–'}–W${weekOf(p.selesai) ?? '–'}</span><span>Bobot ${p.bobot.toFixed(2)}%</span>
        <span>Task ${ts.length} · sepakat ${ok.length} · finish ${fin}${ok.length ? ' (' + pct(fin / ok.length) + ')' : ''}</span>${(() => { const A = TL().filter(a => a.pek === p.nama); return A.length ? `<span>⏱ timeline ${A.length} aktivitas · ${fmt(A.reduce((m, a) => a.mulai < m ? a.mulai : m, A[0].mulai))} → ${fmt(A.reduce((m, a) => a.selesai > m ? a.selesai : m, A[0].selesai))}</span>` : ''; })()}</div>
      <div style="margin-top:8px"><button class="btn sm" data-act="task-new" data-pek="${esc(p.nama)}">+ Task komitmen</button></div>
    </div>`;
  }).join('') || '<div class="empty">Tidak ada pekerjaan pada filter ini.</div>'}`;
}

/* ---------- TIMELINE SCHEDULE: bank aktivitas (PDF MS Project → tools/export_timeline.py → master.timeline) ----------
   Task komitmen diambil dari aktivitas yang aktif di minggu itu: 1 task per aktivitas per minggu (tanggal dipotong ke minggu),
   task menyimpan bl = id aktivitas agar celah komitmen & selisih terhadap timeline bisa dihitung. */
function normTlE(x) { x.id = String(x.id); x.catatan = x.catatan ?? ''; return x; }
let TLc = null; // timeline efektif, dihitung ulang tiap render / simpan
const TL = () => TLc || (TLc = (() => {
  const E = Object.fromEntries((D.tlEdit || []).map(e => [e.id, e]));
  return (D.timeline || []).map(a => { const e = E[String(a.id)]; return e ? { ...a, mulai: e.mulai || a.mulai, selesai: e.selesai || a.selesai, asli: { mulai: a.mulai, selesai: a.selesai }, ubah: e } : a; });
})());
const tlById = id => TL().find(a => String(a.id) === String(id));
const tlAkt = w => { const o = weekObj(w); return o ? TL().filter(a => !a.ms && a.mulai <= o.end && a.selesai >= o.start) : []; };
const tlTasks = (a, w) => D.tasks.filter(t => t.bl === String(a.id) && (w == null || +t.minggu === +w));
const tlKosong = w => tlAkt(w).filter(a => !tlTasks(a, w).length);
const tlSlip = t => { const a = t.bl && tlById(t.bl); return a && t.selesai ? days(a.selesai, t.selesai) : null; }; // + = lebih lambat dari timeline
const tlWeek = () => UI.week !== '' ? +UI.week : Math.min(weekOf(tgl()) + 1, D.meta.weeks.at(-1).w); // default: minggu depan (rapat komitmen)
/* status aktivitas dibanding HARI INI (bukan minggu yang dipilih): belum mulai / berjalan / lewat */
const tlSt = a => { const d = tgl(), n = days(d, a.mulai);
  return a.mulai > d ? ` <span class="chip">mulai ${n === 1 ? 'besok' : n + ' hari lagi'} · ${fmt(a.mulai)}</span>`
    : a.selesai < d ? ` <span class="chip bad">lewat target ${fmt(a.selesai)}</span>`
    : ` <span class="chip warn">sedang berjalan</span>`; };
const tlInfo = a => `${esc([a.lok, a.area].filter(Boolean).join(' · '))} · ${esc(a.dis)} · timeline ${fmt(a.mulai)} – ${fmt(a.selesai)}${a.ubah ? ` <span class="chip info" title="${esc(`Diubah ${a.ubah._by || ''}${a.ubah.catatan ? ': ' + a.ubah.catatan : ''}`)}">diubah · PDF ${fmt(a.asli.mulai)} – ${fmt(a.asli.selesai)}</span>` : ''}`;
function tlGapCard(w) {
  const o = weekObj(w), akt = tlAkt(w), kosong = tlKosong(w), MAX = 25;
  const ms = TL().filter(a => a.ms && a.mulai >= o.start && a.mulai <= o.end);
  return `<details class="card tl-gap"${kosong.length && kosong.length <= 15 ? ' open' : ''}>
    <summary>Timeline W${w}: ${akt.length} aktivitas aktif · <span class="chip ${kosong.length ? 'warn' : 'ok'}">${kosong.length} belum ada task</span>
      ${kosong.filter(a => a.mulai >= o.start).length ? `<span class="chip info">${kosong.filter(a => a.mulai >= o.start).length} mulai minggu ini</span>` : ''}
      ${ms.map(m => `<span class="chip bad">◆ ${esc(m.nama)}</span>`).join('')}</summary>
    <div class="bar" style="margin:8px 0"><button class="btn sm pri" data-act="tl-pick" data-w="${w}">Ambil dari timeline…</button>
      <span class="muted" style="font-size:13px">Pilih beberapa sekaligus; PIC & grid diisi setelahnya.</span></div>
    ${kosong.slice(0, MAX).map(a => `<div class="tl-r"><span><b>${esc(a.nama)}</b>${tlSt(a)}<small>${tlInfo(a)}</small></span>
      <span class="tl-b"><button class="btn sm" data-act="tl-ubah" data-id="${esc(a.id)}" title="Ubah tanggal aktivitas timeline" aria-label="Ubah tanggal">✎</button><button class="btn sm" data-act="tl-add" data-id="${esc(a.id)}" data-w="${w}">+ Task</button></span></div>`).join('')}
    ${kosong.length > MAX ? `<p class="muted" style="font-size:13px">… ${kosong.length - MAX} lagi – pakai <b>Ambil dari timeline</b> (bisa disaring per area / disiplin).</p>` : ''}
  </details>`;
}
function tlPick(w) {
  w = +w;
  const cw = weekOf(tgl()), o = weekObj(w);
  const akt = tlAkt(w).sort((a, b) => a.area.localeCompare(b.area) || a.mulai.localeCompare(b.mulai));
  const areas = [...new Set(akt.map(a => a.area))].sort();
  openDlg(`<h3>Ambil task dari Timeline Schedule</h3>
  <div class="frow">
    <label class="f">Minggu<select id="tl-w">${D.meta.weeks.map(x => opt(x.w, w, `W${x.w} · ${x.label}${x.w === cw ? ' (ini)' : ''} · ${tlKosong(x.w).length} belum`)).join('')}</select></label>
    <label class="f">Area<select id="tl-area">${opt('', '', 'Semua area')}${areas.map(a => opt(a, '')).join('')}</select></label>
  </div>
  <div class="frow">
    <label class="f">Disiplin<select id="tl-dis">${opt('', '', 'Semua disiplin')}${D.disiplin.map(d => opt(d, '')).join('')}</select></label>
    <label class="f">Cari<input type="search" id="tl-cari" placeholder="aktivitas / lokasi / pekerjaan"></label>
  </div>
  <div class="bar" style="margin:2px 0 0"><button type="button" class="btn sm" data-act="tl-manual">+ Task manual (di luar timeline)</button><button type="button" class="btn sm" data-act="tl-all">Centang yang tampil & belum ada task</button><button type="button" class="btn sm" data-act="tl-none">Kosongkan</button><span class="muted" id="tl-n" style="font-size:13px"></span></div>
  <div class="pick" id="tl-pick">${akt.map(a => { const ada = tlTasks(a, w); return `<label class="pick-i" data-area="${esc(a.area)}" data-dis="${esc(a.dis)}" data-cari="${esc((a.nama + ' ' + (a.lok || '') + ' ' + (a.pek || '')).toLowerCase())}">
    <input type="checkbox" name="bl" value="${esc(a.id)}"${ada.length ? ' disabled' : ''}>
    <span><b>${esc(a.nama)}</b>${ada.length ? ` <span class="chip ok">sudah: ${ada.map(t => esc(t.kode)).join(', ')}</span>` : ''}${tlSt(a)}${a.selesai <= o.end ? ' <span class="chip warn">harus selesai</span>' : ''}
    <small>${tlInfo(a)}<br>${a.pek ? esc(a.pek) : '<i>pekerjaan BOQ belum dipetakan</i>'}</small></span></label>`; }).join('') || '<p class="muted" style="padding:10px">Tidak ada aktivitas timeline di minggu ini.</p>'}</div>
  <p class="muted" style="font-size:13px;margin:6px 0 0">Task dibuat untuk <b>W${w}</b> (${fmt(o?.start)} – ${fmt(o?.end)}); tanggal dipotong ke minggu ini, aktivitas panjang diambil lagi minggu berikutnya. Status: belum dibahas.</p>
  <div class="dlg-act"><button value="cancel" class="btn" formnovalidate>Batal</button><button value="save" class="btn pri">Buat task</button></div>`,
  fd => { const n = tlBuat(fd.getAll('bl'), w); if (n) { UI.week = String(w); location.hash === '#komit' ? render() : (location.hash = 'komit'); } });
  tlFilter();
}
function tlFilter() {
  const q = ($('#tl-cari')?.value || '').toLowerCase().trim(), ar = $('#tl-area')?.value, di = $('#tl-dis')?.value;
  let n = 0;
  document.querySelectorAll('#tl-pick .pick-i').forEach(l => { l.hidden = (!!q && !l.dataset.cari.includes(q)) || (!!ar && l.dataset.area !== ar) || (!!di && l.dataset.dis !== di); if (!l.hidden) n++; });
  const c = document.querySelectorAll('#tl-pick input:checked').length;
  if ($('#tl-n')) $('#tl-n').textContent = `${n} tampil · ${c} dicentang`;
}
function tlUbah(id) {
  const a = tlById(id), o = (D.timeline || []).find(x => String(x.id) === String(id)), e = D.tlEdit.find(x => x.id === String(id));
  if (!a || !o) return;
  const ts = tlTasks(a);
  openDlg(`<h3>Ubah tanggal aktivitas timeline</h3>
  <p style="margin:0 0 4px"><b>${esc(a.nama)}</b><br><small class="muted">${esc([a.lok, a.area, a.dis].filter(Boolean).join(' · '))}</small></p>
  <p class="muted" style="font-size:13px">Tanggal di PDF: <b>${fmt(o.mulai)} – ${fmt(o.selesai)}</b>${e ? ` · terakhir diubah ${esc(e._by || '')}` : ''}. Perubahan berlaku untuk seluruh tim (daftar celah, lookahead, selisih task).</p>
  <div class="frow">
    <label class="f">Mulai<input type="date" name="mulai" value="${esc(a.mulai)}" required></label>
    <label class="f">Selesai<input type="date" name="selesai" value="${esc(a.selesai)}" required></label>
  </div>
  <label class="f">Alasan / catatan<input type="text" name="catatan" value="${esc(e?.catatan || '')}" placeholder="mis. ikut jadwal aktual, revisi MS Project"></label>
  ${ts.length ? `<p class="muted" style="font-size:13px">Task yang sudah dibuat (${ts.map(t => esc(t.kode)).join(', ')}) tidak ikut berubah; ubah tanggalnya di task masing-masing bila perlu.</p>` : ''}
  <div class="dlg-act">${e ? `<button type="button" class="btn danger" data-act="tl-reset" data-id="${esc(a.id)}">Kembalikan ke PDF</button>` : ''}<span class="sp"></span>
    <button value="cancel" class="btn" formnovalidate>Batal</button><button value="save" class="btn pri">Simpan</button></div>`,
  fd => {
    const m = fd.get('mulai'), sl = fd.get('selesai');
    if (!m || !sl || sl < m) { toast('Tanggal selesai tidak boleh sebelum mulai'); return; }
    D.tlEdit = D.tlEdit.filter(x => x.id !== String(a.id));
    if (m !== o.mulai || sl !== o.selesai) D.tlEdit.push({ id: String(a.id), mulai: m, selesai: sl, catatan: String(fd.get('catatan') || '').trim() });
    save(); render(); toast(m === o.mulai && sl === o.selesai ? 'Tanggal kembali sesuai PDF' : `Tanggal timeline diubah: ${fmt(m)} – ${fmt(sl)}`);
  });
}
function tlBuat(ids, w) {
  const o = weekObj(w);
  if (!ids.length || !o) { toast('Belum ada aktivitas yang dicentang'); return 0; }
  const baru = [];
  ids.forEach(id => {
    const a = tlById(id);
    if (!a || tlTasks(a, w).length) return;
    const t = normTask({ kode: nextKode(), minggu: +w, disiplin: a.dis, pekerjaan: a.pek || '', area: D.areas.includes(a.area) ? a.area : '', uraian: a.nama,
      lokasi: a.lok || '', pic: '', mulai: a.mulai > o.start ? a.mulai : o.start, selesai: a.selesai < o.end ? a.selesai : o.end,
      disepakati: '', alasan: '', inLPS: false, k: blankK(), catatanKendala: '', penyebabTerlambat: '', prio: '', koord: '', followup: '', bl: String(a.id) });
    D.tasks.push(t); baru.push(t);
  });
  save(); toast(`${baru.length} task dibuat dari timeline untuk W${w}`);
  return baru;
}
const tlChip = t => { const a = t.bl && tlById(t.bl); if (!a) return ''; const s = tlSlip(t);
  return `<a href="#" data-act="tl-ubah" data-id="${esc(a.id)}" title="Ubah tanggal aktivitas timeline">⏱ timeline ${fmt(a.mulai)} – ${fmt(a.selesai)}${a.ubah ? ' (diubah)' : ''}</a>${s > 0 ? `<span class="chip bad">+${s} hr dari timeline</span>` : ''}`; };
function msCard() {
  const M = TL().filter(a => a.ms && days(tgl(), a.mulai) >= -7).sort((a, b) => a.mulai.localeCompare(b.mulai)); // yang lewat > 7 hari disembunyikan
  if (!M.length) return '';
  const now = tgl();
  return `<div class="card"><h3>Milestone (Timeline Schedule)</h3><div class="ms-l">${M.map(m => { const d = days(now, m.mulai);
    return `<div class="ms ${d < 0 ? 'lewat' : d <= 30 ? 'dekat' : ''}"><b>${esc(m.nama)}</b><span>${fmt(m.mulai)}${weekOf(m.mulai) != null ? ' · W' + weekOf(m.mulai) : ''}</span><em>${d < 0 ? 'lewat ' + -d + ' hari' : d === 0 ? 'hari ini' : d + ' hari lagi'}</em></div>`; }).join('')}</div></div>`;
}

/* ---------- KOMITMEN ---------- */
function vKomit() {
  const ts = D.tasks.filter(inWeek).filter(prioOK).sort((a, b) => byPrio(a, b) || a.minggu - b.minggu || a.kode.localeCompare(b.kode));
  let gap = TL().length ? tlGapCard(tlWeek()) : '';
  if (!TL().length && UI.week !== '') {
    const w = weekObj(UI.week);
    const aktif = D.pekerjaan.filter(p => p.mulai && w && p.mulai <= w.end && p.selesai >= w.start);
    const kosong = aktif.filter(p => !D.tasks.some(t => t.pekerjaan === p.nama && +t.minggu === +UI.week));
    gap = `<details class="card"><summary>Pekerjaan MASTER aktif minggu ${UI.week}: ${aktif.length} · <span class="chip warn">${kosong.length} belum ada task</span></summary>
      ${kosong.map(p => `<div style="display:flex;justify-content:space-between;gap:8px;padding:5px 0;border-bottom:1px solid var(--line)"><span><span class="code">${esc(p.kode)}</span> ${esc(p.nama.replace(/^\S+\s/, ''))} <small class="muted">${esc(p.disiplin)}</small></span><button class="btn sm" data-act="task-new" data-pek="${esc(p.nama)}">+ Task</button></div>`).join('')}
    </details>`;
  }
  const n = ts.length, ya = ts.filter(t => t.disepakati === 'Ya').length;
  return `
  ${vh(`Komitmen mingguan ${UI.week === '' ? '' : '· W' + UI.week}`, 'komit')}
  ${prioBar()}
  <p class="muted" style="margin-top:-6px">${n} task · ${ya} disepakati${n ? ' (' + pct(ya / n) + ')' : ''}. Pilih <b>Ya</b> saat rapat mingguan untuk task yang sanggup dikerjakan.</p>
  ${gap}
  ${ts.map(t => `
    <div class="item" id="k-${esc(t.kode)}">
      <div class="hd"><div><span class="code">${esc(t.kode)}</span> <span class="ttl">${esc(t.uraian || '(tanpa uraian)')}</span>
        <div class="sub">${esc(t.pekerjaan)}</div></div>
        <div class="chips">${prioChips(t)}<span class="chip">W${t.minggu}</span><span class="chip">${esc(t.disiplin)}</span></div></div>
      <div class="meta"><span>${esc(t.area)}${t.lokasi ? ' · ' + esc(t.lokasi) : ''}</span><span>PIC ${esc(t.pic || '–')}</span><span>${fmt(t.mulai)} → ${fmt(t.selesai)}${t.mulai && t.selesai ? ' (' + (days(t.mulai, t.selesai) + 1) + ' hr)' : ''}</span>${tlChip(t)}</div>
      <div class="bar" style="margin:10px 0 0">
        <span class="seg">
          <button type="button" data-act="sepakat" data-kode="${esc(t.kode)}" data-v="Ya" class="${t.disepakati === 'Ya' ? 'on-ok' : ''}">Disepakati</button>
          <button type="button" data-act="sepakat" data-kode="${esc(t.kode)}" data-v="Tidak" class="${t.disepakati === 'Tidak' ? 'on-bad' : ''}">Ditunda</button>
        </span>
        ${t.disepakati === 'Ya' ? stChip(stK(t)) : ''}
        <span style="flex:1"></span>
        ${prioSel(t)}
        <button class="btn sm" data-act="task-edit" data-kode="${esc(t.kode)}">Ubah</button>
      </div>
      ${fuNote(t)}
      ${t.disepakati === 'Tidak' ? `<input type="text" style="margin-top:8px" placeholder="Alasan ditunda / catatan rapat" value="${esc(t.alasan)}" data-f="task|${esc(t.kode)}|alasan" class="${t.alasan ? '' : 'need'}">` : ''}
    </div>`).join('') || '<div class="empty">Belum ada task untuk filter ini. Tekan <b>+ Task</b>.</div>'}
  <button class="btn pri fab" data-act="${TL().length ? 'tl-pick' : 'task-new'}" data-w="${TL().length ? tlWeek() : ''}">+ Task</button>`;
}

function taskForm(t, isNew) {
  const pek = D.pekerjaan.map(p => p.nama);
  return `<h3>${isNew ? 'Task baru' : 'Ubah task ' + esc(t.kode)}</h3>
  <div class="frow">
    <label class="f">Kode<input type="text" name="kode" value="${esc(t.kode)}" required ${isNew ? '' : 'readonly'}></label>
    <label class="f">Minggu ke<select name="minggu">${D.meta.weeks.map(x => opt(x.w, t.minggu, `W${x.w} · ${x.label}`)).join('')}</select></label>
  </div>
  <label class="f">Pekerjaan (Master Schedule)<input type="text" name="pekerjaan" list="dl-pek" value="${esc(t.pekerjaan)}" required>
    <datalist id="dl-pek">${pek.map(p => `<option value="${esc(p)}">`).join('')}</datalist></label>
  <div class="frow">
    <label class="f">Disiplin<select name="disiplin">${D.disiplin.map(d => opt(d, t.disiplin)).join('')}</select></label>
    <label class="f">Area gedung<select name="area">${opt('', t.area, '–')}${D.areas.map(a => opt(a, t.area)).join('')}</select></label>
  </div>
  <label class="f">Uraian pekerjaan<textarea name="uraian" required>${esc(t.uraian)}</textarea></label>
  <div class="frow">
    <label class="f">Lokasi / grid<input type="text" name="lokasi" value="${esc(t.lokasi)}"></label>
    <label class="f">PIC<input type="text" name="pic" value="${esc(t.pic)}"></label>
  </div>
  <div class="frow">
    <label class="f">Rencana mulai<input type="date" name="mulai" value="${esc(t.mulai || '')}"></label>
    <label class="f">Rencana selesai<input type="date" name="selesai" value="${esc(t.selesai || '')}"></label>
  </div>
  <div class="frow">
    <label class="f">Disepakati<select name="disepakati">${opt('', t.disepakati, 'Belum dibahas')}${opt('Ya', t.disepakati)}${opt('Tidak', t.disepakati)}</select></label>
    <label class="f">Prioritas<select name="prio">${opt('', t.prio, 'Normal')}${opt('Tinggi', t.prio)}${opt('Urgent', t.prio, 'URGENT')}</select></label>
  </div>
  <div class="f">Koordinasi dengan disiplin lain<div class="ck">${D.disiplin.filter(d => d !== t.disiplin).map(d => `<label><input type="checkbox" name="koord" value="${esc(d)}"${koordList(t).includes(d) ? ' checked' : ''}> ${esc(d)}</label>`).join('')}</div></div>
  <label class="f">Catatan follow-up / koordinasi<textarea name="followup" placeholder="mis. tunggu shop drawing MEP, clash dengan ducting, perlu izin owner…">${esc(t.followup)}</textarea></label>
  <div class="dlg-act">
    ${isNew ? '' : `<button type="button" class="btn danger" data-act="task-del" data-kode="${esc(t.kode)}">Hapus</button>`}
    <span class="sp"></span>
    <button value="cancel" class="btn" formnovalidate>Batal</button>
    <button value="save" class="btn pri">Simpan</button>
  </div>`;
}

/* ---------- MAKE READY (LPS Input + Board 2) ---------- */
function vReady() {
  const ts = planTasks();
  const cnt = k => ts.filter(t => t.k[k].st === 'Pending').length;
  const issues = k => D.issues[k] || [];
  const d = tgl();
  return `
  ${vh('Make Ready · cek kendala', 'ready')}
  ${prioBar()}
  <div class="grid kpis">
    ${KK.map(k => `<div class="kpi ${cnt(k) ? 'warn' : ''}"><b>${cnt(k)}</b><span>Pending ${KAT[k]}</span></div>`).join('')}
    <div class="kpi ok"><b>${ts.filter(t => stK(t) === 'READY').length}</b><span>Task Ready to Work</span></div>
  </div>
  ${ts.filter(prioOK).sort((a, b) => byPrio(a, b) || (stK(a) === 'PENDING' ? 0 : 1) - (stK(b) === 'PENDING' ? 0 : 1) || a.kode.localeCompare(b.kode)).map(t => {
    const p = prog(t.kode), sp = stP(t);
    return `<div class="item" id="r-${esc(t.kode)}">
      <div class="hd"><div><span class="code">${esc(t.kode)}</span> <span class="ttl">${esc(t.uraian)}</span>
        <div class="sub">${esc(t.pekerjaan)}</div></div>
        <div class="chips">${prioChips(t)}${stChip(stK(t))}${stChip(sp)}</div></div>
      <div class="meta"><span>W${t.minggu} · ${esc(t.disiplin)}</span><span>${esc(t.area)}${t.lokasi ? ' · ' + esc(t.lokasi) : ''}</span><span>PIC ${esc(t.pic || '–')}</span><span>${fmt(t.mulai)} → ${fmt(t.selesai)}</span>${prioSel(t)}</div>
      ${fuNote(t)}
      ${dokLine(t)}
      <div style="display:flex;align-items:center;gap:8px;margin:8px 0"><div class="bar-h" style="flex:1"><i style="width:${Math.min(p || 0, 100)}%"></i></div><small class="muted">Progress ${p == null ? '0' : p}% (dari Daily Log)</small>${stK(t) === 'READY' && sp !== 'FINISH' ? `<button class="btn sm" data-act="daily-new" data-kode="${esc(t.kode)}">+ Janji harian</button>` : ''}</div>
      ${KK.map(k => {
        const c = t.k[k], pend = c.st === 'Pending', f = `task|${esc(t.kode)}|k.${k}.`;
        const over = pend && c.target && c.target < d;
        return `<div class="kr">
          <div class="kh"><span>${KAT[k]}</span>
            <span class="seg">
              <button type="button" data-act="kst" data-kode="${esc(t.kode)}" data-k="${k}" data-v="OK" class="${pend ? '' : 'on-ok'}">OK</button>
              <button type="button" data-act="kst" data-kode="${esc(t.kode)}" data-k="${k}" data-v="Pending" class="${pend ? 'on-warn' : ''}">Pending</button>
            </span></div>
          ${pend ? `<div class="kd">
            <select class="full ${c.issue ? '' : 'need'}" data-f="${f}issue">${opt('', c.issue, '– pilih jenis issue –')}${[...new Set([...issues(k), c.issue].filter(Boolean))].map(i => opt(i, c.issue)).join('')}</select>
            <input class="full" type="text" placeholder="Detail issue (no. gambar, material, PO, alat…)" value="${esc(c.detail)}" data-f="${f}detail">
            <input type="text" placeholder="PIC follow up" value="${esc(c.pic)}" data-f="${f}pic" class="${c.pic ? '' : 'need'}">
            <input type="date" value="${esc(c.target || '')}" data-f="${f}target" class="${!c.target ? 'need' : over ? 'over' : ''}" aria-label="Target selesai kendala">
          </div>` : ''}
        </div>`;
      }).join('')}
      ${sp === 'LATE / DELAY' ? `<label class="f" style="margin-top:8px">Penyebab terlambat<select data-f="task|${esc(t.kode)}|penyebabTerlambat" class="${t.penyebabTerlambat ? '' : 'need'}">${opt('', t.penyebabTerlambat, '– pilih –')}${D.penyebab.map(x => opt(x, t.penyebabTerlambat)).join('')}</select></label>` : ''}
      <input type="text" style="margin-top:8px" placeholder="Catatan kendala" value="${esc(t.catatanKendala)}" data-f="task|${esc(t.kode)}|catatanKendala">
    </div>`;
  }).join('') || '<div class="empty">Belum ada task yang disepakati. Buka tab <a href="#komit">Komitmen</a>.</div>'}`;
}

/* kartu satu janji di dalam header task (Harian) */
// = kolom Cek di Daily Log
const cek = (x, t) => !t ? '⚠ Kode tidak ada' : stK(t) !== 'READY' ? '⚠ Task belum READY' : !x.target || x.bobot == null ? '⚠ Isi target & bobot'
  : x.bobot > sisaTask(x) + 0.01 ? '⚠ Bobot > sisa task' : x.capaian == null ? 'Menunggu laporan sore'
  : tercapai(x) === 'Tidak' && !x.penyebab ? '⚠ Isi penyebab' : 'OK';
function janjiCard(x, t, no, n) {
  const c = cek(x, t), f = `daily|${x.id}|`, kem = progKemarin(x), lain = r6(progSebelum(x) - kem);
  return `<div class="janji">
    <div class="hd"><b>Janji ${no}${n > 1 ? '/' + n : ''}</b>
      <div class="chips">${stChip(tercapai(x))}<span class="chip ${c === 'OK' ? 'ok' : c.startsWith('⚠') ? 'warn' : ''}">${esc(c)}</span></div></div>
    <label class="f">Target hari ini (pagi)<input type="text" value="${esc(x.target)}" data-f="${f}target" class="${x.target ? '' : 'need'}"></label>
      <div class="frow">
        <label class="f">Bobot target (% dari task, pagi)<input type="number" min="0" max="100" step="any" inputmode="decimal" value="${x.bobot ?? ''}" data-f="${f}bobot" class="${x.bobot == null ? 'need' : x.bobot > sisaTask(x) + 0.01 ? 'over' : ''}"></label>
        <label class="f">Actual progress hari ini (% dari task, sore)<input type="number" min="0" max="100" step="any" inputmode="decimal" value="${aktual(x) ?? ''}" data-f="${f}aktual" ${x.bobot == null ? 'disabled placeholder="isi bobot dulu"' : ''}></label>
      </div>
      <p class="muted" style="font-size:13px;margin:0 0 8px">${x.capaian == null ? '' : `Tercapai ${pc(x.capaian)} dari target (actual ${pc(aktual(x))} ÷ bobot ${pc(x.bobot)}) · `}Progress s/d kemarin ${pc(kem)}${samaHari(x).length ? ` + janji lain hari ini ${pc(lain)}` : ''} + actual ${pc(kontrib(x))} = <b>kumulatif ${pc(progKum(x))}</b> · sisa task ${pc(Math.max(0, sisaTask(x)))}</p>
      ${tercapai(x) === 'Tidak' ? `<label class="f">Penyebab tidak tercapai<select data-f="${f}penyebab" class="${x.penyebab ? '' : 'need'}">${opt('', x.penyebab, '– pilih –')}${D.penyebab.map(p => opt(p, x.penyebab)).join('')}</select></label>` : ''}
      <div class="bar" style="margin:0"><input class="grow" type="text" placeholder="Catatan" value="${esc(x.catatan)}" data-f="${f}catatan"><button class="btn sm danger" data-act="daily-del" data-id="${x.id}">Hapus</button></div>
  </div>`;
}

/* ---------- HARIAN (Daily Log + PPC Harian + Board 4) ---------- */
/* Tampilan Harian: 1 hari, atau 2 hari sekaligus untuk rutinitas update:
   kh = kemarin (lapor actual) + hari ini (buat janji) · hb = hari ini (lapor actual) + besok (rencana janji) */
const HMODE = { '1': '1 hari', kh: 'Kemarin + hari ini', hb: 'Hari ini + besok' };
function harianHari() {
  const a = UI.harianTgl || tgl(), m = D.settings.hMode || '1';
  if (m === 'kh') return [[addDays(a, -1), 'Kemarin · laporkan actual'], [a, 'Hari ini · buat janji']];
  if (m === 'hb') return [[a, 'Hari ini · laporkan actual'], [addDays(a, 1), 'Besok · rencana janji']];
  return [[a, '']];
}
function harianGroups(d) { // 1 header per task komitmen, berisi janji-janjinya pada tanggal d
  const rows = D.daily.filter(x => x.tgl === d), byKode = Object.fromEntries(D.tasks.map(t => [t.kode, t]));
  return [...new Set(rows.map(x => x.kode))].filter(k => !UI.fPrio || prioOK(byKode[k])).sort((a, b) => (byKode[a] && byKode[b] ? byPrio(byKode[a], byKode[b]) : 0) || a.localeCompare(b)).map(kode => {
    const t = byKode[kode], js = rows.filter(x => x.kode === kode).sort((a, b) => String(a.id).localeCompare(String(b.id)));
    const lap = js.filter(x => tercapai(x)), ya = js.filter(x => tercapai(x) === 'Ya').length;
    const pHari = Math.max(...js.map(progKum)), pAwal = progKemarin(js[0]), bobot = r6(js.reduce((s, x) => s + (nOr(x.bobot) || 0), 0));
    const warn = js.filter(x => cek(x, t).startsWith('⚠')).length, sp = t ? stP(t) : '';
    return `<div class="item grp" id="h-${esc(d)}-${esc(kode)}">
      <div class="hd"><div><span class="code">${esc(kode)}</span> <span class="ttl">${esc(t?.uraian || '(task tidak ada)')}</span>
        <div class="sub">${esc([t?.pekerjaan, t?.lokasi, t?.pic && 'PIC ' + t.pic].filter(Boolean).join(' · '))}</div></div>
        <div class="chips">${t ? prioChips(t) : ''}<span class="chip">${js.length} janji</span>${lap.length ? `<span class="chip ${ya === js.length ? 'ok' : 'bad'}">${ya}/${js.length} tercapai</span>` : '<span class="chip warn">belum lapor</span>'}${warn ? `<span class="chip warn">⚠ ${warn}</span>` : ''}${stChip(sp)}</div></div>
      <div class="grp-prog"><div class="bar-h"><i style="width:${Math.min(pAwal, 100)}%"></i><i class="today" style="width:${Math.max(0, Math.min(pHari, 100) - Math.min(pAwal, 100))}%"></i></div>
        <small class="muted">Progress ${pc(pAwal)} → <b>${pc(pHari)}</b> · bobot ${fmt(d)} ${pc(bobot)} · rencana selesai ${fmt(t?.selesai)}</small></div>
      ${js.map((x, i) => janjiCard(x, t, i + 1, js.length)).join('')}
      <div class="bar" style="margin:8px 0 0"><button class="btn sm" data-act="daily-new" data-kode="${esc(kode)}" data-tgl="${d}">+ Janji untuk task ini</button></div>
    </div>`;
  }).join('') || `<div class="empty">Belum ada janji untuk ${fmt(d)}.</div>`;
}
function vHarian() {
  const hari = harianHari(), dua = hari.length > 1;
  const d = hari[0][0]; // KPI: hari yang dilaporkan actual-nya
  const P = ppcDay(d), tgt = D.meta.targetPPC;
  const dates = [...new Set(D.daily.map(x => x.tgl))].filter(x => x <= tgl()).sort().slice(-14);
  const par = {};
  D.daily.filter(x => tercapai(x) === 'Tidak').forEach(x => { const k = x.penyebab || '(belum diisi)'; par[k] = (par[k] || 0) + 1; });
  const parL = Object.entries(par).sort((a, b) => b[1] - a[1]), parT = parL.reduce((s, x) => s + x[1], 0);
  const weeks = [...new Set(D.daily.map(x => weekOf(x.tgl)))].sort((a, b) => a - b);


  return `
  ${vh('Daily commitment', 'harian')}
  <div class="prio-f">${Object.entries(HMODE).map(([k, l]) => `<button type="button" data-act="hmode" data-v="${k}" class="${(D.settings.hMode || '1') === k ? 'on' : ''}">${l}</button>`).join('')}</div>
  ${prioBar()}
  <div class="bar">
    <button class="btn sm" data-act="day" data-n="-1" aria-label="Hari sebelumnya">‹</button>
    <input type="date" value="${UI.harianTgl || tgl()}" data-ui="harianTgl">
    <button class="btn sm" data-act="day" data-n="1" aria-label="Hari berikutnya">›</button>
    <span class="muted">W${weekOf(UI.harianTgl || tgl())}${dua ? ` · ${hari.map(h => fmt(h[0])).join(' & ')}` : ''}</span>
  </div>
  <div class="grid kpis">
    <div class="kpi"><b>${P.n}</b><span>Dijanjikan</span></div>
    <div class="kpi ok"><b>${P.ya}</b><span>Tercapai</span></div>
    <div class="kpi bad"><b>${P.tidak}</b><span>Tidak tercapai</span></div>
    <div class="kpi warn"><b>${P.n - P.ya - P.tidak}</b><span>Belum dilaporkan</span></div>
    <div class="kpi ${P.ppc == null ? '' : P.ppc >= tgt ? 'ok' : 'bad'}"><b>${pct(P.ppc)}</b><span>PPC ${dua ? fmt(d) : 'hari ini'} (target ${pct(tgt)})</span></div>
    <div class="kpi"><b>${pct(P.rata)}</b><span>Rata-rata % tercapai</span></div>
  </div>
  ${dua ? `<div class="hari2">${hari.map(([hd, lbl], i) => `<section class="hari-kol">
      <div class="hari-hd"><div><b>${esc(lbl)}</b><span class="muted">${fmt(hd)} · W${weekOf(hd)}</span></div>${(() => { const p = ppcDay(hd); return `<span class="chip ${p.n && p.ya + p.tidak === p.n ? 'ok' : p.n ? 'warn' : ''}">${p.n} janji${p.n ? ` · ${p.ya + p.tidak}/${p.n} lapor` : ''}</span>`; })()}</div>
      <div class="bar"><button class="btn ${i ? 'pri' : ''} sm" data-act="daily-new" data-tgl="${hd}">+ Janji ${fmt(hd)}</button>${i ? `<button class="btn sm" data-act="daily-copy" data-tgl="${hd}" data-from="${hari[0][0]}">Salin dari ${fmt(hari[0][0])}</button>` : ''}</div>
      ${harianGroups(hd)}</section>`).join('')}</div>`
    : `<div class="bar"><button class="btn pri" data-act="daily-new">+ Janji hari ini</button><button class="btn" data-act="daily-copy">Salin janji kemarin</button></div>
  ${harianGroups(d)}`}

  <div class="two">
    <div class="card"><h3>PPC harian (14 tanggal terakhir)</h3>${ppcChart(dates)}</div>
    <div class="card">
      <h3>PPC mingguan (Daily Log)</h3>
      <table class="tbl"><tr><th>Minggu</th><th class="n">Dijanjikan</th><th class="n">Tercapai</th><th class="n">PPC</th></tr>
      ${weeks.filter(w => w <= weekOf(tgl())).map(w => { const p = ppcJanji(w); return `<tr><td>W${w}</td><td class="n">${p.n}</td><td class="n">${p.ya}</td><td class="n"><b>${pct(p.ppc)}</b></td></tr>`; }).join('') || '<tr><td colspan="4" class="muted">–</td></tr>'}
      </table>
      <h3 style="margin-top:14px">Pareto penyebab tidak tercapai</h3>
      ${parL.length ? (() => { let cum = 0; return parL.map(([k, v]) => { cum += v; return `<div style="margin:6px 0"><div style="display:flex;justify-content:space-between;font-size:13px"><span>${esc(k)}</span><span class="muted">${v} · ${pct(v / parT)} · kum ${pct(cum / parT)}</span></div><div class="bar-h"><i style="width:${v / parL[0][1] * 100}%;background:var(--bad)"></i></div></div>`; }).join(''); })() : '<p class="muted">Belum ada penyebab tercatat.</p>'}
    </div>
  </div>`;
}

function ppcChart(dates) {
  if (!dates.length) return '<p class="muted">Belum ada data.</p>';
  const W = 460, H = 180, L = 32, B = 26, T = 10, n = dates.length, bw = (W - L - 8) / n;
  const y = v => T + (H - T - B) * (1 - v);
  const tgt = D.meta.targetPPC;
  let g = '';
  for (const v of [0, .5, 1]) g += `<line x1="${L}" x2="${W - 8}" y1="${y(v)}" y2="${y(v)}" stroke="var(--grid)" stroke-width="1"/><text x="${L - 5}" y="${y(v) + 4}" text-anchor="end">${v * 100}%</text>`;
  dates.forEach((d, i) => {
    const p = ppcDay(d), v = p.ppc ?? 0, x = L + i * bw + bw * 0.18;
    g += `<rect x="${x}" y="${y(v)}" width="${bw * 0.64}" height="${H - B - y(v)}" rx="3" fill="${v >= tgt ? 'var(--ok)' : 'var(--bad)'}" stroke="none"><title>${fmt(d)}: ${pct(p.ppc)} (${p.ya}/${p.n})</title></rect>`;
    if (n <= 7 || i % 2 === 0) g += `<text x="${x + bw * 0.32}" y="${H - 8}" text-anchor="middle">${d.slice(8)}/${+d.slice(5, 7)}</text>`;
  });
  g += `<line x1="${L}" x2="${W - 8}" y1="${y(tgt)}" y2="${y(tgt)}" stroke="var(--warn)" stroke-dasharray="5 4" stroke-width="1.5"/>`;
  return `<svg class="chart" viewBox="0 0 ${W} ${H}" role="img" aria-label="PPC harian">${g}</svg>
  <div class="legend"><span><i style="background:var(--ok)"></i>≥ target</span><span><i style="background:var(--bad)"></i>< target</span><span><i style="background:var(--warn)"></i>Target ${pct(tgt)}</span></div>`;
}

/* ---------- ISSUE (Follow Up Issue) ---------- */
function vIssue() {
  const L = issueList().filter(x => prioOK(x.t)).sort((a, b) => byPrio(a.t, b.t));
  return `
  ${vh('Follow up issue / kendala', 'issue')}
  ${prioBar()}
  <div class="card scroll"><table class="tbl">
    <tr><th>Kategori</th><th class="n">Pending</th><th class="n">Lewat target</th></tr>
    ${KK.map(k => { const a = L.filter(x => x.k === k); return `<tr><td>${KAT[k]}</td><td class="n">${a.length}</td><td class="n">${a.filter(x => x.sisa != null && x.sisa < 0).length}</td></tr>`; }).join('')}
    <tr><td><b>Total</b></td><td class="n"><b>${L.length}</b></td><td class="n"><b>${L.filter(x => x.sisa != null && x.sisa < 0).length}</b></td></tr>
  </table></div>
  ${L.map(x => `
    <div class="item">
      <div class="hd"><div><span class="code">${esc(x.t.kode)}</span> <span class="ttl">${esc(x.c.issue || '(issue belum dipilih)')}</span>
        <div class="sub">${esc(KAT[x.k])}${x.c.detail ? ' — ' + esc(x.c.detail) : ''}</div></div>
        <div class="chips">${prioChips(x.t)}${stChip(x.st)}</div></div>
      <div class="meta"><span>${esc(x.t.uraian)}</span><span>${esc(x.t.area)}</span><span>W${x.t.minggu}</span></div>
      <div class="meta"><span>PIC <b>${esc(x.c.pic || '–')}</b></span><span>Target <b>${fmt(x.c.target)}</b></span><span>${x.sisa == null ? '' : x.sisa < 0 ? `<b style="color:var(--bad)">lewat ${-x.sisa} hari</b>` : `sisa ${x.sisa} hari`}</span>
        <a href="#ready" data-act="goto-ready" data-kode="${esc(x.t.kode)}">Update →</a></div>
    </div>`).join('') || '<div class="empty">Tidak ada kendala pending.</div>'}`;
}

/* ---------- DATA & PENGATURAN ---------- */
function settingsForm() {
  const s = D.settings;
  return `<h3>Data & pengaturan</h3>
  <label class="f">Tanggal update<select name="tglMode">${opt('auto', s.tglMode, 'Otomatis = hari ini')}${opt('fixed', s.tglMode, 'Tanggal tertentu')}</select></label>
  <label class="f">Tanggal (bila tertentu)<input type="date" name="tglFixed" value="${esc(s.tglFixed || D.meta.tglUpdate || '')}"></label>
  <label class="f">Target PPC (%)<input type="number" name="targetPPC" min="0" max="100" value="${Math.round(D.meta.targetPPC * 100)}"></label>
  <fieldset style="border:1px solid var(--line);border-radius:10px;padding:10px 12px;margin:0 0 12px">
    <legend style="font-weight:600;font-size:14px;padding:0 4px">Sinkron tim (Google Sheets)</legend>
    <label class="f">URL Web App (…/exec)<input type="text" name="syncUrl" inputmode="url" placeholder="https://script.google.com/macros/s/…/exec" value="${esc(D.sync.url)}"></label>
    <div class="frow">
      <label class="f">Kode akses tim<input type="text" name="syncKey" autocomplete="off" value="${esc(D.sync.key)}"></label>
      <label class="f">Nama Anda<input type="text" name="syncUser" placeholder="mis. Agus SM" value="${esc(D.sync.user)}"></label>
    </div>
    <p class="muted" style="font-size:13px;margin:0 0 8px">${syncOn() ? `Status: ${esc($('#sync')?.title || '')}${D.sync.lastOk ? ' · terakhir ' + esc(new Date(D.sync.lastOk).toLocaleString('id-ID')) : ''}` : 'Belum terhubung. Data hanya di perangkat ini.'}</p>
    ${syncOn() ? '<div style="display:flex;gap:8px;flex-wrap:wrap"><button type="button" class="btn sm" data-act="sync-now">Sinkron sekarang</button><button type="button" class="btn sm danger" data-act="sync-off">Putuskan</button></div>' : ''}
  </fieldset>
  <p class="muted" style="font-size:13px">Sumber: <b>${esc(D.source || '-')}</b><br>Diekspor dari Excel: ${esc((D.exportedAt || '').replace('T', ' '))}<br>
  ${D.tasks.length} task · ${D.daily.length} daily log · ${D.pekerjaan.length} pekerjaan. ${syncOn() ? 'Disalin di perangkat ini & disinkron ke Google Sheet tim.' : 'Data tersimpan di perangkat ini.'}</p>
  <div class="grid" style="grid-template-columns:1fr 1fr">
    <button type="button" class="btn" data-act="exp-json">Backup JSON</button>
    <button type="button" class="btn" data-act="imp-json">Impor JSON</button>
    <button type="button" class="btn" data-act="exp-csv" data-what="tasks">CSV LPS Input</button>
    <button type="button" class="btn" data-act="exp-csv" data-what="daily">CSV Daily Log</button>
    <button type="button" class="btn" data-act="reseed-ref">${syncOn() ? 'Ambil ulang jadwal tim' : 'Perbarui jadwal dari Excel'}</button>
    <button type="button" class="btn danger" data-act="reseed-all">${syncOn() ? 'Reset & ambil ulang data tim' : 'Reset semua ke Excel'}</button>
  </div>
  <details style="margin-top:14px"><summary>Petunjuk singkat LPS</summary>
    <ol style="font-size:13px;padding-left:18px;margin:0">
      <li><b>Jadwal</b>: lihat pekerjaan MASTER yang aktif (lookahead).</li>
      <li><b>Komitmen</b>: buat task detail, saat rapat mingguan tandai <b>Disepakati</b>.</li>
      <li><b>Make Ready</b>: cek 4 kendala; bila Pending wajib isi issue, PIC & target. Semua OK = <b>READY</b>.</li>
      <li><b>Harian</b>: pagi isi janji, target & bobot (% dari task); sore isi <b>actual progress hari ini (% dari task)</b> & penyebab bila actual &lt; bobot. % tercapai, Tercapai? & progress dihitung otomatis.</li>
      <li>Progress 100% = <b>FINISH</b>; lewat rencana selesai = <b>LATE / DELAY</b>. PPC = tercapai ÷ dijanjikan (target ≥ 80%).</li>
    </ol></details>
  <div class="dlg-act"><button value="cancel" class="btn" formnovalidate>Tutup</button><button value="save" class="btn pri">Simpan</button></div>`;
}

function download(name, text, type) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([text], { type }));
  a.download = name; a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 1000);
}
function csv(rows) {
  return '﻿' + rows.map(r => r.map(v => { v = v ?? ''; v = String(v); return /[",\n;]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v; }).join(',')).join('\r\n');
}
function exportCSV(what) {
  if (what === 'tasks') {
    const h = ['Kode Task', 'Minggu Ke', 'Disiplin', 'Pekerjaan (Master Schedule)', 'Area Gedung', 'Uraian Pekerjaan', 'Lokasi / Area', 'PIC', 'Rencana Mulai', 'Rencana Selesai', 'Disepakati', 'Alasan / Catatan Rapat'];
    KK.forEach(k => h.push(KAT[k], 'Issue ' + KAT[k], 'Detail Issue ' + KAT[k], 'PIC ' + KAT[k], 'Target ' + KAT[k]));
    h.push('Catatan Kendala', 'Progress (%)', 'Penyebab Terlambat', 'Status Kendala', 'Status Pelaksanaan', 'Prioritas', 'Koordinasi Disiplin', 'Catatan Follow-up');
    const rows = D.tasks.map(t => [t.kode, t.minggu, t.disiplin, t.pekerjaan, t.area, t.uraian, t.lokasi, t.pic, t.mulai, t.selesai, t.disepakati, t.alasan,
      ...KK.flatMap(k => [t.k[k].st, t.k[k].issue, t.k[k].detail, t.k[k].pic, t.k[k].target]),
      t.catatanKendala, prog(t.kode), t.penyebabTerlambat, t.disepakati === 'Ya' ? stK(t) : '', t.disepakati === 'Ya' ? stP(t) : '', t.prio, t.koord, t.followup]);
    download(`LPS-Input_${tgl()}.csv`, csv([h, ...rows]), 'text/csv');
  } else {
    const h = ['Tanggal', 'Kode Task', 'Target Hari Ini', 'Bobot Target (% dari task)', '% Tercapai Harian (actual ÷ bobot)', 'Tercapai?', 'Progress s/d Kemarin', 'Kontribusi Hari Ini', 'Progress Kumulatif (%)', 'Penyebab Tidak Tercapai', 'Catatan'];
    const rows = [...D.daily].sort((a, b) => a.tgl.localeCompare(b.tgl)).map(x => [x.tgl, x.kode, x.target, x.bobot, x.capaian, tercapai(x), progKemarin(x), kontrib(x), progKum(x), x.penyebab, x.catatan]);
    download(`Daily-Log_${tgl()}.csv`, csv([h, ...rows]), 'text/csv');
  }
}

/* ---------- PDF per PIC (cetak browser → "Simpan sebagai PDF"; tanpa library, jalan offline) ---------- */
/* ---------- SUMBER DAYA: Manpower (MP-Master + aktual harian) & Dokumen (dokument tracking) ---------- */
const GRUP_MP = ['STAF', 'CSA', 'MEP'];
const DOK_ST = ['Draft', 'Submitted', 'Approved', 'Revisi', 'Rejected'];
const dokChip = s => `<span class="chip ${({ Approved: 'ok', Submitted: 'info', Revisi: 'warn', Rejected: 'bad' })[s] || ''}">${esc(s)}</span>`;
function normMp(x) { x.aktual = nOr(x.aktual); x.catatan = x.catatan ?? ''; return x; }
function normDok(x) {
  ['scope', 'nama', 'nomor', 'issue', 'pic', 'kode', 'link', 'riwayat'].forEach(f => { x[f] = x[f] ?? ''; });
  x.status = DOK_ST.includes(x.status) ? x.status : 'Draft'; x.rev = +x.rev || 0;
  ['target', 'tglSubmit', 'tglApprove'].forEach(f => { x[f] = x[f] || null; });
  return x;
}
const mpId = (d, item) => `m${d}_${item}`; // 1 record per aktivitas per tanggal (konflik antar HP: yang terakhir menang)
const mpAkt = (d, item) => D.manpower.find(x => x.id === mpId(d, item))?.aktual ?? null;
const mpPlanW = (it, w) => +(it.plan?.[w] || 0);
const mpHari = d => D.mpPlan.reduce((s, it) => s + (mpAkt(d, it.id) || 0), 0);
/* rencana mingguan (MP-Master) vs aktual = rata-rata harian dari hari yang ada input */
function mpWeek(w) {
  const rows = D.manpower.filter(x => x.aktual != null && weekOf(x.tgl) === +w), days = [...new Set(rows.map(x => x.tgl))];
  const g = Object.fromEntries(GRUP_MP.map(k => {
    const its = D.mpPlan.filter(i => i.grup === k), ids = new Set(its.map(i => i.id));
    return [k, { plan: its.reduce((s, i) => s + mpPlanW(i, w), 0), akt: days.length ? Math.round(rows.filter(x => ids.has(x.item)).reduce((s, x) => s + x.aktual, 0) / days.length) : null }];
  }));
  return { g, hari: days.length, plan: GRUP_MP.reduce((s, k) => s + g[k].plan, 0), akt: days.length ? Math.round(rows.reduce((s, x) => s + x.aktual, 0) / days.length) : null };
}
// dokumen Drive: target = rencana SUBMIT (lewat bila belum submit); dokumen manual: target = rencana APPROVE
const lewatDok = x => x.drive ? !x.tglSubmit && !!x.target && x.target < tgl() : x.status !== 'Approved' && !!x.target && x.target < tgl();
// umur sejak submit aktual: berjalan sampai ada keputusan (Approved / Not approved / Rejected)
const umurDok = x => x.tglSubmit && x.status === 'Submitted' ? days(x.tglSubmit, tgl()) : null;
const umurChip = (x, pj) => { const n = umurDok(x); return n == null ? '' : `<span class="chip ${n > 14 ? 'bad' : n > 7 ? 'warn' : ''}">${n} hari${pj ? ' sejak submit' : ''}</span>`; };
// menunggu siapa (dokumen Submitted/Open): Safran bila CW sudah respon atau remarks "awaiting owner", selain itu CW; For Info & yang sudah dibalas Safran tidak dihitung
const tungguDok = x => umurDok(x) == null || x.own || /for info/i.test(x.stAsli || '') ? '' : x.cw || /awaiting owner/i.test(x.issue || '') ? 'own' : 'cw';
const DOK_FX = { lewat3: x => umurDok(x) > 3, cw: x => tungguDok(x) === 'cw', own: x => tungguDok(x) === 'own' };
const dokOfTask = kode => D.dokumen.filter(x => String(x.kode).split(',').map(s => s.trim()).includes(kode));

/* ---------- Dokumen dari log Google Drive (HANYA BACA lewat link berbagi "Pelihat"; tidak pernah menulis ke Drive) ----------
   Tiap sheet "LOG LIST OF …" (kolom B–L: No, Document Number, Rev, Description, Estimated Submission, Actual Submission,
   Response CW, Response Safran, Status, Remarks, PIC) dibaca via endpoint CSV Google (gviz). Link disimpan di master tim
   (push_master.py --dok-link) atau per perangkat; tidak ada di kode publik. Hasil di-cache lokal agar tetap tampil offline. */
const DOK_KEY = 'lps-batik-dok';
let DRV = (() => { try { return JSON.parse(localStorage.getItem(DOK_KEY)) || {}; } catch { return {}; } })();
let drvBusy = false;
const dokSrc = () => D.settings.dokSrc || D.dokSrc || '';
const drvId = s => (String(s || '').match(/\/d\/([\w-]{20,})/) || String(s || '').trim().match(/^([\w-]{25,})$/) || [])[1] || '';
const drvDocs = () => (DRV.src && DRV.src === drvId(dokSrc()) && DRV.docs) || [];
const allDok = () => [...D.dokumen, ...drvDocs()];
const BLN_N = { jan: 1, feb: 2, mar: 3, apr: 4, mei: 5, may: 5, jun: 6, jul: 7, agu: 8, agt: 8, aug: 8, sep: 9, okt: 10, oct: 10, nov: 11, des: 12, dec: 12 };
function tglTeks(s) {
  s = String(s || '').trim(); if (!s) return null;
  let m = s.match(/^\d{4}-\d{2}-\d{2}/); if (m) return m[0];
  const p2 = v => String(v).padStart(2, '0');
  m = s.match(/^(\d{1,2})[\s-]+([A-Za-z]{3})[A-Za-z]*\.?[\s-]+(\d{4})/);
  if (m && BLN_N[m[2].toLowerCase()]) return `${m[3]}-${p2(BLN_N[m[2].toLowerCase()])}-${p2(m[1])}`;
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/); return m ? `${m[3]}-${p2(m[2])}-${p2(m[1])}` : null;
}
function parseCsv(t) {
  const rows = []; let r = [], f = '', q = false;
  for (let i = 0; i < t.length; i++) {
    const c = t[i];
    if (q) { if (c === '"') { if (t[i + 1] === '"') { f += '"'; i++; } else q = false; } else f += c; }
    else if (c === '"') q = true; else if (c === ',') { r.push(f); f = ''; } else if (c === '\n') { r.push(f); rows.push(r); r = []; f = ''; } else if (c !== '\r') f += c;
  }
  if (f || r.length) { r.push(f); rows.push(r); }
  return rows;
}
const stDrv = s => { s = s.toLowerCase(); return /reject/.test(s) ? 'Rejected' : /not approved|revis/.test(s) ? 'Revisi' : /approved|closed/.test(s) ? 'Approved' : /submit|open|review/.test(s) ? 'Submitted' : 'Draft'; };
function drvSheet(name, gid, rows, id) {
  const hi = rows.findIndex(r => /document number/i.test(r[2] || ''));
  if (hi < 0 || !rows.slice(0, hi).some(r => /log list/i.test(r[1] || ''))) return []; // bukan sheet log (Numbering, Monitoring, …)
  const jenis = name.trim(), by = new Map();
  let scope = '', sub = '';
  for (const r of rows.slice(hi + 1)) {
    const c = i => String(r[i] ?? '').trim(), no = c(2), desc = c(4);
    if (!no && !desc) { // baris judul seksi: "SHOP DRAWING - CSA" → scope, "STRUCTURE" → sub
      const s = c(1); if (!s) continue;
      const m = s.match(/[-–]\s*(CSA|MEP|ARS|PRELIM)\b/i);
      if (m) { scope = m[1].toUpperCase(); sub = ''; } else sub = s.replace(/^(method statement|jsa)\s*[-–]\s*/i, '');
      continue;
    }
    const key = /^SBP-/i.test(no) ? no.toUpperCase() : no + '|' + desc;
    if (!by.has(key)) by.set(key, { no, desc, scope: scope || ((no.match(/-(CSA|MEP)-/i) || [])[1] || '').toUpperCase(), sub, revs: [] });
    by.get(key).revs.push({ rev: +c(3) || 0, est: tglTeks(c(5)), sub: tglTeks(c(6)), cw: tglTeks(c(7)), own: tglTeks(c(8)), st: c(9), rem: c(10), pic: c(11) });
  }
  const link = `https://docs.google.com/spreadsheets/d/${id}/edit#gid=${gid}`;
  return [...by.entries()].map(([key, d]) => {
    const revs = d.revs.sort((a, b) => a.rev - b.rev), x = revs.at(-1), status = stDrv(x.st || '');
    return { id: 'g:' + jenis + ':' + key, drive: true, jenis, scope: d.scope, sub: d.sub, nomor: d.no, nama: d.desc || d.no, rev: x.rev,
      status, stAsli: x.st || 'Belum submit', target: x.est, tglSubmit: x.sub, cw: x.cw, own: x.own, respon: x.own || x.cw || null, tglApprove: status === 'Approved' ? (x.own || x.cw || null) : null,
      issue: x.rem, pic: x.pic, kode: '', link,
      riwayat: revs.length < 2 ? "" : revs.map(v => `Rev ${v.rev} · ${v.st || '-'}${v.sub ? ' · submit ' + fmt(v.sub) : ''}${v.own || v.cw ? ' · respon ' + fmt(v.own || v.cw) : ''}${v.rem ? '\n   ' + v.rem.replace(/\s*\n\s*/g, ' ') : ''}`).join('\n') };
  });
}
async function drvLoad(manual) {
  const id = drvId(dokSrc()); if (!id || drvBusy) return;
  drvBusy = true; if (UI.view === 'sd') render();
  const base = `https://docs.google.com/spreadsheets/d/${id}/`;
  try {
    const html = await (await fetch(base + 'htmlview')).text();
    const sh = [...html.matchAll(/\{name: "((?:[^"\\]|\\.)*)", pageUrl: "[^"]*", gid: "(\d+)"/g)]
      .map(m => [m[1].replace(/\\x([0-9a-f]{2})/gi, (_, h) => String.fromCharCode(parseInt(h, 16))).replace(/\\(.)/g, '$1'), m[2]]);
    if (!sh.length) throw new Error('Daftar sheet tidak terbaca. Pastikan akses link: "Siapa saja yang memiliki link" (Pelihat).');
    const docs = (await Promise.all(sh.map(async ([n, gid]) => {
      const r = await fetch(base + 'gviz/tq?tqx=out:csv&headers=0&gid=' + gid);
      if (!r.ok) throw new Error(`Sheet ${n}: HTTP ${r.status}`);
      return drvSheet(n, gid, parseCsv(await r.text()), id);
    }))).flat();
    DRV = { src: id, at: new Date().toISOString(), docs, err: '' };
  } catch (e) {
    DRV = { src: id, docs: DRV.src === id ? DRV.docs : [], at: new Date().toISOString(), err: /fetch|network/i.test(e.message) ? 'Tidak bisa membaca Drive (offline, atau link belum dibagikan "Siapa saja yang memiliki link")' : e.message };
  }
  drvBusy = false;
  try { localStorage.setItem(DOK_KEY, JSON.stringify(DRV)); } catch { /* penyimpanan penuh: tetap tampil di sesi ini */ }
  render();
  if (manual) toast(DRV.err ? DRV.err : `${DRV.docs.length} dokumen dibaca dari Drive`);
}
// buka tab Dokumen → baca ulang otomatis bila cache > 10 menit
const drvAuto = () => { if (drvId(dokSrc()) && !drvBusy && (!DRV.at || DRV.src !== drvId(dokSrc()) || Date.now() - new Date(DRV.at) > 6e5)) setTimeout(() => drvLoad(false)); };

function vSD() {
  return `<div class="vh"><h2>Sumber daya</h2></div>
  <div class="prio-f">${[['mp', 'Manpower'], ['dok', 'Dokumen']].map(([k, l]) => `<button type="button" data-act="sd-tab" data-v="${k}" class="${UI.sd === k ? 'on' : ''}">${l}</button>`).join('')}</div>
  ${UI.sd === 'dok' ? vDok() : vMp()}`;
}

function vMp() {
  if (!D.mpPlan.length) return '<div class="empty">Rencana manpower belum ada. Taruh <b>Man Power Tracking.xlsx</b> di folder yang sama dengan Excel LPS, lalu jalankan export_excel.py.</div>';
  const d = UI.mpTgl || tgl(), w = weekOf(d), W = mpWeek(w), tot = mpHari(d), ada = D.mpPlan.some(it => mpAkt(d, it.id) != null);
  const items = D.mpPlan.filter(it => UI.mpAll || mpPlanW(it, w) > 0 || mpAkt(d, it.id) != null);
  const gHari = g => D.mpPlan.filter(i => i.grup === g).reduce((s, i) => s + (mpAkt(d, i.id) || 0), 0);
  let list = '', lastG = '', lastS = null;
  for (const it of items) {
    if (it.grup !== lastG) { list += `<h3 class="mp-g">${esc(it.grup)}<small class="muted">rencana ${W.g[it.grup]?.plan ?? 0} · hari ini ${gHari(it.grup)}</small></h3>`; lastG = it.grup; lastS = null; }
    if (it.sub !== lastS) { if (it.sub) list += `<div class="mp-s">${esc(it.sub)}</div>`; lastS = it.sub; }
    const p = mpPlanW(it, w), a = mpAkt(d, it.id), st = a == null ? '' : a < p ? 'bad' : a > p ? 'info' : 'ok';
    list += `<div class="mp-row"><div class="mp-n"><b>${esc(it.nama)}</b><small class="muted">rencana ${p} org</small></div>
      <input type="number" min="0" max="999" step="1" inputmode="numeric" value="${a ?? ''}" data-mp="${esc(it.id)}" aria-label="Aktual ${esc(it.nama)}">
      <span class="chip ${st}">${a == null ? '–' : (a - p > 0 ? '+' : '') + (a - p)}</span></div>`;
  }
  const weeks = D.meta.weeks.filter(x => x.w <= weekOf(tgl()) + 1 && x.w >= weekOf(tgl()) - 5);
  return `
  <div class="bar">
    <button class="btn sm" data-act="mp-day" data-n="-1" aria-label="Hari sebelumnya">‹</button>
    <input type="date" value="${d}" data-ui="mpTgl">
    <button class="btn sm" data-act="mp-day" data-n="1" aria-label="Hari berikutnya">›</button>
    <span class="muted">W${w}</span>
  </div>
  <div class="grid kpis">
    <div class="kpi"><b>${W.plan}</b><span>Rencana MP W${w} (MP-Master)</span></div>
    <div class="kpi ${!ada ? 'warn' : tot >= W.plan ? 'ok' : 'bad'}"><b>${ada ? tot : '–'}</b><span>Aktual hari ini${ada ? ' · ' + pct(W.plan ? tot / W.plan : null) + ' rencana' : ' · belum diisi'}</span></div>
    ${GRUP_MP.map(g => `<div class="kpi"><b>${ada ? gHari(g) : '–'}<small class="muted" style="font-size:13px"> / ${W.g[g].plan}</small></b><span>${g} aktual / rencana</span></div>`).join('')}
  </div>
  <div class="bar"><button class="btn" data-act="mp-copy">Salin aktual hari sebelumnya</button>
    <label class="muted" style="font-size:13px"><input type="checkbox" data-act="mp-all" ${UI.mpAll ? 'checked' : ''} style="width:auto;min-height:0"> tampilkan semua aktivitas</label></div>
  <div class="card mp-list">${list || '<p class="muted">Tidak ada aktivitas berencana minggu ini.</p>'}</div>
  <div class="two">
    <div class="card"><h3>Manpower loading – rencana vs aktual</h3>${mpChart()}</div>
    <div class="card"><h3>Rekap mingguan (rata-rata harian)</h3>
      <div class="scroll"><table class="tbl"><tr><th>Minggu</th>${GRUP_MP.map(g => `<th class="n">${g}</th>`).join('')}<th class="n">Total</th><th class="n">%</th></tr>
      ${weeks.map(x => { const r = mpWeek(x.w); return `<tr${x.w === weekOf(tgl()) ? ' style="background:var(--warn-bg)"' : ''}><td>W${x.w}<br><small class="muted">${esc(x.label)}</small></td>${GRUP_MP.map(g => `<td class="n">${r.g[g].akt ?? '–'}<small class="muted"> / ${r.g[g].plan}</small></td>`).join('')}<td class="n"><b>${r.akt ?? '–'}</b><small class="muted"> / ${r.plan}</small></td><td class="n">${pct(r.akt != null && r.plan ? r.akt / r.plan : null)}</td></tr>`; }).reverse().join('')}
      </table></div>
      <p class="muted" style="font-size:12px">Aktual mingguan = rata-rata jumlah orang per hari dari hari yang ada input.</p></div>
  </div>`;
}
function mpChart() {
  const ws = D.meta.weeks, W = 640, H = 220, L = 34, R = 10, T = 12, B = 26, n = ws.length;
  const data = ws.map(x => ({ w: x.w, plan: mpWeek(x.w).plan, akt: mpWeek(x.w).akt }));
  const max = Math.max(...data.map(x => Math.max(x.plan, x.akt || 0)), 10), bw = (W - L - R) / n;
  const y = v => T + (H - T - B) * (1 - v / max), cw = weekOf(tgl());
  let g = '';
  for (const v of [0, .25, .5, .75, 1]) g += `<line x1="${L}" x2="${W - R}" y1="${y(max * v)}" y2="${y(max * v)}" stroke="var(--grid)"/><text x="${L - 5}" y="${y(max * v) + 4}" text-anchor="end">${Math.round(max * v)}</text>`;
  data.forEach((x, i) => {
    const cx = L + bw * i + bw / 2;
    g += `<rect x="${cx - bw * .35}" y="${y(x.plan)}" width="${bw * .7}" height="${H - B - y(x.plan)}" fill="var(--plan)" opacity="${x.w === cw ? .9 : .45}"><title>W${x.w}: rencana ${x.plan}</title></rect>`;
    if (x.akt != null) g += `<circle cx="${cx}" cy="${y(x.akt)}" r="4" fill="var(--ok)" stroke="var(--surface)" stroke-width="1.5"><title>W${x.w}: aktual rata-rata ${x.akt}</title></circle>`;
    if (i % 3 === 0) g += `<text x="${cx}" y="${H - 8}" text-anchor="middle">W${x.w}</text>`;
  });
  const act = data.map((x, i) => x.akt == null ? null : `${L + bw * i + bw / 2},${y(x.akt)}`).filter(Boolean);
  if (act.length > 1) g += `<polyline points="${act.join(' ')}" stroke="var(--ok)" stroke-width="2"/>`;
  return `<svg class="chart" viewBox="0 0 ${W} ${H}" role="img" aria-label="Manpower loading">${g}</svg>
    <div class="legend"><span><i style="background:var(--plan);opacity:.6"></i>Rencana (MP-Master)</span><span><i style="background:var(--ok)"></i>Aktual rata-rata</span></div>`;
}

function vDok() {
  drvAuto();
  const f = UI.dokF, sc = UI.dokScope, jn = UI.dokJenis, sb = UI.dokSub, q = UI.dokQ.toLowerCase().trim(), all0 = allDok();
  const kat = x => x.jenis || 'Manual';
  const urut = new Map(); all0.forEach(x => { if (!urut.has(kat(x))) urut.set(kat(x), urut.size); if (x.sub && !urut.has(kat(x) + '|' + x.sub)) urut.set(kat(x) + '|' + x.sub, urut.size); });
  const scopeL = [...new Set([...D.disiplin, ...all0.map(x => x.scope).filter(Boolean)])];
  // filter selain kategori → dipakai untuk angka di chip kategori
  const base = all0.filter(x => (!sc || x.scope === sc) && (!q || [x.nomor, x.nama, x.pic, x.sub, x.issue].join(' ').toLowerCase().includes(q))
      && (!f || (DOK_FX[f] ? DOK_FX[f](x) : f === 'open' ? x.status !== 'Approved' : f === 'lewat' ? lewatDok(x) : x.status === f)));
  const all = jn ? all0.filter(x => kat(x) === jn) : all0; // KPI mengikuti kategori terpilih
  const subL = jn ? [...new Set(all.map(x => x.sub).filter(Boolean))] : [];
  const L = base.filter(x => (!jn || kat(x) === jn) && (!sb || x.sub === sb))
    .sort((a, b) => urut.get(kat(a)) - urut.get(kat(b)) || (urut.get(kat(a) + '|' + a.sub) ?? -1) - (urut.get(kat(b) + '|' + b.sub) ?? -1)
      || (lewatDok(b) - lewatDok(a)) || ((a.status === 'Approved') - (b.status === 'Approved')) || String(a.target || '9999').localeCompare(String(b.target || '9999')));
  const katL = [...urut.keys()].filter(k => !k.includes('|'));
  const nKat = k => base.filter(x => kat(x) === k).length;
  const n = s => all.filter(x => x.status === s).length, lw = all.filter(lewatDok).length;
  const FD = { '': 'Semua', open: 'Belum approve', lewat3: 'Submit > 3 hari', cw: 'Awaiting feedback (CW)', own: 'Awaiting owner (Safran)', lewat: 'Lewat target', Draft: 'Draft', Submitted: 'Submitted', Revisi: 'Revisi / Not Approved', Rejected: 'Rejected', Approved: 'Approved' };
  const id = drvId(dokSrc());
  const src = id ? `<div class="card" style="padding:10px 14px"><div class="bar" style="margin:0">
      <span class="muted" style="font-size:13px">Log dokumen Drive (hanya baca) · ${drvBusy ? 'membaca…' : DRV.at && DRV.src === id ? `${drvDocs().length} dokumen · dibaca ${new Date(DRV.at).toLocaleString('id-ID', { day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' })}` : 'belum dibaca'}</span>
      <span style="flex:1"></span><button class="btn sm" data-act="drv-load" ${drvBusy ? 'disabled' : ''}>Muat ulang</button><button class="btn sm" data-act="drv-set">Link</button></div>
      ${DRV.err && DRV.src === id ? `<p class="muted" style="color:var(--bad);margin:6px 0 0;font-size:13px">${esc(DRV.err)}</p>` : ''}</div>`
    : `<div class="card"><b>Ambil dokumen dari log Google Drive</b><p class="muted" style="font-size:13px">Tempel link Google Sheet log dokumen (akses link: "Siapa saja yang memiliki link" – Pelihat). Aplikasi hanya membaca, tidak mengubah file.</p><button class="btn pri" data-act="drv-set">Atur link Drive</button></div>`;
  return `${src}
  <div class="grid kpis">
    <div class="kpi"><b>${all.length}</b><span>Total dokumen</span></div>
    <div class="kpi ok"><b>${n('Approved')}</b><span>Approved</span></div>
    <div class="kpi info"><b>${n('Submitted')}</b><span>Menunggu review</span></div>
    <div class="kpi warn"><b>${n('Revisi') + n('Rejected')}</b><span>Not approved / Revisi / Rejected</span></div>
    <div class="kpi ${lw ? 'bad' : ''}"><b>${lw}</b><span>Lewat target</span></div>
  </div>
  <div class="prio-f">${Object.entries(FD).map(([k, l]) => `<button type="button" data-act="dok-f" data-v="${k}" class="${f === k ? 'on' : ''}">${l}${DOK_FX[k] ? ` <b>${all.filter(DOK_FX[k]).length}</b>` : ''}</button>`).join('')}</div>
  <div class="prio-f kat-f"><button type="button" data-act="dok-kat" data-v="" class="${jn ? '' : 'on'}">Semua kategori <b>${base.length}</b></button>${katL.map(k => `<button type="button" data-act="dok-kat" data-v="${esc(k)}" class="${jn === k ? 'on' : ''}">${esc(k)} <b>${nKat(k)}</b></button>`).join('')}</div>
  <div class="bar">${subL.length ? `<select data-ui="dokSub" style="width:auto">${opt('', sb, 'Semua sub-kategori')}${subL.map(d => opt(d, sb)).join('')}</select>` : ''}
    <select data-ui="dokScope" style="width:auto">${opt('', sc, 'Semua scope')}${scopeL.map(d => opt(d, sc)).join('')}</select>
    <input type="search" data-ui="dokQ" value="${esc(UI.dokQ)}" placeholder="Cari nomor / nama / PIC" style="flex:1;min-width:160px"></div>
  ${L.length ? `<div class="card dok-t"><table class="tbl">
    <thead><tr><th>Dokumen</th><th class="w-l">PIC</th><th>Status</th><th class="n">Tanggal</th></tr></thead>
    <tbody>${(() => {
      let h = '', gK = null, gS = null;
      for (const x of L) {
        const k = kat(x), tutup = UI.dokTutup.includes(k);
        if (k !== gK) {
          const g = L.filter(d => kat(d) === k), bu = g.filter(d => d.status !== 'Approved').length, lw = g.filter(lewatDok).length;
          h += `<tr class="dok-grp" data-act="dok-grp" data-v="${esc(k)}"><td colspan="4"><span class="cv">${tutup ? '▸' : '▾'}</span> <b>${esc(k)}</b>
            <small class="muted">${g.length} dokumen · ${bu} belum approve${lw ? ` · <span style="color:var(--bad)">${lw} lewat</span>` : ''}</small></td></tr>`;
          gK = k; gS = null;
        }
        if (tutup) continue;
        if (x.sub && x.sub !== gS && !sb) { h += `<tr class="dok-sub"><td colspan="4">${esc(x.sub)}</td></tr>`; gS = x.sub; }
        const lw = lewatDok(x), tl = (l, v) => v ? `<small class="muted">${l}</small> ${fmt(v)}` : '';
        const tanggal = [x.tglSubmit ? tl('Submit', x.tglSubmit) : tl(x.drive ? 'Rencana' : 'Target', x.target),
          x.drive ? [tl('CW', x.cw), tl('Safran', x.own)].filter(Boolean).join('<br>') : tl('Approve', x.tglApprove)].filter(Boolean).join('<br>');
        h += `<tr id="d-${esc(x.id)}" data-act="dok-info" data-id="${esc(x.id)}">
          <td><span class="code">${esc(x.nomor || '(tanpa nomor)')}${x.rev ? ' · Rev ' + x.rev : ''}</span><br>${esc(x.nama)}${x.issue ? `<small class="dok-rem">${esc(x.issue)}</small>` : ''}</td>
          <td class="w-l">${esc(x.pic || '–')}</td>
          <td>${dokStChip(x)}${lw ? `<br><span class="chip bad">lewat</span>` : umurDok(x) != null ? '<br>' + umurChip(x) : ''}</td>
          <td class="n dok-tg">${tanggal || '–'}</td></tr>`;
      }
      return h;
    })()}</tbody></table></div>
    <p class="muted" style="font-size:12px">${L.length} dokumen · ketuk judul kategori untuk buka/tutup, ketuk baris untuk detail</p>`
    : `<div class="empty">${drvBusy ? 'Membaca log dokumen dari Drive…' : 'Belum ada dokumen untuk filter ini.'}</div>`}
  <button class="btn pri fab" data-act="dok-new">+ Dokumen</button>`;
}
const dokStChip = x => x.drive && x.stAsli !== x.status ? `<span class="chip ${({ Approved: 'ok', Submitted: 'info', Revisi: 'warn', Rejected: 'bad' })[x.status] || ''}">${esc(x.stAsli)}</span>` : dokChip(x.status);
/* detail satu dokumen (dialog) – dokumen Drive hanya baca, dokumen manual bisa diubah statusnya */
function dokInfo(id) {
  const x = allDok().find(d => d.id === id); if (!x) return;
  const sisa = x.target ? days(tgl(), x.target) : null, tasks = String(x.kode).split(',').map(s => s.trim()).filter(Boolean);
  const next = x.drive ? [] : { Draft: [['Submitted', 'Submit']], Submitted: [['Approved', 'Approve'], ['Revisi', 'Revisi'], ['Rejected', 'Reject']], Revisi: [['Submitted', `Submit ulang (Rev ${x.rev + 1})`]], Rejected: [['Submitted', `Submit ulang (Rev ${x.rev + 1})`]] }[x.status] || [];
  const row = (k, v) => v ? `<tr><th>${k}</th><td>${v}</td></tr>` : '';
  openDlg(`<h3 style="margin-bottom:2px">${esc(x.nama)}</h3>
    <p class="code" style="margin:0 0 8px">${esc(x.nomor || '(tanpa nomor)')} · Rev ${x.rev}</p>
    <div class="chips" style="justify-content:flex-start">${dokStChip(x)}${lewatDok(x) ? `<span class="chip bad">lewat ${-sisa} hari</span>` : umurDok(x) != null ? umurChip(x, 1) : !x.drive && x.status !== 'Approved' && sisa != null ? `<span class="chip">sisa ${sisa} hari</span>` : ''}</div>
    <table class="tbl dok-d">
      ${row('Jenis', esc(x.jenis || 'Manual'))}${row('Scope', esc([x.scope, x.sub].filter(Boolean).join(' · ')))}${row('PIC', esc(x.pic))}
      ${x.drive ? row('Rencana submit', fmt(x.target)) + row('Submit aktual', x.tglSubmit && `${fmt(x.tglSubmit)} <small class="muted">· ${days(x.tglSubmit, tgl())} hari lalu</small>`)
          + row('Respon CW', x.cw && `${fmt(x.cw)} <small class="muted">· ${days(x.tglSubmit || x.cw, x.cw)} hari setelah submit</small>`)
          + row('Respon Safran', x.own && `${fmt(x.own)} <small class="muted">· ${days(x.tglSubmit || x.own, x.own)} hari setelah submit</small>`)
          + (!x.cw && !x.own && x.respon ? row('Respon', fmt(x.respon)) : '')
        : row('Target approve', fmt(x.target)) + row('Submit', fmt(x.tglSubmit)) + row('Approved', x.tglApprove && fmt(x.tglApprove))}
      ${row('Remarks / issue', x.issue && esc(x.issue).replace(/\n/g, '<br>'))}
      ${row('Task LPS', tasks.map(k => `<a href="#ready" data-act="goto-ready" data-kode="${esc(k)}" class="chip info" style="text-decoration:none">${esc(k)}</a>`).join(' '))}
    </table>
    ${x.riwayat ? `<details open style="margin-top:8px"><summary class="muted" style="font-size:13px">Riwayat revisi</summary><pre class="riwayat">${esc(x.riwayat)}</pre></details>` : ''}
    ${x.drive ? '<p class="muted" style="font-size:12px">Dari log Google Drive (hanya baca). Perubahan dilakukan di file Drive.</p>' : ''}
    <div class="dlg-act" style="flex-wrap:wrap">${next.map(([v, l]) => `<button type="button" class="btn sm ${v === 'Approved' ? 'pri' : v === 'Rejected' ? 'danger' : ''}" data-act="dok-st" data-id="${esc(x.id)}" data-v="${v}">${esc(l)}</button>`).join('')}
      <span class="sp"></span>${x.link ? `<a class="btn" href="${esc(x.link)}" target="_blank" rel="noopener">${x.drive ? 'Buka log' : 'Buka file'}</a>` : ''}${x.drive ? '' : `<button type="button" class="btn" data-act="dok-edit" data-id="${esc(x.id)}">Ubah</button>`}
      <button value="cancel" class="btn" formnovalidate>Tutup</button></div>`, null);
}
function drvSet() {
  openDlg(`<h3>Link log dokumen (Google Drive)</h3>
    <label class="f">Link Google Sheet<input type="text" name="src" inputmode="url" value="${esc(D.settings.dokSrc || '')}" placeholder="${D.dokSrc ? 'Dari Sheet tim (biarkan kosong)' : 'https://docs.google.com/spreadsheets/d/…'}"></label>
    <p class="muted" style="font-size:13px">Aplikasi hanya membaca file ini; tidak ada yang ditulis ke Drive. Akses link harus "Siapa saja yang memiliki link" (Pelihat).
    ${D.dokSrc ? '<br>Link tim sudah diatur admin (push_master.py --dok-link). Isi di sini hanya bila ingin memakai link lain di perangkat ini.' : '<br>Agar berlaku untuk semua anggota tim, admin menjalankan push_master.py dengan --dok-link.'}</p>
    <div class="dlg-act"><span class="sp"></span><button value="cancel" class="btn" formnovalidate>Batal</button><button value="save" class="btn pri">Simpan</button></div>`, fd => {
    const s = String(fd.get('src') || '').trim();
    if (s && !drvId(s)) { toast('Link tidak dikenali – salin link dari tombol Bagikan Google Sheet'); return; }
    D.settings.dokSrc = s; persist(); DRV = {}; render(); drvLoad(true);
  });
}
function dokForm(x, isNew) {
  const kode = String(x.kode).split(',').map(s => s.trim());
  return `<h3>${isNew ? 'Dokumen baru' : 'Ubah dokumen'}</h3>
  <label class="f">Nama dokumen<input type="text" name="nama" value="${esc(x.nama)}" required placeholder="mis. Shop drawing overlay lantai GD 4"></label>
  <div class="frow">
    <label class="f">Nomor dokumen<input type="text" name="nomor" value="${esc(x.nomor)}"></label>
    <label class="f">Revisi<input type="number" name="rev" min="0" step="1" value="${x.rev}"></label>
  </div>
  <div class="frow">
    <label class="f">Scope<select name="scope">${D.disiplin.map(d => opt(d, x.scope)).join('')}</select></label>
    <label class="f">Status<select name="status">${DOK_ST.map(s => opt(s, x.status)).join('')}</select></label>
  </div>
  <div class="frow">
    <label class="f">Target approved<input type="date" name="target" value="${esc(x.target || '')}"></label>
    <label class="f">Tanggal submit<input type="date" name="tglSubmit" value="${esc(x.tglSubmit || '')}"></label>
  </div>
  <div class="frow">
    <label class="f">PIC<input type="text" name="pic" value="${esc(x.pic)}"></label>
    <label class="f">Link file (Google Drive)<input type="text" name="link" inputmode="url" value="${esc(x.link)}" placeholder="https://drive.google.com/…"></label>
  </div>
  <label class="f">Issue (jika revisi / reject)<textarea name="issue">${esc(x.issue)}</textarea></label>
  <div class="f">Task LPS terkait (dokumen ini syarat Engineering task tsb)<div class="ck">${D.tasks.map(t => `<label><input type="checkbox" name="kode" value="${esc(t.kode)}"${kode.includes(t.kode) ? ' checked' : ''}> ${esc(t.kode)} <small class="muted">${esc(t.uraian.slice(0, 28))}</small></label>`).join('') || '<span class="muted">Belum ada task</span>'}</div></div>
  <div class="dlg-act">${isNew ? '' : `<button type="button" class="btn danger" data-act="dok-del" data-id="${esc(x.id)}">Hapus</button>`}<span class="sp"></span>
    <button value="cancel" class="btn" formnovalidate>Batal</button><button value="save" class="btn pri">Simpan</button></div>`;
}
const riwayat = (x, s) => { x.riwayat = (x.riwayat ? x.riwayat + '\n' : '') + `${fmt(tgl())} · ${s}${D.sync.user ? ' · ' + D.sync.user : ''}`; };
function editDok(x, isNew) {
  openDlg(dokForm(x, isNew), fd => {
    const v = Object.fromEntries(fd.entries()); v.kode = fd.getAll('kode').join(','); v.rev = +v.rev || 0;
    ['target', 'tglSubmit'].forEach(k => { v[k] = v[k] || null; });
    const r = isNew ? normDok({ id: 'doc' + Date.now().toString(36) + Math.random().toString(36).slice(2, 5), riwayat: '', tglApprove: null }) : D.dokumen.find(d => d.id === x.id);
    if (r.status !== v.status) riwayat(r, `${v.status} (Rev ${v.rev})`);
    if (v.status === 'Approved' && !r.tglApprove) v.tglApprove = tgl();
    Object.assign(r, v); normDok(r);
    if (isNew) D.dokumen.push(r);
    save(); UI.focus = 'd-' + r.id; render(); toast('Dokumen tersimpan');
  });
}
function dokStatus(x, to) {
  const apply = issue => {
    if (to === 'Submitted') { if (x.status === 'Revisi' || x.status === 'Rejected') x.rev++; x.tglSubmit = tgl(); x.tglApprove = null; }
    if (to === 'Approved') x.tglApprove = tgl();
    if (issue != null) x.issue = issue;
    x.status = to; riwayat(x, `${to} (Rev ${x.rev})${issue ? ': ' + issue : ''}`);
    save(); render(); toast(`${x.nomor || x.nama} → ${to}`);
  };
  if (to !== 'Revisi' && to !== 'Rejected') return apply(null);
  openDlg(`<h3>${to === 'Revisi' ? 'Revisi' : 'Reject'} · ${esc(x.nomor || x.nama)}</h3>
    <label class="f">Issue / komentar reviewer<textarea name="issue" required>${esc(x.issue)}</textarea></label>
    <div class="dlg-act"><button value="cancel" class="btn" formnovalidate>Batal</button><button value="save" class="btn pri">Simpan</button></div>`, fd => apply(fd.get('issue')));
}
/* dokumen task belum approve → kendala Engineering task tsb */
function dokKendala(t) {
  const open = dokOfTask(t.kode).filter(x => x.status !== 'Approved');
  if (!open.length) return;
  const n = open.map(x => x.nama).join(' ').toLowerCase(), E = D.issues.eng || [];
  const pick = [[/method/, 'Method statement'], [/submittal/, 'Material submittal'], [/rfi/, 'RFI'], [/izin|ptw/, 'Izin kerja'], [/./, 'Shop drawing belum approve']]
    .map(([re, key]) => re.test(n) && E.find(i => i.startsWith(key))).find(Boolean) || '';
  t.k.eng = { st: 'Pending', issue: pick, detail: open.map(x => `${x.nomor || x.nama} (${x.status}, Rev ${x.rev})`).join('; '), pic: open[0].pic || t.k.eng.pic,
    target: open.map(x => x.target).filter(Boolean).sort().at(-1) || t.k.eng.target };
  save(); render(); toast(`${t.kode}: kendala Engineering → Pending (${open.length} dokumen)`);
}
const dokLine = t => {
  const ds = dokOfTask(t.kode); if (!ds.length) return '';
  const open = ds.filter(x => x.status !== 'Approved');
  return `<div class="meta" style="margin-top:6px"><span>Dokumen:</span>${ds.map(x => `<a href="#sd" data-act="goto-dok" data-id="${esc(x.id)}" class="chip ${x.status === 'Approved' ? 'ok' : lewatDok(x) ? 'bad' : 'warn'}" style="text-decoration:none">${esc(x.nomor || x.nama)} · ${esc(x.status)}</a>`).join('')}
    ${open.length && t.k.eng.st !== 'Pending' ? `<button class="btn sm" data-act="dok-kendala" data-kode="${esc(t.kode)}">⚠ ${open.length} belum approve → jadikan kendala Engineering</button>` : ''}</div>`;
};
const vh = (title, r) => `<div class="vh"><h2>${esc(title)}</h2><button class="btn sm" data-act="pdf" data-r="${r}">PDF per PIC</button></div>`;
const picOf = s => String(s || '').trim() || '(tanpa PIC)';
const prioTxt = t => [t.prio ? t.prio.toUpperCase() : '', koordList(t).length ? '↔ ' + koordList(t).join(', ') : '', t.followup].filter(Boolean).join('\n');
const rentang = t => t.mulai || t.selesai ? `${fmt(t.mulai)} – ${fmt(t.selesai)}` : '–';
/* Tiap laporan: judul, kolom, dan baris {pic, cells}. Make Ready / Komitmen / Harian = PIC task; Issue = PIC kendala. */
const REPORTS = {
  komit: { title: 'Komitmen Mingguan', week: true,
    head: ['Kode', 'Prioritas · koordinasi', 'Pekerjaan', 'Uraian', 'Lokasi / Area', 'Rencana', 'Disepakati', 'Alasan / catatan rapat'],
    rows: o => D.tasks.filter(inW(o.week)).filter(prioOK).sort((a, b) => byPrio(a, b) || a.kode.localeCompare(b.kode)).map(t => ({ pic: picOf(t.pic),
      cells: [t.kode, prioTxt(t), t.pekerjaan, t.uraian, [t.lokasi, t.area].filter(Boolean).join(' · '), rentang(t), t.disepakati || '–', t.alasan] })) },
  ready: { title: 'Make Ready – Cek Kendala', week: true,
    head: ['Kode', 'Prioritas · koordinasi', 'Uraian', 'Lokasi', 'Rencana', 'Status kendala', 'Kendala pending (issue · PIC · target)', 'Progress', 'Pelaksanaan'],
    rows: o => D.tasks.filter(t => t.disepakati === 'Ya' && inW(o.week)(t) && prioOK(t)).sort((a, b) => byPrio(a, b) || a.kode.localeCompare(b.kode)).map(t => ({ pic: picOf(t.pic),
      cells: [t.kode, prioTxt(t), t.uraian, t.lokasi, rentang(t), stK(t),
        KK.filter(k => t.k[k].st === 'Pending').map(k => { const c = t.k[k]; return `${KAT[k]}: ${c.issue || '?'}${c.detail ? ' – ' + c.detail : ''} · ${c.pic || 'PIC?'} · ${fmt(c.target)}`; }).join('\n') || 'Semua OK',
        pc(prog(t.kode)), stP(t)] })) },
  harian: { title: 'Daily Commitment', date: true,
    head: ['Kode', 'Uraian / lokasi', 'Target hari ini', 'Bobot', 'Actual', '% tercapai', 'Tercapai?', 'Progress kumulatif', 'Penyebab tidak tercapai', 'Catatan'],
    rows: o => D.daily.filter(x => x.tgl === o.tgl).map(x => { const t = D.tasks.find(t => t.kode === x.kode); return (!UI.fPrio || prioOK(t)) && { pic: picOf(t?.pic),
      cells: [x.kode + (t?.prio ? '\n' + t.prio.toUpperCase() : ''), [t?.uraian, t?.lokasi].filter(Boolean).join(' · '), x.target, pc(x.bobot), x.capaian == null ? '' : pc(aktual(x)), x.capaian == null ? '' : pc(x.capaian), tercapai(x),
        x.capaian == null ? '' : pc(progKum(x)), x.penyebab, x.catatan] }; }) },
  issue: { title: 'Follow Up Issue / Kendala', week: true,
    head: ['Kode', 'Task', 'Kategori', 'Issue', 'Detail', 'Target', 'Sisa hari', 'Status'],
    rows: o => issueList(o.week).filter(x => prioOK(x.t)).sort((a, b) => byPrio(a.t, b.t)).map(x => ({ pic: picOf(x.c.pic || x.t.pic),
      cells: [x.t.kode + (x.t.prio ? '\n' + x.t.prio.toUpperCase() : ''), [x.t.uraian, x.t.lokasi].filter(Boolean).join(' · '), KAT[x.k], x.c.issue, x.c.detail, fmt(x.c.target), x.sisa ?? '–', x.st] })) },
};
function pdfOpts(r) { return { week: UI.week, tgl: UI.harianTgl || tgl(), pic: '' , r }; }
function pdfForm(r) {
  const R = REPORTS[r], o = pdfOpts(r);
  const pics = [...new Set(R.rows(o).filter(Boolean).map(x => x.pic))].sort();
  return `<h3>PDF ${esc(R.title)}</h3>
  ${R.week ? `<label class="f">Minggu<select name="week">${opt('', o.week, 'Semua minggu')}${D.meta.weeks.map(x => opt(x.w, o.week, `W${x.w} · ${x.label}`)).join('')}</select></label>` : ''}
  ${R.date ? `<label class="f">Tanggal<input type="date" name="tgl" value="${o.tgl}"></label>` : ''}
  <label class="f">PIC<select name="pic">${opt('', '', 'Semua PIC – 1 halaman per PIC')}${pics.map(p => opt(p, '')).join('')}</select></label>
  <p class="muted" style="font-size:13px">Daftar PIC mengikuti ${R.date ? 'tanggal' : 'minggu'} yang sedang dibuka. Di jendela cetak pilih <b>Simpan sebagai PDF</b>, lalu kirim ke PIC.${UI.fPrio ? ` Filter prioritas aktif: <b>${esc(FPRIO[UI.fPrio])}</b>.` : ''}${r === 'issue' ? ' Issue dikelompokkan per PIC kendala.' : ''}</p>
  <div class="dlg-act"><button value="cancel" class="btn" formnovalidate>Batal</button><button value="save" class="btn pri">Buat PDF</button></div>`;
}
function printReport(r, o) {
  const R = REPORTS[r], rows = R.rows(o).filter(x => x && (!o.pic || x.pic === o.pic));
  if (!rows.length) { toast('Tidak ada data untuk dicetak'); return; }
  const groups = {};
  rows.forEach(x => (groups[x.pic] ||= []).push(x.cells));
  const wo = o.week === '' ? null : weekObj(o.week);
  const per = R.date ? fmt(o.tgl) + ` (W${weekOf(o.tgl)})` : wo ? `W${wo.w} · ${fmt(wo.start)} – ${fmt(wo.end)}` : 'Semua minggu';
  const cell = v => esc(v ?? '').replace(/\n/g, '<br>');
  $('#print').innerHTML = Object.keys(groups).sort().map(pic => `<section class="pg">
    <header><div><b>${esc(R.title)}</b><span>${esc(D.meta.project)}</span></div>
      <div class="pr-meta"><span>PIC: <b>${esc(pic)}</b></span><span>${esc(per)}</span><span>Dicetak ${fmt(todayISO())}${D.sync.user ? ' oleh ' + esc(D.sync.user) : ''}</span></div></header>
    <table><thead><tr><th>No</th>${R.head.map(h => `<th>${esc(h)}</th>`).join('')}</tr></thead>
      <tbody>${groups[pic].map((c, i) => `<tr><td>${i + 1}</td>${c.map(v => `<td>${cell(v)}</td>`).join('')}</tr>`).join('')}</tbody></table>
    <p class="pr-n">${groups[pic].length} baris</p>
    <footer><div>Dibuat oleh<br><br><br>( ${esc(D.sync.user || '..................')} )</div><div>Diterima PIC<br><br><br>( ${esc(pic)} )</div></footer>
  </section>`).join('');
  const judul = document.title;
  document.title = [R.title.split(' ')[0], o.pic || 'Semua-PIC', R.date ? o.tgl : wo ? 'W' + wo.w : 'semua'].join('_').replace(/[\\/:*?"<>|\s]+/g, '-'); // nama file PDF
  addEventListener('afterprint', () => { document.title = judul; $('#print').innerHTML = ''; }, { once: true });
  setTimeout(() => print(), 60); // beri waktu dialog tertutup
}

/* ---------- PDF Jadwal / Lookahead: 1 minggu atau beberapa minggu ke depan, bar per minggu, opsional task komitmen ---------- */
function pdfJadwalForm() {
  const cw = weekOf(tgl()), from = UI.jFrom ?? cw, look = UI.jLook === 'all' ? '4' : UI.jLook;
  return `<h3>PDF Jadwal / Lookahead</h3>
  <div class="frow">
    <label class="f">Mulai minggu<select name="from">${D.meta.weeks.map(x => opt(x.w, from, `W${x.w} · ${x.label}${x.w === cw ? ' (ini)' : ''}`)).join('')}</select></label>
    <label class="f">Rentang<select name="look">${[['1', '1 minggu (weekly)'], ['2', '2 minggu'], ['3', '3 minggu'], ['4', '4 minggu'], ['6', '6 minggu'], ['8', '8 minggu'], ['all', 'Semua jadwal']].map(([v, l]) => opt(v, look, l)).join('')}</select></label>
  </div>
  <div class="frow">
    <label class="f">Disiplin<select name="dis">${opt('', UI.jDis, 'Semua disiplin')}${D.disiplin.map(d => opt(d, UI.jDis)).join('')}</select></label>
    <label class="f">Susunan<select name="grp">${opt('dis', 'dis', '1 halaman per disiplin')}${opt('satu', 'dis', 'Satu tabel')}</select></label>
  </div>
  <label class="f" style="flex-direction:row;align-items:center;gap:8px"><input type="checkbox" name="task" value="1" checked style="width:auto;min-height:0"> Sertakan task komitmen (kode · uraian · PIC) di bawah tiap pekerjaan</label>
  <p class="muted" style="font-size:13px">Kolom W… menandai minggu aktif tiap pekerjaan (maks. 8 minggu; "Semua jadwal" tanpa kolom minggu). Di jendela cetak pilih <b>Simpan sebagai PDF</b>.</p>
  <div class="dlg-act"><button value="cancel" class="btn" formnovalidate>Batal</button><button value="save" class="btn pri">Buat PDF</button></div>`;
}
function printJadwal(o) {
  const { wF, wT, list } = jadwalList(o.from, o.look, o.dis);
  if (!list.length) { toast('Tidak ada pekerjaan di rentang ini'); return; }
  const weeks = o.look === 'all' ? [] : D.meta.weeks.filter(x => x.w >= wF.w && x.w <= wT.w);
  const gantt = weeks.length && weeks.length <= 8 ? weeks : [];
  const aktifW = (a, b, w) => !!(a && b && a <= w.end && b >= w.start);
  const per = o.look === 'all' ? 'Semua jadwal' : `W${wF.w}${wT.w !== wF.w ? '–W' + wT.w : ''} · ${fmt(wF.start)} – ${fmt(wT.end)}`;
  const head = ['No', 'Kode', 'Pekerjaan / task', 'Area / lokasi', 'Mulai', 'Selesai', 'Bobot', 'Status · PIC'];
  const cw = weekOf(tgl());
  const gCells = (a, b, cls) => gantt.map(w => `<td class="g${aktifW(a, b, w) ? ' on' + cls : ''}${w.w === cw ? ' now' : ''}"></td>`).join('');
  const rowsOf = ps => { let n = 0; return ps.map(p => {
    const ts = o.task ? D.tasks.filter(t => t.pekerjaan === p.nama && (o.look === 'all' || (+t.minggu >= wF.w && +t.minggu <= wT.w))).sort((a, b) => a.minggu - b.minggu || a.kode.localeCompare(b.kode)) : [];
    return `<tr class="pj"><td>${++n}</td><td>${esc(p.kode)}</td><td><b>${esc(p.nama.replace(/^\S+\s/, ''))}</b><br><small>${esc(p.bill)}</small></td><td>${esc(p.area || '')}</td>
      <td>${fmt(p.mulai)}</td><td>${fmt(p.selesai)}</td><td>${p.bobot.toFixed(2)}%</td><td>${esc(stJadwal(p))}</td>${gCells(p.mulai, p.selesai, '')}</tr>`
      + ts.map(t => `<tr class="tk"><td></td><td>${esc(t.kode)}</td><td>↳ ${esc(t.uraian)}${t.prio ? ` <b>[${esc(t.prio.toUpperCase())}]</b>` : ''}</td><td>${esc(t.lokasi || '')}</td>
      <td>${fmt(t.mulai)}</td><td>${fmt(t.selesai)}</td><td>W${t.minggu}</td><td>${esc([t.disepakati === 'Ya' ? stP(t) : 'Belum sepakat', t.pic].filter(Boolean).join(' · '))}</td>
      ${gCells(t.mulai || weekObj(t.minggu)?.start, t.selesai || weekObj(t.minggu)?.end, ' tk')}</tr>`).join('');
  }).join(''); };
  const groups = o.grp === 'dis' ? [...new Set(list.map(p => p.disiplin))].map(d => [d, list.filter(p => p.disiplin === d)]) : [[o.dis || 'Semua disiplin', list]];
  $('#print').innerHTML = groups.map(([g, ps]) => `<section class="pg jdw">
    <header><div><b>Jadwal / Lookahead${o.look === '1' ? ' Mingguan' : o.look === 'all' ? '' : ' ' + o.look + ' Minggu'}</b><span>${esc(D.meta.project)}</span></div>
      <div class="pr-meta"><span>Disiplin: <b>${esc(g)}</b></span><span>${esc(per)}</span><span>Dicetak ${fmt(todayISO())}${D.sync.user ? ' oleh ' + esc(D.sync.user) : ''}</span></div></header>
    <table><thead><tr>${head.map(h => `<th>${esc(h)}</th>`).join('')}${gantt.map(w => `<th class="g${w.w === cw ? ' now' : ''}">W${w.w}<br><small>${fmt(w.start).slice(0, 6)}</small></th>`).join('')}</tr></thead>
      <tbody>${rowsOf(ps)}</tbody></table>
    <p class="pr-n">${ps.length} pekerjaan · bobot ${ps.reduce((s, p) => s + p.bobot, 0).toFixed(2)}%${gantt.length ? ' · kotak gelap = minggu aktif pekerjaan, kotak muda = minggu task, garis merah = minggu ini' : ''}</p>
  </section>`).join('');
  const judul = document.title;
  document.title = ['Jadwal', o.look === 'all' ? 'semua' : `W${wF.w}${wT.w !== wF.w ? '-W' + wT.w : ''}`, o.dis || 'Semua'].join('_');
  addEventListener('afterprint', () => { document.title = judul; $('#print').innerHTML = ''; }, { once: true });
  setTimeout(() => print(), 60);
}

/* ---------- dialog ---------- */
let dlgSubmit = null;
function openDlg(html, onSave) {
  $('#dlg-form').innerHTML = html;
  dlgSubmit = onSave;
  if (!$('#dlg').open) $('#dlg').showModal();
}
$('#dlg').addEventListener('close', () => {
  const dlg = $('#dlg');
  if (dlg.returnValue === 'save' && dlgSubmit) dlgSubmit(new FormData($('#dlg-form')));
  dlg.returnValue = ''; dlgSubmit = null;
});

function editTask(t, isNew) {
  openDlg(taskForm(t, isNew), fd => {
    const v = Object.fromEntries(fd.entries());
    v.minggu = +v.minggu;
    v.koord = fd.getAll('koord').join(','); // checkbox: bisa lebih dari satu
    if (isNew && D.tasks.some(x => x.kode === v.kode)) { toast('Kode ' + v.kode + ' sudah dipakai'); return; }
    const p = D.pekerjaan.find(x => x.nama === v.pekerjaan);
    if (p && !v.disiplin) v.disiplin = p.disiplin;
    if (isNew) D.tasks.push(Object.assign(t, v));
    else Object.assign(D.tasks.find(x => x.kode === t.kode), v);
    save(); UI.focus = 'k-' + v.kode; render(); toast('Task tersimpan');
  });
}

function newTask(pek) {
  const p = D.pekerjaan.find(x => x.nama === pek);
  const w = UI.week !== '' ? +UI.week : weekOf(tgl());
  const wo = weekObj(w);
  editTask({
    kode: nextKode(), minggu: w, disiplin: p?.disiplin === 'PRELIM' ? 'CSA' : (p?.disiplin || 'CSA'), pekerjaan: pek || '',
    area: D.areas.includes(p?.area) ? p.area : '', uraian: '', lokasi: '', pic: '',
    mulai: wo?.start || null, selesai: wo?.end || null, disepakati: '', alasan: '',
    inLPS: false, k: blankK(), catatanKendala: '', penyebabTerlambat: '', prio: '', koord: '', followup: ''
  }, true);
}

/* Satu task komitmen → beberapa janji harian (beberapa tanggal, atau beberapa janji di hari yang sama) */
const janjiRow = (d, target = '', bobot = '') => `<div class="jr">
  <input type="date" name="jtgl" value="${d}" aria-label="Tanggal"><input type="number" name="jbobot" min="0" max="100" step="any" inputmode="decimal" placeholder="Bobot %" value="${bobot}" aria-label="Bobot %">
  <button type="button" class="btn sm" data-act="jr-del" aria-label="Hapus baris">✕</button>
  <input type="text" name="jtarget" placeholder="Target (lokasi / volume)" value="${esc(target)}" aria-label="Target"></div>`;
function jrInfo(kode) {
  const t = D.tasks.find(x => x.kode === kode); if (!t) return '';
  return `Rencana ${fmt(t.mulai)} – ${fmt(t.selesai)} · progress ${pc(prog(kode) || 0)} · janji belum dilaporkan ${pc(belumLapor(kode))} · <b>sisa untuk direncanakan ${pc(sisaRencana(kode))}</b>`;
}
let jrTgl = null; // tanggal default baris janji di dialog
function newDaily(kode, tglJanji) {
  const d = jrTgl = tglJanji || UI.harianTgl || tgl();
  const ready = D.tasks.filter(t => t.disepakati === 'Ya' && stK(t) === 'READY' && stP(t) !== 'FINISH');
  if (!ready.length) { toast('Belum ada task READY'); return; }
  const k0 = ready.some(t => t.kode === kode) ? kode : ready[0].kode;
  openDlg(`<h3>Janji harian</h3>
    <div class="f">Task (hanya READY & belum FINISH)
      ${ready.length > 5 ? '<input type="search" id="jr-cari" placeholder="Cari kode / uraian / lokasi / PIC…" style="margin:3px 0 6px">' : ''}
      <div class="pick" id="jr-pick" role="radiogroup">${ready.map(t => `<label class="pick-i" data-cari="${esc((t.kode + ' ' + t.uraian + ' ' + t.lokasi + ' ' + t.pic).toLowerCase())}">
        <input type="radio" name="kode" value="${esc(t.kode)}"${t.kode === k0 ? ' checked' : ''}>
        <span><b>${esc(t.kode)}</b> ${esc(t.uraian)}<small>${esc([t.lokasi, t.pic && 'PIC ' + t.pic].filter(Boolean).join(' · '))}</small></span></label>`).join('')}</div></div>
    <p class="muted" id="jr-info" style="font-size:13px;margin:-4px 0 10px">${jrInfo(k0)}</p>
    <div class="muted" style="font-size:12px;margin-bottom:4px">Tanggal · bobot (% dari task) · target terukur. Boleh beberapa baris di tanggal yang sama.</div>
    <div id="jr-rows">${janjiRow(d)}</div>
    <div class="bar" style="margin:4px 0 0"><button type="button" class="btn sm" data-act="jr-add">+ Baris</button><button type="button" class="btn sm" data-act="jr-auto" data-n="7">7 hari ke depan</button><button type="button" class="btn sm" data-act="jr-auto" data-n="14">Bagi rata s/d rencana selesai</button></div>
    <label class="muted" style="font-size:13px;display:block;margin-top:6px"><input type="checkbox" id="jr-minggu" style="width:auto;min-height:0"> lewati hari Minggu</label>
    <div class="dlg-act"><button value="cancel" class="btn" formnovalidate>Batal</button><button value="save" class="btn pri">Simpan janji</button></div>`,
  fd => {
    const k = fd.get('kode'), tg = fd.getAll('jtgl'), tr = fd.getAll('jtarget'), bb = fd.getAll('jbobot');
    const rows = tg.map((x, i) => ({ tgl: x, target: tr[i].trim(), bobot: nOr(bb[i]) })).filter(r => r.tgl && r.target && r.bobot > 0);
    if (!rows.length) { toast('Isi tanggal, target & bobot minimal satu baris'); return; }
    const sisa = sisaRencana(k), tot = r6(rows.reduce((s, r) => s + r.bobot, 0));
    rows.forEach(r => D.daily.push({ id: newId(), tgl: r.tgl, kode: k, target: r.target, bobot: r.bobot, capaian: null, penyebab: '', catatan: '' }));
    save(); render();
    toast(`${rows.length} janji ${k} disimpan${tot > sisa + 0.01 ? ` · ⚠ total bobot ${pc(tot)} > sisa ${pc(sisa)}` : ''}`);
  });
}
/* Satu baris janji per hari kerja dari tanggal janji (atau rencana mulai) s/d rencana selesai, maks `limit` hari.
   Bobot sisa dibagi rata ke SEMUA hari kerja s/d rencana selesai; yang dibuat hanya `limit` hari pertama
   (mis. 7 hari ke depan), sisanya direncanakan minggu berikutnya. Tiap janji tampil di Harian pada tanggalnya. */
function jrAuto(limit) {
  const t = D.tasks.find(x => x.kode === $('#jr-pick input:checked')?.value);
  const from = [jrTgl || UI.harianTgl || tgl(), t.mulai].filter(Boolean).sort().at(-1);
  const to = t.selesai && t.selesai >= from ? t.selesai : addDays(from, limit - 1), libur = $('#jr-minggu')?.checked;
  const hari = [];
  for (let d = from; d <= to; d = addDays(d, 1)) if (!(libur && new Date(d + 'T00:00:00').getDay() === 0)) hari.push(d);
  const sisa = sisaRencana(t.kode);
  if (!sisa) { toast('Bobot task sudah habis direncanakan'); return; }
  if (!hari.length) { toast('Tidak ada hari kerja di rentang ini'); return; }
  const dibuat = hari.slice(0, limit), b = Math.floor(sisa / hari.length * 10) / 10, habis = dibuat.length === hari.length;
  $('#jr-rows').innerHTML = dibuat.map((d, i) => janjiRow(d, '', habis && i === dibuat.length - 1 ? r6(sisa - b * (hari.length - 1)) : b)).join('');
  toast(habis ? `${dibuat.length} hari s/d rencana selesai ${fmt(t.selesai)}` : `${dibuat.length} hari dibuat · ${hari.length - dibuat.length} hari lagi s/d ${fmt(to)} direncanakan berikutnya`);
}

/* ---------- event ---------- */
function setPath(o, path, v) {
  const ks = path.split('.'); const last = ks.pop();
  ks.reduce((a, k) => a[k], o)[last] = v;
}

document.addEventListener('click', e => {
  const el = e.target.closest('[data-act]');
  if (!el) return;
  const a = el.dataset.act, kode = el.dataset.kode, task = kode && D.tasks.find(t => t.kode === kode);
  // tombol di dialog detail dokumen: tutup detail dulu (dialog berikutnya memakai #dlg yang sama)
  if (['dok-st', 'dok-edit', 'goto-ready'].includes(a) && el.closest('#dlg')?.open) $('#dlg').close();
  switch (a) {
    case 'settings':
      openDlg(settingsForm(), fd => {
        D.settings.tglMode = fd.get('tglMode');
        D.settings.tglFixed = fd.get('tglFixed') || null;
        D.meta.targetPPC = Math.max(0, Math.min(100, +fd.get('targetPPC') || 80)) / 100;
        const S = D.sync, key = fd.get('syncKey').trim(), user = fd.get('syncUser').trim();
        const url = fd.get('syncUrl').replace(/[\s\u200B-\u200D\u2060\uFEFF]/g, ''); // spasi & karakter tak terlihat dari copy-paste
        if (url && (!/^https?:\/\//.test(url) || !key || !user)) { toast('Isi URL (https://…), kode akses & nama untuk sinkron'); return; }
        if (url !== S.url || key !== S.key) { S.rev = 0; S.masterVer = ''; S.error = ''; } // sambungan baru: tarik data tim dulu
        Object.assign(S, { url, key, user });
        save(); render(); toast('Pengaturan disimpan');
        if (syncOn()) syncNow(true);
      });
      break;
    case 'aktual': openDlg(aktualForm(), simpanAktual); break;
    case 'sd-tab': UI.sd = el.dataset.v; render(); break;
    case 'mp-day': UI.mpTgl = addDays(UI.mpTgl || tgl(), +el.dataset.n); render(); break;
    case 'mp-all': UI.mpAll = el.checked; render(); break;
    case 'mp-copy': {
      const d = UI.mpTgl || tgl(), prev = [...new Set(D.manpower.filter(x => x.tgl < d && x.aktual != null).map(x => x.tgl))].sort().at(-1);
      if (!prev) { toast('Belum ada data hari sebelumnya'); break; }
      let n = 0;
      D.manpower.filter(x => x.tgl === prev && x.aktual != null).forEach(x => { if (mpAkt(d, x.item) == null) { D.manpower.push({ id: mpId(d, x.item), tgl: d, item: x.item, aktual: x.aktual, catatan: '' }); n++; } });
      save(); render(); toast(`${n} aktivitas disalin dari ${fmt(prev)} – sesuaikan bila berubah`);
      break;
    }
    case 'dok-f': UI.dokF = el.dataset.v; render(); break;
    case 'drv-load': drvLoad(true); break;
    case 'dok-kat': UI.dokJenis = el.dataset.v; UI.dokSub = ''; UI.dokTutup = []; render(); break;
    case 'dok-grp': { const k = el.dataset.v; UI.dokTutup = UI.dokTutup.includes(k) ? UI.dokTutup.filter(x => x !== k) : [...UI.dokTutup, k]; render(); break; }
    case 'drv-set': drvSet(); break;
    case 'dok-new': editDok(normDok({ scope: D.disiplin[0], status: 'Draft', rev: 0 }), true); break;
    case 'dok-info': dokInfo(el.dataset.id); break;
    case 'dok-edit': editDok({ ...D.dokumen.find(x => x.id === el.dataset.id) }, false); break;
    case 'dok-del':
      if (confirm('Hapus dokumen ini?')) { D.dokumen = D.dokumen.filter(x => x.id !== el.dataset.id); save(); $('#dlg').close(); render(); }
      break;
    case 'dok-st': dokStatus(D.dokumen.find(x => x.id === el.dataset.id), el.dataset.v); break;
    case 'dok-kendala': dokKendala(task); break;
    case 'goto-dok': UI.sd = 'dok'; UI.dokF = ''; UI.focus = 'd-' + el.dataset.id; if (UI.view === 'sd') render(); dokInfo(el.dataset.id); break;
    case 'fprio': UI.fPrio = el.dataset.v; if (el.tagName === 'A') break; render(); break; // <a> → hashchange render
    case 'pdf': {
      const r = el.dataset.r;
      openDlg(pdfForm(r), fd => printReport(r, { week: fd.has('week') ? fd.get('week') : UI.week, tgl: fd.get('tgl') || UI.harianTgl || tgl(), pic: fd.get('pic') || '' }));
      break;
    }
    case 'pdf-jadwal': openDlg(pdfJadwalForm(), fd => printJadwal({ from: fd.get('from'), look: fd.get('look'), dis: fd.get('dis'), grp: fd.get('grp'), task: fd.has('task') })); break;
    case 'tl-pick': tlPick(el.dataset.w || tlWeek()); break;
    case 'tl-ubah': e.preventDefault(); tlUbah(el.dataset.id); break;
    case 'tl-reset': D.tlEdit = D.tlEdit.filter(x => x.id !== el.dataset.id); dlgSubmit = null; $('#dlg').close(); save(); render(); toast('Tanggal kembali sesuai PDF'); break;
    case 'tl-add': { const t = tlBuat([el.dataset.id], +el.dataset.w)[0]; if (t) { render(); editTask({ ...t }, false); } break; }
    case 'tl-all': document.querySelectorAll('#tl-pick .pick-i:not([hidden]) input:not(:disabled)').forEach(c => { c.checked = true; }); tlFilter(); break;
    case 'tl-none': document.querySelectorAll('#tl-pick input:checked').forEach(c => { c.checked = false; }); tlFilter(); break;
    case 'tl-manual': { const w = $('#tl-w')?.value; if (w != null) UI.week = w; newTask(); break; } // isi dialog diganti form task biasa
    case 'tl-goto': UI.week = el.dataset.w; if (UI.view === 'komit') render(); break;
    case 'task-new': newTask(el.dataset.pek); break;
    case 'task-edit': editTask({ ...task }, false); break;
    case 'task-del':
      if (confirm(`Hapus task ${kode}? Daily log task ini juga dihapus.`)) {
        D.tasks = D.tasks.filter(t => t.kode !== kode);
        D.daily = D.daily.filter(x => x.kode !== kode);
        save(); $('#dlg').close(); render(); toast('Task dihapus');
      }
      break;
    case 'sepakat': task.disepakati = task.disepakati === el.dataset.v ? '' : el.dataset.v; task.inLPS = task.disepakati === 'Ya'; save(); render(); break;
    case 'kst': task.k[el.dataset.k].st = el.dataset.v; save(); render(); break;
    case 'daily-new': newDaily(el.dataset.kode, el.dataset.tgl); break;
    case 'hmode': D.settings.hMode = el.dataset.v; persist(); render(); break; // per perangkat, tidak disinkron
    case 'jr-add': {
      const last = [...document.querySelectorAll('#jr-rows [name=jtgl]')].at(-1)?.value || jrTgl || tgl();
      $('#jr-rows').insertAdjacentHTML('beforeend', janjiRow(addDays(last, 1)));
      break;
    }
    case 'jr-del': if (document.querySelectorAll('#jr-rows .jr').length > 1) el.closest('.jr').remove(); break;
    case 'jr-auto': jrAuto(+el.dataset.n || 14); break;
    case 'daily-del': D.daily = D.daily.filter(r => r.id !== el.dataset.id); save(); render(); break;
    case 'daily-copy': {
      const d = el.dataset.tgl || UI.harianTgl || tgl();
      const prev = D.daily.some(x => x.tgl === el.dataset.from) ? el.dataset.from : [...new Set(D.daily.filter(x => x.tgl < d).map(x => x.tgl))].sort().at(-1);
      if (!prev) { toast('Tidak ada log sebelumnya'); break; }
      let n = 0;
      D.daily.filter(x => x.tgl === prev).forEach(x => {
        const t = D.tasks.find(t => t.kode === x.kode);
        if (!t || stP(t) === 'FINISH' || D.daily.some(y => y.tgl === d && y.kode === x.kode)) return;
        D.daily.push({ id: newId(), tgl: d, kode: x.kode, target: '', bobot: null, capaian: null, penyebab: '', catatan: '' }); n++;
      });
      save(); render(); toast(`${n} janji disalin dari ${fmt(prev)} ke ${fmt(d)} – isi target & bobot`);
      break;
    }
    case 'day': UI.harianTgl = addDays(UI.harianTgl || tgl(), +el.dataset.n); render(); break;
    case 'goto-ready': UI.focus = 'r-' + kode; if (UI.view === 'ready') render(); break;
    case 'goto-komit': UI.focus = 'k-' + kode; if (UI.view === 'komit') render(); break;
    case 'exp-json': download(`LPS-Batik_backup_${tgl()}.json`, JSON.stringify({ ...D, sync: undefined }, null, 1), 'application/json'); break; // tanpa kode akses
    case 'sync-now': syncNow(true); break;
    case 'sync-off':
      if (confirm('Putuskan sinkron tim? Data tetap ada di perangkat ini.')) { Object.assign(D.sync, { url: '', key: '', rev: 0, masterVer: '', error: '' }); persist(); $('#dlg').close(); render(); }
      break;
    case 'imp-json': {
      const inp = document.createElement('input'); inp.type = 'file'; inp.accept = '.json,application/json';
      inp.onchange = async () => {
        try {
          const s = JSON.parse(await inp.files[0].text());
          if (!s.meta?.weeks?.length || !Array.isArray(s.tasks)) throw new Error('format tidak dikenal');
          const sync = D.sync; D = normalize(s); D.sync = sync; // impor = perubahan lokal, ikut terkirim ke tim
          save(); $('#dlg').close(); render(); toast('Data diimpor');
        } catch (err) { toast('Gagal impor: ' + err.message); }
      };
      inp.click();
      break;
    }
    case 'exp-csv': exportCSV(el.dataset.what); break;
    case 'reseed-ref':
      if (syncOn()) { D.sync.masterVer = ''; persist(); $('#dlg').close(); syncNow(true); break; } // jadwal dari Sheet tim
      fetchSeed().then(s => {
        applyMaster(s); // jadwal & referensi dari Excel; task & daily log lapangan dipertahankan
        const known = new Set(D.tasks.map(t => t.kode));
        const add = s.tasks.filter(t => !known.has(t.kode));
        D.tasks.push(...normalize({ tasks: add }).tasks);
        save(); $('#dlg').close(); render(); toast(`Jadwal diperbarui · ${add.length} task baru dari Excel`);
      }).catch(err => toast(err.message));
      break;
    case 'reseed-all':
      if (confirm(syncOn() ? 'Data di perangkat ini dibuang lalu diambil ulang dari Excel + data tim. Lanjut?' : 'Semua data di perangkat ini diganti dengan data Excel (seed.json). Lanjut?'))
        fetchSeed().catch(err => { if (syncOn()) return blank(); throw err; }).then(s => {
          const sync = D.sync; D = normalize(s);
          D.sync = { ...sync, rev: 0, masterVer: '', h: {}, dirty: {}, tomb: {}, error: '' }; // rev 0 → data tim menang saat sinkron
          save(); $('#dlg').close(); render(); toast('Data direset'); syncNow();
        }).catch(err => toast(err.message));
      break;
  }
});

document.addEventListener('change', e => {
  const el = e.target;
  if (el.dataset.ui) { UI[el.dataset.ui] = el.value; render(); return; }
  if (el.dataset.mp) { // aktual manpower per aktivitas per tanggal; kosong = hapus
    const d = UI.mpTgl || tgl(), id = mpId(d, el.dataset.mp), v = el.value === '' ? null : Math.max(0, Math.round(+el.value));
    const r = D.manpower.find(x => x.id === id);
    if (v == null) D.manpower = D.manpower.filter(x => x !== r);
    else if (r) r.aktual = v; else D.manpower.push({ id, tgl: d, item: el.dataset.mp, aktual: v, catatan: '' });
    save(); render(); return;
  }
  if (el.id === 'tl-w') { tlPick(el.value); return; }
  if (el.closest?.('#tl-pick') || el.id === 'tl-area' || el.id === 'tl-dis') { tlFilter(); return; }
  if (el.name === 'kode' && el.closest('#jr-pick')) { $('#jr-info').innerHTML = jrInfo(el.value); return; }
  if (!el.dataset.f) return;
  const [kind, id, path] = el.dataset.f.split('|');
  let v = el.value;
  if (el.type === 'date') v = v || null;
  if (el.type === 'number') v = v === '' ? null : Math.max(0, Math.min(+el.max || 100, +v));
  if (kind === 'task') { const t = D.tasks.find(x => x.kode === id); setPath(t, path, v); }
  else if (kind === 'daily') {
    const x = D.daily.find(x => x.id === id);
    if (path === 'aktual') x.capaian = v == null || !x.bobot ? null : r6(v / x.bobot * 100); // actual → % tercapai
    else if (path === 'bobot') { const a = aktual(x); x.bobot = v; if (a != null) x.capaian = v ? r6(a / v * 100) : null; } // actual tetap
    else setPath(x, path, v);
  }
  save(); render();
});
// Pencarian langsung saat mengetik
document.addEventListener('input', e => {
  if (e.target.id === 'tl-cari') { tlFilter(); return; }
  if (e.target.id === 'jr-cari') { // saring daftar task di dialog janji
    const q = e.target.value.toLowerCase().trim();
    document.querySelectorAll('#jr-pick .pick-i').forEach(l => { l.hidden = !!q && !l.dataset.cari.includes(q); });
    return;
  }
  const k = e.target.dataset.ui;
  if (k !== 'jSearch' && k !== 'dokQ') return;
  UI[k] = e.target.value;
  const pos = e.target.selectionStart;
  render();
  const s = document.querySelector(`[data-ui="${k}"]`); s.focus(); s.setSelectionRange(pos, pos);
});
$('#f-week').addEventListener('change', e => { UI.week = e.target.value; render(); });

function route() {
  const v = location.hash.slice(1);
  UI.view = VIEWS.includes(v) ? v : 'dash';
  const keep = UI.focus;
  render();
  if (!keep) window.scrollTo(0, 0);
}
window.addEventListener('hashchange', route);

(async () => {
  try { await load(); } catch (err) { $('#view').innerHTML = `<div class="empty">Gagal memuat data: ${esc(err.message)}<br>Jalankan aplikasi lewat server (http://), bukan file://</div>`; return; }
  route();
  syncNow();
  setInterval(() => { if (document.visibilityState === 'visible') syncNow(); }, 60000);
  addEventListener('online', () => syncNow());
  addEventListener('offline', renderSync);
  document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible') syncNow(); });
  if ('serviceWorker' in navigator && location.protocol !== 'file:') navigator.serviceWorker.register('sw.js').catch(() => {});
})();
