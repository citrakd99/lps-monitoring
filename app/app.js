'use strict';
/* LPS Batik – PWA monitoring Last Planner System.
   Data = seed.json (hasil tools/export_excel.py) + perubahan lapangan di localStorage.
   Versi yang di-hosting publik tidak membawa seed.json: jadwal (master) diambil dari Google Sheet tim setelah kode akses diisi.
   Peta Excel → field ada di PETA-EXCEL-KE-APLIKASI.md */

const KEY = 'lps-batik-v1';
const KAT = { eng: 'Engineering / HSE', mat: 'Material & Alat', mp: 'Man Power', oth: 'Other' };
const KK = Object.keys(KAT);
const BLN = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];
const VIEWS = ['dash', 'jadwal', 'komit', 'ready', 'harian', 'issue'];

const $ = s => document.querySelector(s);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const opt = (v, cur, label) => `<option value="${esc(v)}"${String(v) === String(cur ?? '') ? ' selected' : ''}>${esc(label ?? v)}</option>`;
const pct = (x, d = 0) => x == null || isNaN(x) ? '–' : (x * 100).toFixed(d) + '%';

let D = null;
const UI = { view: 'dash', week: '', jSearch: '', jDis: '', jLook: '2', jFrom: null, harianTgl: null, focus: null };

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
function prog(kode) {
  let m = null;
  for (const x of D.daily) if (x.kode === kode && x.progress !== null && x.progress !== '' && !isNaN(x.progress)) m = Math.max(m ?? 0, +x.progress);
  return m;
}
const stK = t => KK.some(k => t.k[k].st === 'Pending') ? 'PENDING' : 'READY';
function stP(t) {
  if (stK(t) !== 'READY') return '';
  if ((prog(t.kode) || 0) >= 100) return 'FINISH';
  if (t.selesai && tgl() > t.selesai) return 'LATE / DELAY';
  return 'ON PROGRESS';
}
const inWeek = t => UI.week === '' || +t.minggu === +UI.week;
const planTasks = () => D.tasks.filter(t => t.disepakati === 'Ya' && inWeek(t));
function stJadwal(p) {
  const d = tgl();
  if (!p.mulai) return 'Tanpa jadwal';
  if (p.mulai > d) return 'Belum mulai';
  if (p.selesai < d) return 'Selesai (jadwal)';
  return 'Berjalan';
}
function issueList() {
  const out = [];
  for (const t of D.tasks) {
    if (t.disepakati !== 'Ya' || !inWeek(t)) continue;
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
  const ya = rows.filter(x => x.tercapai === 'Ya').length;
  return { n: rows.length, ya, tidak: rows.filter(x => x.tercapai === 'Tidak').length, ppc: rows.length ? ya / rows.length : null };
}
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
const MASTER_F = ['meta', 'pekerjaan', 'areas', 'areaMap', 'issues', 'penyebab', 'disiplin', 'source', 'exportedAt'];
const blank = () => ({ meta: { project: '', tglUpdate: null, targetPPC: 0.8, weeks: [], scurve: [] }, pekerjaan: [], areas: [], areaMap: {},
  issues: Object.fromEntries(KK.map(k => [k, []])), penyebab: [], disiplin: ['CSA', 'MEP', 'ARS', 'PRELIM'], tasks: [], daily: [] });
const hasMaster = () => D.meta.weeks.length > 0;
function applyMaster(m) { // jadwal & referensi diganti; task, daily log, target PPC lokal dipertahankan
  const ppc = hasMaster() ? D.meta.targetPPC : m.meta.targetPPC;
  MASTER_F.forEach(f => { if (m[f] !== undefined) D[f] = m[f]; });
  D.meta = { ...m.meta, targetPPC: ppc ?? 0.8 };
}
const TXT_T = ['disiplin', 'pekerjaan', 'area', 'uraian', 'lokasi', 'pic', 'disepakati', 'alasan', 'catatanKendala', 'penyebabTerlambat'];
function normTask(t) {
  t.k = Object.assign(blankK(), t.k || {});
  KK.forEach(k => { const c = t.k[k]; c.st = c.st || 'OK'; ['issue', 'detail', 'pic'].forEach(f => { c[f] = c[f] ?? ''; }); c.target = c.target || null; });
  TXT_T.forEach(f => { t[f] = t[f] ?? ''; });
  t.minggu = +t.minggu || 0; t.mulai = t.mulai || null; t.selesai = t.selesai || null;
  t.inLPS = t.inLPS === true || t.inLPS === 'TRUE';
  return t;
}
function normDaily(x) {
  ['target', 'tercapai', 'penyebab', 'catatan'].forEach(f => { x[f] = x[f] ?? ''; });
  x.progress = x.progress === null || x.progress === undefined || x.progress === '' ? null : +x.progress;
  return x;
}
function normalize(s) {
  s.settings = s.settings || { tglMode: 'auto', tglFixed: null };
  s.sync = Object.assign({ url: '', key: '', user: '', rev: 0, masterVer: '', h: {}, dirty: {}, tomb: {}, lastOk: null, error: '' }, s.sync || {});
  s.tasks.forEach(normTask);
  (s.daily || []).forEach(normDaily);
  return s;
}
function persist() {
  try { localStorage.setItem(KEY, JSON.stringify(D)); } catch (e) { toast('Gagal menyimpan: ' + e.message); }
}
function save() { stamp(); persist(); scheduleSync(); }

/* ---------- sinkron tim (Google Sheets via Apps Script, lihat gas/Code.gs) ----------
   Setiap record (task / daily log) diberi _u (waktu ubah) & _by (nama). Perubahan dideteksi dengan
   membandingkan hash isi record, jadi semua aksi cukup memanggil save(). Konflik: _u terbaru menang. */
const TBL = { tasks: 'kode', daily: 'id' };
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
function applyRemote(res, force) {
  const S = D.sync;
  let n = 0;
  for (const [t, key] of Object.entries(TBL)) for (const r of res[t] || []) {
    const id = t + ':' + r[key], del = r._del === true || r._del === 'TRUE';
    delete r._del; delete r._rev;
    if (!force && S.dirty[id]) { const l = localRec(id); if (l && (l.r._u || '') > (r._u || '')) continue; } // lokal lebih baru → dikirim nanti
    const arr = D[t], i = arr.findIndex(x => String(x[key]) === String(r[key]));
    if (del) { if (i >= 0) { arr.splice(i, 1); n++; } delete S.h[id]; }
    else {
      const rec = t === 'tasks' ? normTask(r) : normDaily(r), h = recHash(rec);
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
    const sent = {}, body = { ...base, since: S.rev, masterVer: S.masterVer, tasks: [], daily: [] };
    for (const id of Object.keys(S.dirty)) { const l = localRec(id); if (l) { body[l.t].push(l.r); sent[id] = l.r._u; } else delete S.dirty[id]; }
    changed += applyRemote(master(await post(body)), false);
    for (const id in sent) { const l = localRec(id); if (!l || l.r._u === sent[id]) delete S.dirty[id]; }
    S.lastOk = new Date().toISOString(); S.error = '';
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
  document.querySelector('.wk').hidden = !hasMaster();
  if (!hasMaster()) {
    $('#hdr-sub').textContent = 'Belum terhubung'; $('#f-week').innerHTML = ''; $('#issue-badge').hidden = true;
    $('#view').innerHTML = vSetup(); renderSync(); return;
  }
  renderHeader();
  $('#view').innerHTML = { dash: vDash, jadwal: vJadwal, komit: vKomit, ready: vReady, harian: vHarian, issue: vIssue }[UI.view]();
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
  const pw = ppcWeek(UI.week), pd = ppcDay(tgl()), tgt = D.meta.targetPPC;
  const iss = issueList(), over = iss.filter(x => x.sisa != null && x.sisa < 0);
  const disc = [...new Set(D.tasks.map(t => t.disiplin).filter(Boolean))];
  const komitAll = D.tasks.filter(inWeek).length;

  return `
  <div class="grid kpis">
    <div class="kpi"><b>${ts.length}</b><span>Task Work Plan (disepakati)</span></div>
    <div class="kpi ok"><b>${ready}</b><span>Task Final · READY</span></div>
    <div class="kpi warn"><b>${pend}</b><span>Pending (make ready)</span></div>
    <div class="kpi"><b>${by('ON PROGRESS')}</b><span>On Progress</span></div>
    <div class="kpi bad"><b>${by('LATE / DELAY')}</b><span>Late / Delay</span></div>
    <div class="kpi info"><b>${by('FINISH')}</b><span>Finish</span></div>
    <div class="kpi ${pw.ppc == null ? '' : pw.ppc >= tgt ? 'ok' : 'bad'}"><b>${pct(pw.ppc)}</b><span>PPC ${UI.week === '' ? 'total' : 'minggu ' + UI.week} (target ${pct(tgt)})</span></div>
    <div class="kpi ${pd.ppc == null ? '' : pd.ppc >= tgt ? 'ok' : 'bad'}"><b>${pct(pd.ppc)}</b><span>PPC hari ini · ${pd.ya}/${pd.n} janji</span></div>
  </div>

  <div class="card">
    <h3>Kurva-S rencana (MASTER)</h3>
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
  S.forEach((s, i) => { g += `<circle cx="${x(i)}" cy="${y(s.cum)}" r="6" fill="transparent" stroke="none"><title>W${s.w}: kumulatif ${s.cum.toFixed(2)}%</title></circle>`; });
  if (now) g += `<line x1="${x(ci)}" x2="${x(ci)}" y1="${T}" y2="${H - B}" stroke="var(--bad)" stroke-dasharray="4 3" stroke-width="1.5"/>
      <circle cx="${x(ci)}" cy="${y(now.cum)}" r="4.5" fill="var(--bad)" stroke="var(--surface)" stroke-width="2"/>
      <text x="${Math.min(x(ci) + 6, W - 120)}" y="${Math.max(y(now.cum) - 8, T + 10)}" style="fill:var(--text);font-weight:600">Rencana s/d W${now.w}: ${now.cum.toFixed(2)}%</text>`;
  return `<svg class="chart" viewBox="0 0 ${W} ${H}" role="img" aria-label="Kurva S rencana">${g}</svg>
    <div class="legend"><span><i style="background:var(--plan)"></i>Kumulatif rencana</span><span><i style="background:var(--brand-2);opacity:.55"></i>Rencana per minggu</span><span><i style="background:var(--bad)"></i>Minggu berjalan</span></div>`;
}

/* ---------- JADWAL (Daftar Pekerjaan + Lookahead) ---------- */
function vJadwal() {
  const cw = weekOf(tgl());
  const from = UI.jFrom ?? cw;
  const look = UI.jLook;
  const wF = weekObj(from), wT = weekObj(Math.min(+from + (+look || 0) - 1, D.meta.weeks.at(-1).w));
  const q = UI.jSearch.toLowerCase();
  const list = D.pekerjaan.filter(p =>
    (!UI.jDis || p.disiplin === UI.jDis) &&
    (!q || (p.nama + ' ' + p.area + ' ' + p.bill).toLowerCase().includes(q)) &&
    (look === 'all' || (p.mulai && wF && wT && p.mulai <= wT.end && p.selesai >= wF.start)));
  return `
  <h2>Jadwal & Lookahead</h2>
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
        <span>Task ${ts.length} · sepakat ${ok.length} · finish ${fin}${ok.length ? ' (' + pct(fin / ok.length) + ')' : ''}</span></div>
      <div style="margin-top:8px"><button class="btn sm" data-act="task-new" data-pek="${esc(p.nama)}">+ Task komitmen</button></div>
    </div>`;
  }).join('') || '<div class="empty">Tidak ada pekerjaan pada filter ini.</div>'}`;
}

/* ---------- KOMITMEN ---------- */
function vKomit() {
  const ts = D.tasks.filter(inWeek).sort((a, b) => a.minggu - b.minggu || a.kode.localeCompare(b.kode));
  let gap = '';
  if (UI.week !== '') {
    const w = weekObj(UI.week);
    const aktif = D.pekerjaan.filter(p => p.mulai && w && p.mulai <= w.end && p.selesai >= w.start);
    const kosong = aktif.filter(p => !D.tasks.some(t => t.pekerjaan === p.nama && +t.minggu === +UI.week));
    gap = `<details class="card"><summary>Pekerjaan MASTER aktif minggu ${UI.week}: ${aktif.length} · <span class="chip warn">${kosong.length} belum ada task</span></summary>
      ${kosong.map(p => `<div style="display:flex;justify-content:space-between;gap:8px;padding:5px 0;border-bottom:1px solid var(--line)"><span><span class="code">${esc(p.kode)}</span> ${esc(p.nama.replace(/^\S+\s/, ''))} <small class="muted">${esc(p.disiplin)}</small></span><button class="btn sm" data-act="task-new" data-pek="${esc(p.nama)}">+ Task</button></div>`).join('')}
    </details>`;
  }
  const n = ts.length, ya = ts.filter(t => t.disepakati === 'Ya').length;
  return `
  <h2>Komitmen mingguan ${UI.week === '' ? '' : '· W' + UI.week}</h2>
  <p class="muted" style="margin-top:-6px">${n} task · ${ya} disepakati${n ? ' (' + pct(ya / n) + ')' : ''}. Pilih <b>Ya</b> saat rapat mingguan untuk task yang sanggup dikerjakan.</p>
  ${gap}
  ${ts.map(t => `
    <div class="item" id="k-${esc(t.kode)}">
      <div class="hd"><div><span class="code">${esc(t.kode)}</span> <span class="ttl">${esc(t.uraian || '(tanpa uraian)')}</span>
        <div class="sub">${esc(t.pekerjaan)}</div></div>
        <div class="chips"><span class="chip">W${t.minggu}</span><span class="chip">${esc(t.disiplin)}</span></div></div>
      <div class="meta"><span>${esc(t.area)}${t.lokasi ? ' · ' + esc(t.lokasi) : ''}</span><span>PIC ${esc(t.pic || '–')}</span><span>${fmt(t.mulai)} → ${fmt(t.selesai)}${t.mulai && t.selesai ? ' (' + (days(t.mulai, t.selesai) + 1) + ' hr)' : ''}</span></div>
      <div class="bar" style="margin:10px 0 0">
        <span class="seg">
          <button type="button" data-act="sepakat" data-kode="${esc(t.kode)}" data-v="Ya" class="${t.disepakati === 'Ya' ? 'on-ok' : ''}">Disepakati</button>
          <button type="button" data-act="sepakat" data-kode="${esc(t.kode)}" data-v="Tidak" class="${t.disepakati === 'Tidak' ? 'on-bad' : ''}">Ditunda</button>
        </span>
        ${t.disepakati === 'Ya' ? stChip(stK(t)) : ''}
        <span style="flex:1"></span>
        <button class="btn sm" data-act="task-edit" data-kode="${esc(t.kode)}">Ubah</button>
      </div>
      ${t.disepakati === 'Tidak' ? `<input type="text" style="margin-top:8px" placeholder="Alasan ditunda / catatan rapat" value="${esc(t.alasan)}" data-f="task|${esc(t.kode)}|alasan" class="${t.alasan ? '' : 'need'}">` : ''}
    </div>`).join('') || '<div class="empty">Belum ada task untuk filter ini. Tekan <b>+ Task</b>.</div>'}
  <button class="btn pri fab" data-act="task-new">+ Task</button>`;
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
  <label class="f">Disepakati<select name="disepakati">${opt('', t.disepakati, 'Belum dibahas')}${opt('Ya', t.disepakati)}${opt('Tidak', t.disepakati)}</select></label>
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
  <h2>Make Ready · cek kendala</h2>
  <div class="grid kpis">
    ${KK.map(k => `<div class="kpi ${cnt(k) ? 'warn' : ''}"><b>${cnt(k)}</b><span>Pending ${KAT[k]}</span></div>`).join('')}
    <div class="kpi ok"><b>${ts.filter(t => stK(t) === 'READY').length}</b><span>Task Ready to Work</span></div>
  </div>
  ${ts.sort((a, b) => (stK(a) === 'PENDING' ? 0 : 1) - (stK(b) === 'PENDING' ? 0 : 1) || a.kode.localeCompare(b.kode)).map(t => {
    const p = prog(t.kode), sp = stP(t);
    return `<div class="item" id="r-${esc(t.kode)}">
      <div class="hd"><div><span class="code">${esc(t.kode)}</span> <span class="ttl">${esc(t.uraian)}</span>
        <div class="sub">${esc(t.pekerjaan)}</div></div>
        <div class="chips">${stChip(stK(t))}${stChip(sp)}</div></div>
      <div class="meta"><span>W${t.minggu} · ${esc(t.disiplin)}</span><span>${esc(t.area)}${t.lokasi ? ' · ' + esc(t.lokasi) : ''}</span><span>PIC ${esc(t.pic || '–')}</span><span>${fmt(t.mulai)} → ${fmt(t.selesai)}</span></div>
      <div style="display:flex;align-items:center;gap:8px;margin:8px 0"><div class="bar-h" style="flex:1"><i style="width:${Math.min(p || 0, 100)}%"></i></div><small class="muted">Progress ${p == null ? '0' : p}% (dari Daily Log)</small></div>
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

/* ---------- HARIAN (Daily Log + PPC Harian + Board 4) ---------- */
function vHarian() {
  const d = UI.harianTgl || tgl();
  const rows = D.daily.filter(x => x.tgl === d);
  const P = ppcDay(d), tgt = D.meta.targetPPC;
  const byKode = Object.fromEntries(D.tasks.map(t => [t.kode, t]));
  const dates = [...new Set(D.daily.map(x => x.tgl))].sort().slice(-14);
  const par = {};
  D.daily.filter(x => x.tercapai === 'Tidak').forEach(x => { const k = x.penyebab || '(belum diisi)'; par[k] = (par[k] || 0) + 1; });
  const parL = Object.entries(par).sort((a, b) => b[1] - a[1]), parT = parL.reduce((s, x) => s + x[1], 0);
  const weeks = [...new Set(D.daily.map(x => weekOf(x.tgl)))].sort((a, b) => a - b);

  const cek = (x, t) => !t ? '⚠ Kode tidak ada' : stK(t) !== 'READY' ? '⚠ Task belum READY' : !x.target ? '⚠ Isi target hari ini'
    : !x.tercapai ? 'Menunggu laporan sore' : x.tercapai === 'Tidak' && !x.penyebab ? '⚠ Isi penyebab' : (x.progress === null || x.progress === '') ? '⚠ Isi progress' : 'OK';

  return `
  <h2>Daily commitment</h2>
  <div class="bar">
    <button class="btn sm" data-act="day" data-n="-1" aria-label="Hari sebelumnya">‹</button>
    <input type="date" value="${d}" data-ui="harianTgl">
    <button class="btn sm" data-act="day" data-n="1" aria-label="Hari berikutnya">›</button>
    <span class="muted">W${weekOf(d)}</span>
  </div>
  <div class="grid kpis">
    <div class="kpi"><b>${P.n}</b><span>Dijanjikan</span></div>
    <div class="kpi ok"><b>${P.ya}</b><span>Tercapai</span></div>
    <div class="kpi bad"><b>${P.tidak}</b><span>Tidak tercapai</span></div>
    <div class="kpi warn"><b>${P.n - P.ya - P.tidak}</b><span>Belum dilaporkan</span></div>
    <div class="kpi ${P.ppc == null ? '' : P.ppc >= tgt ? 'ok' : 'bad'}"><b>${pct(P.ppc)}</b><span>PPC hari ini (target ${pct(tgt)})</span></div>
  </div>
  <div class="bar"><button class="btn pri" data-act="daily-new">+ Janji hari ini</button><button class="btn" data-act="daily-copy">Salin janji kemarin</button></div>
  ${rows.map(x => {
    const t = byKode[x.kode], c = cek(x, t), f = `daily|${x.id}|`;
    return `<div class="item">
      <div class="hd"><div><span class="code">${esc(x.kode)}</span> <span class="ttl">${esc(t ? label(t) : '')}</span>
        <div class="sub">${esc(t?.pekerjaan || '')}</div></div>
        <div class="chips">${stChip(x.tercapai)}<span class="chip ${c === 'OK' ? 'ok' : c.startsWith('⚠') ? 'warn' : ''}">${esc(c)}</span></div></div>
      <label class="f" style="margin-top:8px">Target hari ini (pagi)<input type="text" value="${esc(x.target)}" data-f="${f}target" class="${x.target ? '' : 'need'}"></label>
      <div class="frow">
        <label class="f">Tercapai? (sore)<span class="seg" style="display:flex;margin-top:3px">
          <button type="button" style="flex:1" data-act="capai" data-id="${x.id}" data-v="Ya" class="${x.tercapai === 'Ya' ? 'on-ok' : ''}">Ya</button>
          <button type="button" style="flex:1" data-act="capai" data-id="${x.id}" data-v="Tidak" class="${x.tercapai === 'Tidak' ? 'on-bad' : ''}">Tidak</button></span></label>
        <label class="f">Progress kumulatif (%)<input type="number" min="0" max="100" step="1" inputmode="numeric" value="${x.progress ?? ''}" data-f="${f}progress"></label>
      </div>
      ${x.tercapai === 'Tidak' ? `<label class="f">Penyebab tidak tercapai<select data-f="${f}penyebab" class="${x.penyebab ? '' : 'need'}">${opt('', x.penyebab, '– pilih –')}${D.penyebab.map(p => opt(p, x.penyebab)).join('')}</select></label>` : ''}
      <div class="bar" style="margin:0"><input class="grow" type="text" placeholder="Catatan" value="${esc(x.catatan)}" data-f="${f}catatan"><button class="btn sm danger" data-act="daily-del" data-id="${x.id}">Hapus</button></div>
    </div>`;
  }).join('') || '<div class="empty">Belum ada janji untuk tanggal ini.</div>'}

  <div class="two">
    <div class="card"><h3>PPC harian (14 tanggal terakhir)</h3>${ppcChart(dates)}</div>
    <div class="card">
      <h3>PPC mingguan (Daily Log)</h3>
      <table class="tbl"><tr><th>Minggu</th><th class="n">Dijanjikan</th><th class="n">Tercapai</th><th class="n">PPC</th></tr>
      ${weeks.map(w => { const r = D.daily.filter(x => weekOf(x.tgl) === w), y = r.filter(x => x.tercapai === 'Ya').length; return `<tr><td>W${w}</td><td class="n">${r.length}</td><td class="n">${y}</td><td class="n"><b>${pct(r.length ? y / r.length : null)}</b></td></tr>`; }).join('') || '<tr><td colspan="4" class="muted">–</td></tr>'}
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
  const L = issueList();
  return `
  <h2>Follow up issue / kendala</h2>
  <div class="card scroll"><table class="tbl">
    <tr><th>Kategori</th><th class="n">Pending</th><th class="n">Lewat target</th></tr>
    ${KK.map(k => { const a = L.filter(x => x.k === k); return `<tr><td>${KAT[k]}</td><td class="n">${a.length}</td><td class="n">${a.filter(x => x.sisa != null && x.sisa < 0).length}</td></tr>`; }).join('')}
    <tr><td><b>Total</b></td><td class="n"><b>${L.length}</b></td><td class="n"><b>${L.filter(x => x.sisa != null && x.sisa < 0).length}</b></td></tr>
  </table></div>
  ${L.map(x => `
    <div class="item">
      <div class="hd"><div><span class="code">${esc(x.t.kode)}</span> <span class="ttl">${esc(x.c.issue || '(issue belum dipilih)')}</span>
        <div class="sub">${esc(KAT[x.k])}${x.c.detail ? ' — ' + esc(x.c.detail) : ''}</div></div>
        <div class="chips">${stChip(x.st)}</div></div>
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
      <li><b>Harian</b>: pagi isi janji & target, sore isi tercapai, progress % & penyebab.</li>
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
    h.push('Catatan Kendala', 'Progress (%)', 'Penyebab Terlambat', 'Status Kendala', 'Status Pelaksanaan');
    const rows = D.tasks.map(t => [t.kode, t.minggu, t.disiplin, t.pekerjaan, t.area, t.uraian, t.lokasi, t.pic, t.mulai, t.selesai, t.disepakati, t.alasan,
      ...KK.flatMap(k => [t.k[k].st, t.k[k].issue, t.k[k].detail, t.k[k].pic, t.k[k].target]),
      t.catatanKendala, prog(t.kode), t.penyebabTerlambat, t.disepakati === 'Ya' ? stK(t) : '', t.disepakati === 'Ya' ? stP(t) : '']);
    download(`LPS-Input_${tgl()}.csv`, csv([h, ...rows]), 'text/csv');
  } else {
    const h = ['Tanggal', 'Kode Task', 'Target Hari Ini', 'Tercapai?', 'Progress Kumulatif (%)', 'Penyebab Tidak Tercapai', 'Catatan'];
    const rows = [...D.daily].sort((a, b) => a.tgl.localeCompare(b.tgl)).map(x => [x.tgl, x.kode, x.target, x.tercapai, x.progress, x.penyebab, x.catatan]);
    download(`Daily-Log_${tgl()}.csv`, csv([h, ...rows]), 'text/csv');
  }
}

/* ---------- dialog ---------- */
let dlgSubmit = null;
function openDlg(html, onSave) {
  $('#dlg-form').innerHTML = html;
  dlgSubmit = onSave;
  $('#dlg').showModal();
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
    inLPS: false, k: blankK(), catatanKendala: '', penyebabTerlambat: ''
  }, true);
}

function newDaily() {
  const d = UI.harianTgl || tgl();
  const ready = D.tasks.filter(t => t.disepakati === 'Ya' && stK(t) === 'READY' && stP(t) !== 'FINISH');
  if (!ready.length) { toast('Belum ada task READY'); return; }
  openDlg(`<h3>Janji harian · ${fmt(d)}</h3>
    <label class="f">Task (hanya READY & belum FINISH)<select name="kode" required>${ready.map(t => opt(t.kode, '', `${t.kode} · ${t.uraian}${t.lokasi ? ' · ' + t.lokasi : ''}`)).join('')}</select></label>
    <label class="f">Target hari ini (terukur: lokasi / volume)<input type="text" name="target" required></label>
    <div class="dlg-act"><button value="cancel" class="btn" formnovalidate>Batal</button><button value="save" class="btn pri">Tambah</button></div>`,
  fd => {
    D.daily.push({ id: newId(), tgl: d, kode: fd.get('kode'), target: fd.get('target'), tercapai: '', progress: null, penyebab: '', catatan: '' });
    save(); render();
  });
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
    case 'capai': { const x = D.daily.find(r => r.id === el.dataset.id); x.tercapai = x.tercapai === el.dataset.v ? '' : el.dataset.v; save(); render(); break; }
    case 'daily-new': newDaily(); break;
    case 'daily-del': D.daily = D.daily.filter(r => r.id !== el.dataset.id); save(); render(); break;
    case 'daily-copy': {
      const d = UI.harianTgl || tgl();
      const prev = [...new Set(D.daily.filter(x => x.tgl < d).map(x => x.tgl))].sort().at(-1);
      if (!prev) { toast('Tidak ada log sebelumnya'); break; }
      let n = 0;
      D.daily.filter(x => x.tgl === prev).forEach(x => {
        const t = D.tasks.find(t => t.kode === x.kode);
        if (!t || stP(t) === 'FINISH' || D.daily.some(y => y.tgl === d && y.kode === x.kode)) return;
        D.daily.push({ id: newId(), tgl: d, kode: x.kode, target: '', tercapai: '', progress: null, penyebab: '', catatan: '' }); n++;
      });
      save(); render(); toast(`${n} janji disalin dari ${fmt(prev)} – isi target hari ini`);
      break;
    }
    case 'day': UI.harianTgl = addDays(UI.harianTgl || tgl(), +el.dataset.n); render(); break;
    case 'goto-ready': UI.focus = 'r-' + kode; if (UI.view === 'ready') render(); break;
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
  if (!el.dataset.f) return;
  const [kind, id, path] = el.dataset.f.split('|');
  let v = el.value;
  if (el.type === 'date') v = v || null;
  if (el.type === 'number') v = v === '' ? null : Math.max(0, Math.min(100, +v));
  if (kind === 'task') { const t = D.tasks.find(x => x.kode === id); setPath(t, path, v); }
  else if (kind === 'daily') setPath(D.daily.find(x => x.id === id), path, v);
  save(); render();
});
// Pencarian langsung saat mengetik
document.addEventListener('input', e => {
  if (e.target.dataset.ui !== 'jSearch') return;
  UI.jSearch = e.target.value;
  const pos = e.target.selectionStart;
  render();
  const s = document.querySelector('[data-ui="jSearch"]'); s.focus(); s.setSelectionRange(pos, pos);
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
