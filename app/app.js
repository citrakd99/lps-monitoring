'use strict';
/* LPS – PWA monitoring Last Planner System.
   Data = seed.json (hasil tools/export_excel.py) + perubahan lapangan di localStorage.
   Versi yang di-hosting publik tidak membawa seed.json: jadwal (master) diambil dari Google Sheet tim setelah kode akses diisi.
   Peta Excel → field: dokumen internal (tidak ikut repo publik). */

/* Mode SPV (spv.html): HANYA LIHAT. Tidak ada tombol ubah / buat / hapus / cetak / ekspor; tidak pernah mengirim data.
   Server juga menolak tulis dari kode akses SPV (VIEW_KEY di Code.gs). Data lokal terpisah dari aplikasi tim. */
const RO = document.documentElement.dataset.mode === 'spv';
// kunci penyimpanan; nama lama (versi ≤ v31: lps-<nama>-…) dipindah sekali ke nama baru tanpa kehilangan data
function lsKunci(baru, akhir) {
  try {
    if (localStorage.getItem(baru) != null) return baru;
    const re = new RegExp('^lps-[a-z]+-' + akhir + '$'), lama = Object.keys(localStorage).find(k => re.test(k));
    if (!lama) return baru;
    try { localStorage.setItem(baru, localStorage.getItem(lama)); } catch { return lama; } // penyimpanan penuh: tetap pakai yang lama
    localStorage.removeItem(lama);
  } catch { /* localStorage tidak tersedia */ }
  return baru;
}
const KEY = RO ? lsKunci('lps-spv', 'spv') : lsKunci('lps-v1', 'v1');
// aksi yang boleh di mode SPV: navigasi, filter, detail, sinkron (tarik)
const RO_OK = new Set(['settings', 'sync-now', 'sync-off', 'sd-tab', 'mp-day', 'mp-all', 'dok-f', 'drv-load', 'dok-kat', 'dok-grp', 'dok-info',
  'goto-dok', 'fprio', 'day', 'goto-ready', 'goto-komit', 'tl-goto', 'hmode', 'h-tab', 'log-goto', 'log-all', 'log-buka', 'kw', 'rantai', 'r-tgl', 'r-all']);
const RO_TAMPIL = new Set(['kst', 'sepakat', 'waste']); // tombol status: tetap terlihat (menunjukkan status) tapi nonaktif
function roLock(root) {
  if (!RO || !root) return;
  root.querySelectorAll('[data-act]').forEach(el => {
    const a = el.dataset.act;
    if (RO_OK.has(a)) return;
    if (RO_TAMPIL.has(a)) { el.disabled = true; return; }
    if (el.matches('button, .btn, .fab') || (el.tagName === 'A' && !el.children.length && a !== 'tl-ubah')) el.hidden = true;
    else { el.removeAttribute('data-act'); el.removeAttribute('href'); } // mis. chip timeline: tetap tampil sebagai teks
  });
  root.querySelectorAll('input, select, textarea').forEach(el => { if (!el.dataset.ui && !RO_OK.has(el.dataset.act) && el.type !== 'search' && !el.closest('.ro-ok')) el.disabled = true; });
}
const KAT = { eng: 'Engineering / HSE', mat: 'Material & Alat', mp: 'Man Power', oth: 'Other' };
/* TIMWOODS: 8 jenis pemborosan (lean) di janji harian, + X = bukan waste / faktor eksternal (cuaca dll.) */
const WASTE = {
  T: ['Transportation', 'Material dipindah berkali-kali, stockyard jauh, double handling'],
  I: ['Inventory', 'Material menumpuk / rusak di site, pekerjaan setengah jadi menggantung'],
  M: ['Motion', 'Bolak-balik ambil alat, cari gambar, pindah-pindah scaffolding'],
  W: ['Waiting', 'Menunggu material, alat, approval, inspeksi, ijin, pekerjaan pendahulu'],
  OP: ['Overproduction', 'Mengerjakan yang belum dibutuhkan / area belum siap, fabrikasi berlebih'],
  OV: ['Over-processing', 'Melebihi spek, laporan / inspeksi ganda, kerja ulang karena perubahan'],
  D: ['Defects', 'Rework, NCR, punch list, gagal test'],
  S: ['Skills', 'Keahlian tidak dipakai, tidak dilibatkan dalam perencanaan, kurang training'],
  X: ['Bukan waste', 'Faktor eksternal, mis. cuaca']
};
const wasteL = x => String(x.waste || '').split(',').filter(k => WASTE[k]);
const wNama = k => (k === 'X' ? '' : k[0] + ' · ') + WASTE[k][0];
const isWaste = x => wasteL(x).some(k => k !== 'X');
const jamOrang = x => x.jamHilang == null ? null : r6(x.jamHilang * (x.orang || 1)); // jam hilang × orang terdampak
const r1 = v => Math.round(v * 10) / 10;
// saran awal dari penyebab (bisa diganti di janji)
const WASTE_SARAN = [[/cuaca|hujan/i, 'X'], [/produktiv/i, 'M'], [/instruksi|perubahan.*owner|revisi/i, 'OV'], [/prioritas/i, 'OP'], [/rework|cacat|ncr|defect/i, 'D'],
  [/material|alat|tenaga|pendahulu|ijin|izin|hse|akses|area|engineering|desain/i, 'W']];
const saranWaste = p => (WASTE_SARAN.find(([re]) => re.test(p || '')) || [])[1] || '';
const KK = Object.keys(KAT);
const BLN = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];
const VIEWS = ['dash', 'jadwal', 'komit', 'ready', 'harian', 'issue', 'sd'];

const $ = s => document.querySelector(s);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const opt = (v, cur, label) => `<option value="${esc(v)}"${String(v) === String(cur ?? '') ? ' selected' : ''}>${esc(label ?? v)}</option>`;
const pct = (x, d = 0) => x == null || isNaN(x) ? '–' : (x * 100).toFixed(d) + '%';

let D = null;
const UI = { hTab: '', lgFrom: null, lgTo: null, lgPeny: '', lgDis: '', lgPic: '', lgQ: '', lgAll: false, lgWaste: '', wBuka: {}, wPilih: false, rBuka: {}, jdBuka: {}, view: 'dash', week: '', jSearch: '', jDis: '', jLook: '2', jFrom: null, harianTgl: null, focus: null, fPrio: '', sd: 'mp', mpTgl: null, mpAll: false, dokF: 'open', dokScope: '', dokJenis: '', dokSub: '', dokQ: '', dokTutup: [], llF: '', llDis: '', llQ: '', kTo: '' };

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
// status kendala: OK / Pending (ada issue) / Cek (belum dicek – default saat task disepakati, wajib dibahas di Make Ready)
const stK = t => KK.some(k => t.k[k].st !== 'OK') ? 'PENDING' : 'READY';
const belumCek = t => KK.filter(k => t.k[k].st === 'Cek');
// task baru disepakati → semua kendala yang "OK" dijadikan "belum dicek" agar benar-benar dibahas
function wajibCek(t) { KK.forEach(k => { if (t.k[k].st === 'OK') t.k[k].st = 'Cek'; }); }
function stP(t) {
  if (stK(t) !== 'READY') return '';
  if ((prog(t.kode) || 0) >= 100) return 'FINISH';
  if (t.selesai && tgl() > t.selesai) return 'LATE / DELAY';
  return 'ON PROGRESS';
}
const inW = w => t => w === '' || +t.minggu === +w;
const rangeOn = (w, w2) => w !== '' && w != null && w2 !== '' && w2 != null && +w2 > +w;
const inWR = (w, w2) => rangeOn(w, w2) ? t => +t.minggu >= +w && +t.minggu <= +w2 : inW(w);
const kRange = () => rangeOn(UI.week, UI.kTo);
const perLabel = (w, w2) => { const a = w === '' ? null : weekObj(w), b = rangeOn(w, w2) ? weekObj(w2) : null;
  return !a ? 'Semua minggu' : b ? `W${a.w}–W${b.w} · ${fmt(a.start)} – ${fmt(b.end)}` : `W${a.w} · ${fmt(a.start)} – ${fmt(a.end)}`; };
const inWeek = t => inW(UI.week)(t);
const planTasks = () => D.tasks.filter(t => t.disepakati === 'Ya' && inWeek(t));
function stJadwal(p) {
  const d = tgl();
  if (!p.mulai) return 'Tanpa jadwal';
  if (p.mulai > d) return 'Belum mulai';
  if (p.selesai < d) return 'Selesai (jadwal)';
  return 'Berjalan';
}
function issueList(w = UI.week, w2 = '') {
  const out = [];
  for (const t of D.tasks) {
    if (t.disepakati !== 'Ya' || !inWR(w, w2)(t)) continue;
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
/* RANTAI PEKERJAAN MULTI-MINGGU: task mingguan → pekerjaan total.
   induk = kode task pertama rantai (atau dikenali dari keterangan "(lanjutan T-…)"); porsi = % task ini terhadap pekerjaan total.
   Progres total = Σ porsi × progres mingguan. Porsi kosong: task pertama 100%, lanjutan = sisa porsi task sebelumnya. */
const LJ_K = /\(lanjutan (T-[^,)\s]+)/;
function rootOf(t, seen = {}) {
  if (t.induk) return t.induk;
  const m = LJ_K.exec(t.uraian || '');
  if (m && !seen[t.kode]) { const p = D.tasks.find(x => x.kode === m[1]); if (p) { seen[t.kode] = 1; return rootOf(p, seen); } }
  return t.kode;
}
function rantai(t) {
  const root = rootOf(t), M = D.tasks.filter(x => rootOf(x) === root).sort((a, b) => a.minggu - b.minggu || a.kode.localeCompare(b.kode));
  if (M.length < 2 && t.porsi == null) return null;
  const eff = [], pr = M.map(x => prog(x.kode) || 0);
  const lj = M.map(x => LJ_K.test(x.uraian || '')); // lanjutan = sisa porsi task sebelumnya (bukan cakupan baru)
  // task pertama tanpa porsi = sisa setelah task mingguan lain yang direncanakan manual (bukan lanjutan)
  const awal = Math.max(0, 100 - M.reduce((s, x, i) => s + (i && !lj[i] && x.porsi != null ? x.porsi : 0), 0));
  M.forEach((x, i) => { eff[i] = x.porsi ?? (i ? r6(eff[i - 1] * (100 - pr[i - 1]) / 100) : awal); });
  const total = Math.min(100, r6(M.reduce((s, x, i) => s + eff[i] * pr[i] / 100, 0)));
  return { root, M, eff, pr, total, sum: r6(eff.reduce((a, b, i) => a + (lj[i] ? 0 : b), 0)), idx: M.indexOf(t) };
}
const porsiEff = t => { const r = rantai(t); return r ? r.eff[r.idx] : (t.porsi ?? 100); };
// chip progres: mingguan (task ini) + total (rantai)
function progChips(t) {
  const p = prog(t.kode) || 0, r = rantai(t);
  return `<span class="chip ${p >= 100 ? 'info' : p ? 'prog' : ''}">Mingguan ${pc(p)}</span>${r ? `<a href="#" class="chip ${r.total >= 100 ? 'info' : 'warn'}" data-act="rantai" data-kode="${esc(t.kode)}" title="Lihat riwayat pekerjaan total per minggu">Total ${pc(r.total)} · minggu ${r.idx + 1}/${r.M.length} · porsi ${pc(r.eff[r.idx])}</a>` : ''}`;
}
function rantaiDlg(t) {
  const r = rantai(t);
  if (!r) return;
  const root = D.tasks.find(x => x.kode === r.root) || r.M[0];
  openDlg(`<h3>Pekerjaan total · ${esc(r.root)}</h3>
  <p style="margin:0 0 6px"><b>${esc(root.uraian.replace(LJ_RE, ''))}</b><br><small class="muted">${esc([root.pekerjaan, root.lokasi, root.pic && 'PIC ' + root.pic].filter(Boolean).join(' · '))}</small></p>
  <div class="bar-h" style="margin:6px 0"><i style="width:${r.total}%"></i></div>
  <p class="muted" style="font-size:13px;margin-top:0">Progres total <b>${pc(r.total)}</b> dari ${r.M.length} task mingguan · porsi direncanakan ${pc(r.sum)}${r.sum > 100.5 ? ' <b style="color:var(--bad)">(lebih dari 100% – periksa porsi)</b>' : r.sum < 99.5 ? ' (sisa belum direncanakan)' : ''}</p>
  <div class="scroll"><table class="tbl"><tr><th>Minggu</th><th>Task</th><th class="n">Porsi</th><th class="n">Mingguan</th><th class="n">Kontribusi</th><th>Status</th></tr>
  ${r.M.map((x, i) => `<tr${x === t ? ' style="font-weight:600"' : ''}><td>W${x.minggu}</td><td><span class="code">${esc(x.kode)}</span> <small>${esc(x.uraian.replace(LJ_RE, ''))}</small></td><td class="n">${pc(r.eff[i])}${x.porsi == null ? '<small class="muted">*</small>' : ''}</td><td class="n">${pc(r.pr[i])}</td><td class="n">${pc(r6(r.eff[i] * r.pr[i] / 100))}</td><td>${x.disepakati === 'Ya' ? stChip(stP(x) || stK(x)) : x.disepakati === 'Tidak' ? '<span class="chip bad">ditunda</span>' : '<span class="chip">belum dibahas</span>'}</td></tr>`).join('')}
  </table></div>
  <p class="muted" style="font-size:12px">* porsi otomatis (task pertama 100%, lanjutan = sisa task sebelumnya). Ubah porsi di form task bila pekerjaan sudah dibagi per minggu.</p>
  <div class="dlg-act"><span class="sp"></span><button value="cancel" class="btn">Tutup</button></div>`, null);
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
const MASTER_F = ['meta', 'pekerjaan', 'areas', 'areaMap', 'issues', 'penyebab', 'disiplin', 'source', 'exportedAt', 'mpPlan', 'dokSrc', 'llSrc', 'timeline'];
const blank = () => ({ meta: { project: '', tglUpdate: null, targetPPC: 0.8, weeks: [], scurve: [] }, pekerjaan: [], areas: [], areaMap: {},
  issues: Object.fromEntries(KK.map(k => [k, []])), penyebab: [], disiplin: ['CSA', 'MEP', 'ARS', 'PRELIM'], tasks: [], daily: [], progress: [], mpPlan: [], manpower: [], dokumen: [], timeline: [], tlEdit: [], jdEdit: [] });
const hasMaster = () => D.meta.weeks.length > 0;
function applyMaster(m) { // jadwal & referensi diganti; task, daily log, target PPC lokal dipertahankan
  const ppc = hasMaster() ? D.meta.targetPPC : m.meta.targetPPC;
  MASTER_F.forEach(f => { if (m[f] !== undefined) D[f] = m[f]; });
  D.meta = { ...m.meta, targetPPC: ppc ?? 0.8 };
}
const TXT_T = ['disiplin', 'pekerjaan', 'area', 'uraian', 'lokasi', 'pic', 'disepakati', 'alasan', 'catatanKendala', 'penyebabTerlambat', 'prio', 'koord', 'followup', 'induk'];
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
  t.minggu = +t.minggu || 0; t.mulai = t.mulai || null; t.selesai = t.selesai || null; t.porsi = nOr(t.porsi);
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
  ['target', 'penyebab', 'catatan', 'masalah', 'tindakan', 'waste'].forEach(f => { x[f] = x[f] ?? ''; });
  x.bobot = nOr(x.bobot); x.capaian = nOr(x.capaian); x.jamHilang = nOr(x.jamHilang); x.orang = nOr(x.orang);
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
  s.jdEdit = (s.jdEdit || []).map(normTlE);
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
function save() { TLc = null; if (RO) { persist(); return; } stamp(); persist(); scheduleSync(); } // SPV: hanya pengaturan perangkat

/* ---------- sinkron tim (Google Sheets via Apps Script, lihat gas/Code.gs) ----------
   Setiap record (task / daily log) diberi _u (waktu ubah) & _by (nama). Perubahan dideteksi dengan
   membandingkan hash isi record, jadi semua aksi cukup memanggil save(). Konflik: _u terbaru menang. */
const TBL = { tasks: 'kode', daily: 'id', progress: 'w', manpower: 'id', dokumen: 'id', tlEdit: 'id', jdEdit: 'id' };
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
    if (t === 'tasks' && !del && !('induk' in r) && i >= 0) { r.induk = arr[i].induk; r.porsi = arr[i].porsi; } // Code.gs lama: kolom rantai belum ada
    if (t === 'daily' && !del && !('masalah' in r) && i >= 0) { r.masalah = arr[i].masalah; r.tindakan = arr[i].tindakan; } // Code.gs lama: kolom log masalah belum ada
    if (t === 'daily' && !del && !('waste' in r) && i >= 0) { r.waste = arr[i].waste; r.jamHilang = arr[i].jamHilang; r.orang = arr[i].orang; } // Code.gs lama: kolom TIMWOODS belum ada
    if (t === 'tasks' && !del && !('prio' in r)) { PRIO_T.forEach(f => { r[f] = i >= 0 ? arr[i][f] : ''; }); oldGas = true; }
    if (del) { if (i >= 0) { arr.splice(i, 1); n++; } delete S.h[id]; }
    else {
      const rec = ({ tasks: normTask, daily: normDaily, progress: normProg, manpower: normMp, dokumen: normDok, tlEdit: normTlE, jdEdit: normTlE })[t](r), h = recHash(rec);
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
  if (!RO && j.ro) throw new Error('kode akses ini khusus SPV (hanya lihat) – pakai kode akses tim');
  if (RO && !j.ro) throw new Error('aplikasi SPV hanya menerima kode akses SPV – minta ke admin');
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
    if (RO) S.dirty = {}; // SPV tidak pernah mengirim data
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

/* koreksi tanggal pekerjaan (tab Jadwal): master tetap utuh di _m0/_s0, tanggal efektif = koreksi tim (jdEdit, id = nama pekerjaan) */
function applyJd() {
  const E = Object.fromEntries((D.jdEdit || []).map(e => [e.id, e]));
  for (const p of D.pekerjaan) {
    if (p._m0 === undefined) { p._m0 = p.mulai ?? null; p._s0 = p.selesai ?? null; } // master baru → simpan tanggal asli
    const e = E[p.nama];
    p.mulai = e ? e.mulai : p._m0; p.selesai = e ? e.selesai : p._s0; p.ubah = !!e;
  }
}
function render() {
  applyJd();
  TLc = null;
  document.querySelector('.wk').hidden = !hasMaster();
  if (!hasMaster()) {
    $('#hdr-sub').textContent = 'Belum terhubung'; $('#f-week').innerHTML = ''; $('#issue-badge').hidden = true;
    $('#view').innerHTML = vSetup(); renderSync(); return;
  }
  renderHeader();
  $('#view').innerHTML = { dash: vDash, jadwal: vJadwal, komit: vKomit, ready: vReady, harian: vHarian, issue: vIssue, sd: vSD }[UI.view]();
  roLock($('#view'));
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
  ${wasteCard()}
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

// Dashboard: 3 waste terbesar minggu ini (jam-orang, lalu kejadian)
function wasteCard() {
  if (!D.daily.some(isWaste)) return '';
  const cw = weekOf(tgl()), byK = Object.fromEntries(D.tasks.map(t => [t.kode, t]));
  const R = D.daily.filter(x => weekOf(x.tgl) === cw && x.tgl <= tgl() && isWaste(x)).map(x => ({ x, t: byK[x.kode] }));
  const par = wastePar(R).filter(([k]) => k !== 'X'), jo = R.reduce((s, r) => s + (jamOrang(r.x) || 0), 0);
  return `<div class="card"><div class="vh"><h3 style="margin:0">Top 3 waste minggu ini (W${cw})</h3><a href="#harian" data-act="log-buka" style="font-size:13px">Log masalah →</a></div>
    ${par.length ? `${wasteParH(par.slice(0, 3))}<p class="muted" style="font-size:12px;margin:4px 0 0">${R.length} janji dengan pemborosan · ${r1(jo)} jam-orang terbuang minggu ini</p>` : '<p class="muted">Belum ada pemborosan tercatat minggu ini.</p>'}</div>`;
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
  <div class="vh"><h2>Jadwal & Lookahead</h2><button class="btn sm" data-act="pdf-jadwal">PDF / Excel</button></div>
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
      ${p.ubah ? `<div class="fu">Tanggal diubah dari master ${fmt(p._m0)} → ${fmt(p._s0)}${(() => { const e = D.jdEdit.find(x => x.id === p.nama); return e ? `${e.catatan ? ' · ' + esc(e.catatan) : ''}${e._by ? ' · oleh ' + esc(e._by) : ''}` : ''; })()}</div>` : ''}
      ${(() => { const A = TL().filter(a => a.pek === p.nama && !a.ms).sort((a, b) => a.mulai.localeCompare(b.mulai) || a.selesai.localeCompare(b.selesai));
        const Av = look === 'all' || !wF || !wT ? A : A.filter(a => a.mulai <= wT.end && a.selesai >= wF.start), kos = Av.filter(a => !tlTasks(a).length);
        if (!A.length) return `<div class="bar" style="margin:8px 0 0"><button class="btn sm" data-act="task-new" data-pek="${esc(p.nama)}">+ Task komitmen</button><button class="btn sm" data-act="jd-ubah" data-pek="${esc(p.nama)}">✎ Ubah tanggal</button></div>`;
        return `<details class="tl-jd" data-pek="${esc(p.nama)}"${UI.jdBuka[p.nama] ? ' open' : ''}><summary>Aktivitas timeline${look === 'all' ? '' : ' di lookahead'}: ${Av.length}${Av.length < A.length ? ` dari ${A.length}` : ''} · <span class="chip ${kos.length ? 'warn' : 'ok'}">${kos.length} belum jadi task</span></summary>
          ${Av.map(a => { const ts = tlTasks(a), w = tlWeekFor(a), ada = tlTasks(a, w).length;
            return `<div class="tl-r"><span><b>${esc(a.nama)}</b>${tlSt(a)}${ts.length ? ` <span class="chip ok">task: ${ts.map(t => esc(t.kode) + ' W' + t.minggu).join(', ')}</span>` : ''}<small>${tlInfo(a)}</small></span>
              <span class="tl-b"><button class="btn sm" data-act="tl-ubah" data-id="${esc(a.id)}" title="Ubah tanggal aktivitas timeline" aria-label="Ubah tanggal">✎</button>${ada ? '' : `<button class="btn sm" data-act="tl-add" data-id="${esc(a.id)}" data-w="${w}" title="Jadikan task komitmen W${w}">+ Task W${w}</button>`}</span></div>`; }).join('') || '<p class="muted" style="font-size:13px">Tidak ada aktivitas di rentang lookahead ini.</p>'}
        </details>
        <div class="bar" style="margin:8px 0 0"><button class="btn sm pri" data-act="tl-pick" data-pek="${esc(p.nama)}" data-w="${tlWeekFor(kos[0] || Av[0] || A[0])}">+ Task dari timeline…</button><button class="btn sm" data-act="task-new" data-pek="${esc(p.nama)}">+ Task manual</button><button class="btn sm" data-act="jd-ubah" data-pek="${esc(p.nama)}">✎ Ubah tanggal</button></div>`; })()}
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
/* minggu task untuk aktivitas: minggu pertama (≥ minggu ini) aktivitas aktif yang belum ada task-nya */
const tlWeekFor = a => { const cw = weekOf(tgl()) ?? D.meta.weeks[0].w, last = D.meta.weeks.at(-1).w;
  if (!a) return Math.min(cw + 1, last);
  const ws = D.meta.weeks.filter(x => x.w >= cw && a.mulai <= x.end && a.selesai >= x.start), free = ws.find(x => !tlTasks(a, x.w).length);
  return (free || ws[0])?.w ?? Math.min(Math.max(cw, weekOf(a.mulai) ?? cw), last); };
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
function tlPick(w, pek = '') {
  w = +w;
  const cw = weekOf(tgl()), o = weekObj(w);
  const akt = tlAkt(w).filter(a => !pek || a.pek === pek).sort((a, b) => a.area.localeCompare(b.area) || a.mulai.localeCompare(b.mulai));
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
  ${pek ? `<p class="muted" style="font-size:13px;margin:4px 0">Pekerjaan: <b>${esc(pek.replace(/^\S+\s/, ''))}</b> · <a href="#" data-act="tl-pek-all">tampilkan semua pekerjaan</a></p>` : ''}
  <div class="bar" style="margin:2px 0 0"><button type="button" class="btn sm" data-act="tl-manual">+ Task manual (di luar timeline)</button><button type="button" class="btn sm" data-act="tl-all">Centang yang tampil & belum ada task</button><button type="button" class="btn sm" data-act="tl-none">Kosongkan</button><span class="muted" id="tl-n" style="font-size:13px"></span></div>
  <div class="pick" id="tl-pick" data-pek="${esc(pek)}">${akt.map(a => { const ada = tlTasks(a, w); return `<label class="pick-i" data-area="${esc(a.area)}" data-dis="${esc(a.dis)}" data-cari="${esc((a.nama + ' ' + (a.lok || '') + ' ' + (a.pek || '')).toLowerCase())}">
    <input type="checkbox" name="bl" value="${esc(a.id)}"${ada.length ? ' disabled' : ''}>
    <span><b>${esc(a.nama)}</b>${ada.length ? ` <span class="chip ok">sudah: ${ada.map(t => esc(t.kode)).join(', ')}</span>` : ''}${tlSt(a)}${a.selesai <= o.end ? ' <span class="chip warn">harus selesai</span>' : ''}
    <small>${tlInfo(a)}<br>${a.pek ? esc(a.pek) : '<i>pekerjaan BOQ belum dipetakan</i>'}</small></span></label>`; }).join('') || `<p class="muted" style="padding:10px">Tidak ada aktivitas timeline${pek ? ' pekerjaan ini' : ''} di minggu ini.</p>`}</div>
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
function jdUbah(nama) {
  applyJd();
  const p = D.pekerjaan.find(x => x.nama === nama), e = D.jdEdit.find(x => x.id === nama);
  if (!p) return;
  const ts = D.tasks.filter(t => t.pekerjaan === nama);
  openDlg(`<h3>Ubah tanggal pekerjaan</h3>
  <p style="margin:0 0 4px"><span class="code">${esc(p.kode)}</span> <b>${esc(p.nama.replace(/^\S+\s/, ''))}</b><br><small class="muted">${esc([p.bill, p.area, p.disiplin].filter(Boolean).join(' · '))}</small></p>
  <p class="muted" style="font-size:13px">Tanggal di master: <b>${fmt(p._m0)} – ${fmt(p._s0)}</b>${e?._by ? ` · terakhir diubah ${esc(e._by)}` : ''}. Hanya mulai & selesai yang berubah; berlaku untuk seluruh tim (lookahead, status jadwal, PDF). Bobot & kurva-S tidak berubah.</p>
  <div class="frow">
    <label class="f">Mulai<input type="date" name="mulai" value="${esc(p.mulai || '')}" required></label>
    <label class="f">Selesai<input type="date" name="selesai" value="${esc(p.selesai || '')}" required></label>
  </div>
  <label class="f">Alasan / catatan<input type="text" name="catatan" value="${esc(e?.catatan || '')}" placeholder="mis. ikut jadwal aktual, revisi baseline"></label>
  ${ts.length ? `<p class="muted" style="font-size:13px">${ts.length} task komitmen di pekerjaan ini tidak ikut berubah; ubah tanggalnya di task masing-masing bila perlu.</p>` : ''}
  <div class="dlg-act">${e ? `<button type="button" class="btn danger" data-act="jd-reset" data-pek="${esc(p.nama)}">Kembalikan ke master</button>` : ''}<span class="sp"></span>
    <button value="cancel" class="btn" formnovalidate>Batal</button><button value="save" class="btn pri">Simpan</button></div>`,
  fd => {
    const m = fd.get('mulai'), sl = fd.get('selesai');
    if (!m || !sl || sl < m) { toast('Tanggal selesai tidak boleh sebelum mulai'); return; }
    D.jdEdit = D.jdEdit.filter(x => x.id !== nama);
    const sama = m === p._m0 && sl === p._s0;
    if (!sama) D.jdEdit.push({ id: nama, mulai: m, selesai: sl, catatan: String(fd.get('catatan') || '').trim() });
    save(); render(); toast(sama ? 'Tanggal kembali sesuai master' : `Tanggal pekerjaan diubah: ${fmt(m)} – ${fmt(sl)}`);
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
/* LANJUTKAN KE MINGGU DEPAN: task disepakati yang belum FINISH (atau ditunda) dibuat ulang di minggu berikutnya.
   Task lama tetap di minggunya (PPC minggu itu tetap jujur); task baru = sisa pekerjaan, status disepakati dikosongkan untuk rapat. */
const LJ_RE = /\s*\(lanjutan T-[^)]*\)$/;
const lanjutanDari = t => D.tasks.find(x => +x.minggu === +t.minggu + 1 && x.uraian.includes(`(lanjutan ${t.kode}`));
const perluLanjut = t => !!weekObj(+t.minggu + 1) && +t.minggu <= weekOf(tgl()) && !lanjutanDari(t)
  && ((t.disepakati === 'Ya' && (prog(t.kode) || 0) < 100) || t.disepakati === 'Tidak');
function lanjutkan(t) {
  const w = +t.minggu + 1, wo = weekObj(w), p = prog(t.kode) || 0, pe = porsiEff(t);
  if (!wo) return null;
  const n = normTask({ ...JSON.parse(JSON.stringify(t)), kode: nextKode(), minggu: w, disepakati: '', alasan: '', inLPS: false, penyebabTerlambat: '', followup: '',
    uraian: `${t.uraian.replace(LJ_RE, '')} (lanjutan ${t.kode}${p ? `, s/d ${r1(p)}%` : ''})`,
    mulai: wo.start, selesai: !t.selesai || t.selesai < wo.start || t.selesai > wo.end ? wo.end : t.selesai,
    induk: rootOf(t), porsi: r1(pe * (100 - Math.min(100, p)) / 100) });
  delete n._u; delete n._by; delete n._rev;
  D.tasks.push(n);
  return n;
}
function vKomit() {
  const rg = kRange(), ts = D.tasks.filter(inWR(UI.week, UI.kTo)).filter(prioOK).sort((a, b) => (rg ? a.minggu - b.minggu : 0) || byPrio(a, b) || a.minggu - b.minggu || a.kode.localeCompare(b.kode));
  let gap = TL().length ? tlGapCard(tlWeek()) : '';
  if (!TL().length && UI.week !== '') {
    const w = weekObj(UI.week);
    const aktif = D.pekerjaan.filter(p => p.mulai && w && p.mulai <= w.end && p.selesai >= w.start);
    const kosong = aktif.filter(p => !D.tasks.some(t => t.pekerjaan === p.nama && +t.minggu === +UI.week));
    gap = `<details class="card"><summary>Pekerjaan MASTER aktif minggu ${UI.week}: ${aktif.length} · <span class="chip warn">${kosong.length} belum ada task</span></summary>
      ${kosong.map(p => `<div style="display:flex;justify-content:space-between;gap:8px;padding:5px 0;border-bottom:1px solid var(--line)"><span><span class="code">${esc(p.kode)}</span> ${esc(p.nama.replace(/^\S+\s/, ''))} <small class="muted">${esc(p.disiplin)}</small></span><button class="btn sm" data-act="task-new" data-pek="${esc(p.nama)}">+ Task</button></div>`).join('')}
    </details>`;
  }
  const n = ts.length, ya = ts.filter(t => t.disepakati === 'Ya').length, cw = weekOf(tgl());
  const lj = UI.week === '' || rg ? [] : D.tasks.filter(t => +t.minggu === +UI.week && perluLanjut(t));
  const kw = [...(weekObj(cw - 1) ? [[String(cw - 1), `Minggu lalu · W${cw - 1}`]] : []), [String(cw), `Minggu ini · W${cw}`], ...(weekObj(cw + 1) ? [[String(cw + 1), `Minggu depan · W${cw + 1}`]] : []), ['', 'Semua minggu']];
  return `
  ${vh(`Komitmen mingguan ${UI.week === '' ? '' : '· W' + UI.week + (rg ? '–W' + UI.kTo : '')}`, 'komit')}
  <div class="prio-f">${kw.map(([v, l]) => `<button type="button" data-act="kw" data-v="${v}" class="${!rg && String(UI.week) === v ? 'on' : ''}">${l}</button>`).join('')}${UI.week !== '' && !rg && !kw.some(([v]) => v === String(UI.week)) ? `<button type="button" class="on" data-act="kw" data-v="${UI.week}">W${UI.week}</button>` : ''}<button type="button" data-act="k-rg" class="${rg ? 'on' : ''}">Beberapa minggu…</button></div>
  ${rg || UI.kRgBuka ? `<div class="bar k-rg"><span class="muted" style="font-size:13px">Rentang</span>
    <select data-ui="week" aria-label="Dari minggu">${D.meta.weeks.map(x => opt(x.w, UI.week === '' ? cw : UI.week, `W${x.w} · ${x.label}${x.w === cw ? ' (ini)' : ''}`)).join('')}</select>
    <span class="muted">s/d</span>
    <select data-ui="kTo" aria-label="Sampai minggu">${opt('', rg ? UI.kTo : '', '— pilih —')}${D.meta.weeks.filter(x => UI.week === '' || x.w > +UI.week).map(x => opt(x.w, UI.kTo, `W${x.w} · ${x.label}`)).join('')}</select>
    ${rg ? '<button type="button" class="btn sm" data-act="kw" data-v="' + UI.week + '">Hanya 1 minggu</button>' : ''}</div>` : ''}
  ${prioBar()}
  ${rg ? (() => { const ws = D.meta.weeks.filter(x => x.w >= +UI.week && x.w <= +UI.kTo);
    return `<div class="card" style="padding:10px 12px"><b>Ringkasan W${UI.week}–W${UI.kTo}</b><div class="scroll"><table class="tbl"><tr><th>Minggu</th><th class="n">Task</th><th class="n">Disepakati</th><th class="n">Selesai 100%</th><th class="n">PPC</th></tr>
      ${ws.map(x => { const T = D.tasks.filter(t => +t.minggu === x.w && prioOK(t)), S = T.filter(t => t.disepakati === 'Ya'), F = S.filter(t => (prog(t.kode) || 0) >= 100), P = ppcWeek(x.w);
        return `<tr${x.w === cw ? ' style="background:var(--warn-bg)"' : ''}><td>W${x.w} <small class="muted">${esc(x.label)}</small></td><td class="n">${T.length}</td><td class="n">${S.length}</td><td class="n">${F.length}</td><td class="n">${x.w <= cw ? pct(P.ppc) : '–'}</td></tr>`; }).join('')}</table></div></div>`; })() : ''}
  ${!rg && UI.week !== '' && +UI.week <= cw ? (() => { const w = +UI.week, S = D.tasks.filter(t => +t.minggu === w && t.disepakati === 'Ya'), fin = S.filter(t => (prog(t.kode) || 0) >= 100), P = ppcWeek(w);
    return `<div class="card" style="padding:10px 12px"><b>${w === cw ? `Progres W${w} (berjalan)` : `Hasil W${w}`}</b>: ${S.length} disepakati · <b>${fin.length} selesai 100%</b> · ${S.length - fin.length} belum · PPC (Finish ÷ Final) <b>${pct(P.ppc)}</b>
      ${S.length ? `<div class="bar-h" style="margin-top:6px"><i style="width:${fin.length / S.length * 100}%"></i></div>` : ''}</div>`; })() : ''}
  ${lj.length ? `<div class="fu" style="margin-bottom:10px;display:flex;gap:8px;align-items:center;flex-wrap:wrap"><span style="flex:1;min-width:200px"><b>${lj.length} task W${UI.week}</b> belum selesai / ditunda dan belum dilanjutkan ke W${+UI.week + 1}.</span><button class="btn sm" data-act="lanjut-semua" data-w="${UI.week}">Lanjutkan semua ke W${+UI.week + 1}</button></div>` : ''}
  <p class="muted" style="margin-top:-6px">${n} task · ${ya} disepakati${n ? ' (' + pct(ya / n) + ')' : ''}. Pilih <b>Ya</b> saat rapat mingguan untuk task yang sanggup dikerjakan.</p>
  ${gap}
  ${ts.map((t, i) => `${rg && (i === 0 || ts[i - 1].minggu !== t.minggu) ? (() => { const o = weekObj(t.minggu), T = ts.filter(x => x.minggu === t.minggu);
      return `<h3 class="k-wk">W${t.minggu}${o ? ` <small class="muted">${fmt(o.start)} – ${fmt(o.end)} · ${T.length} task · ${T.filter(x => x.disepakati === 'Ya').length} disepakati</small>` : ''}</h3>`; })() : ''}
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
        ${t.disepakati === 'Ya' ? stChip(stK(t)) + stChip(stP(t)) + `<a href="#ready" class="r-ptr" data-act="goto-ready" data-kode="${esc(t.kode)}">${belumCek(t).length ? `⚠ ${belumCek(t).length} kendala belum dicek` : 'Make ready'} →</a>` : ''}
        ${t.disepakati === 'Ya' || rantai(t) ? progChips(t) : ''}
        <span style="flex:1"></span>
        ${prioSel(t)}
        ${perluLanjut(t) ? `<button class="btn sm" data-act="lanjut" data-kode="${esc(t.kode)}" title="Buat task lanjutan (sisa pekerjaan) di minggu berikutnya">→ W${+t.minggu + 1}</button>` : ''}
        <button class="btn sm" data-act="task-edit" data-kode="${esc(t.kode)}">Ubah</button>
      </div>
      ${(() => { const l = lanjutanDari(t); return l ? `<div class="muted" style="font-size:12px;margin-top:6px">↪ dilanjutkan ke W${l.minggu} sebagai <a href="#komit" data-act="kw" data-v="${l.minggu}">${esc(l.kode)}</a></div>` : ''; })()}
      ${fuNote(t)}
      ${t.disepakati === 'Tidak' ? `<input type="text" style="margin-top:8px" placeholder="Alasan ditunda / catatan rapat" value="${esc(t.alasan)}" data-f="task|${esc(t.kode)}|alasan" class="${t.alasan ? '' : 'need'}">` : ''}
    </div>`).join('') || '<div class="empty">Belum ada task untuk filter ini. Tekan <b>+ Task</b>.</div>'}
  <button class="btn pri fab" data-act="${TL().length ? 'tl-pick' : 'task-new'}" data-w="${TL().length ? tlWeek() : ''}">+ Task</button>`;
}

/* pilihan template = aktivitas timeline yang aktif di minggu task; aktivitas pekerjaan yang sama ditaruh paling atas */
function tplOpts(w, pek, sel) {
  const akt = tlAkt(w), lab = a => { const ts = tlTasks(a, w); return `${a.nama}${a.lok ? ' · ' + a.lok : ''} (${fmt(a.mulai)} – ${fmt(a.selesai)})${ts.length ? ' · sudah ' + ts.map(t => t.kode).join(', ') : ''}`; };
  const og = (l, A) => A.length ? `<optgroup label="${esc(l)}">${A.map(a => opt(a.id, sel, lab(a))).join('')}</optgroup>` : '';
  const P = pek ? akt.filter(a => a.pek === pek) : [], R = akt.filter(a => !P.includes(a));
  return opt('', sel, akt.length ? `– Manual (tanpa template) · ${akt.length} aktivitas W${w} tersedia –` : `– Manual · tidak ada aktivitas timeline di W${w} –`)
    + og('Pekerjaan ini', P) + [...new Set(R.map(a => a.area))].sort().map(ar => og(ar || 'Lainnya', R.filter(a => a.area === ar))).join('');
}
function tplFill(id) {
  const E = $('#dlg-form').elements, a = tlById(id), o = weekObj(E.minggu.value);
  E.bl.value = a ? String(a.id) : '';
  if (!a) return;
  const set = (n, v) => { if (E[n] && v) E[n].value = v; };
  set('uraian', a.nama); set('lokasi', a.lok); set('pekerjaan', a.pek);
  if (D.areas.includes(a.area)) set('area', a.area);
  if (D.disiplin.includes(a.dis)) set('disiplin', a.dis);
  if (o) { set('mulai', a.mulai > o.start ? a.mulai : o.start); set('selesai', a.selesai < o.end ? a.selesai : o.end); }
}
function taskForm(t, isNew) {
  const pek = D.pekerjaan.map(p => p.nama);
  return `<h3>${isNew ? 'Task baru' : 'Ubah task ' + esc(t.kode)}</h3>
  <div class="frow">
    <label class="f">Kode<input type="text" name="kode" value="${esc(t.kode)}" required ${isNew ? '' : 'readonly'}></label>
    <label class="f">Minggu ke<select name="minggu">${D.meta.weeks.map(x => opt(x.w, t.minggu, `W${x.w} · ${x.label}`)).join('')}</select></label>
  </div>
  ${isNew && TL().length ? `<label class="f">Template dari timeline <small class="muted">(opsional – pilih untuk mengisi otomatis, atau biarkan Manual)</small><select name="tpl" id="tf-tpl">${tplOpts(t.minggu, t.pekerjaan, t.bl || '')}</select></label>
  <input type="hidden" name="bl" value="${esc(t.bl || '')}">` : ''}
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
  ${(() => { const roots = [...new Set(D.tasks.filter(x => x.kode !== t.kode && +x.minggu < +t.minggu).sort((a, b) => (b.pekerjaan === t.pekerjaan) - (a.pekerjaan === t.pekerjaan) || b.minggu - a.minggu).map(rootOf))].filter(k => k !== t.kode);
    const cur = t.induk || (rootOf(t) !== t.kode ? rootOf(t) : '');
    return `<div class="frow">
    <label class="f">Bagian dari pekerjaan multi-minggu<select name="induk">${opt('', cur, 'Tidak – selesai dalam minggu ini')}${[...new Set([cur, ...roots].filter(Boolean))].map(k => { const x = D.tasks.find(y => y.kode === k); return opt(k, cur, `${k} · ${x ? (x.uraian.replace(LJ_RE, '').slice(0, 50) + ' (W' + x.minggu + ')') : ''}`); }).join('')}</select></label>
    <label class="f">Porsi dari pekerjaan total (%)<input type="number" name="porsi" min="0" max="100" step="any" inputmode="decimal" value="${t.porsi ?? ''}" placeholder="kosong = otomatis"></label>
  </div>`; })()}
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
// Make Ready: task READY diringkas (judul + progres), PENDING / LATE terbuka; pilihan buka/tutup diingat selama sesi
const rBuka = t => UI.rBuka[t.kode] ?? (stK(t) === 'PENDING' || stP(t) === 'LATE / DELAY');
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
    ${(() => { const n = ts.filter(t => belumCek(t).length).length; return n ? `<div class="kpi warn"><b>${n}</b><span>Task kendalanya belum dicek</span></div>` : ''; })()}
  </div>
  <div class="bar"><span class="muted" style="font-size:13px;flex:1">Task READY otomatis diringkas – ketuk judul untuk membuka.</span>
    <button class="btn sm" data-act="r-all" data-v="1">Buka semua</button><button class="btn sm" data-act="r-all" data-v="0">Ringkas semua</button></div>
  ${ts.filter(prioOK).sort((a, b) => byPrio(a, b) || (stK(a) === 'PENDING' ? 0 : 1) - (stK(b) === 'PENDING' ? 0 : 1) || a.kode.localeCompare(b.kode)).map(t => {
    const p = prog(t.kode), sp = stP(t), buka = rBuka(t);
    const kr = KK.filter(k => t.k[k].st !== 'OK'), kSum = !kr.length ? 'Kendala: semua OK' : KK.map(k => t.k[k].st === 'OK' ? '' : `${KAT[k]} ${t.k[k].st === 'Cek' ? 'belum dicek' : 'pending'}`).filter(Boolean).join(' · ');
    const progLine = `${buka ? '' : `<a href="#" class="r-ptr" data-act="r-tgl" data-kode="${esc(t.kode)}">${esc(kSum)} · lihat / ubah status →</a>`}<div style="display:flex;align-items:center;gap:8px;margin:8px 0 ${buka ? 8 : 0}px"><div class="bar-h" style="flex:1"><i style="width:${Math.min(p || 0, 100)}%"></i></div><small class="muted">Progress ${p == null ? '0' : p}%${buka ? ' (dari Daily Log)' : ''}</small>${rantai(t) ? progChips(t).replace(/^<span[^>]*>[^<]*<\/span>/, '') : ''}${stK(t) === 'READY' && sp !== 'FINISH' ? `<button class="btn sm" data-act="daily-new" data-kode="${esc(t.kode)}">+ Janji harian</button>` : ''}</div>`;
    return `<div class="item${buka ? '' : ' tutup'}" id="r-${esc(t.kode)}">
      <div class="hd lipat" data-act="r-tgl" data-kode="${esc(t.kode)}" role="button" aria-expanded="${buka}"><div><span class="lp">${buka ? '▾' : '▸'}</span> <span class="code">${esc(t.kode)}</span> <span class="ttl">${esc(t.uraian)}</span>
        <div class="sub">${esc(t.pekerjaan)}${buka ? '' : ` · W${t.minggu} · PIC ${esc(t.pic || '–')}`}</div></div>
        <div class="chips">${prioChips(t)}${stChip(stK(t))}${stChip(sp)}</div></div>
      ${buka ? '' : progLine}
      ${!buka ? '' : `<div class="meta"><span>W${t.minggu} · ${esc(t.disiplin)}</span><span>${esc(t.area)}${t.lokasi ? ' · ' + esc(t.lokasi) : ''}</span><span>PIC ${esc(t.pic || '–')}</span><span>${fmt(t.mulai)} → ${fmt(t.selesai)}</span>${prioSel(t)}</div>
      ${fuNote(t)}
      ${dokLine(t)}
      ${progLine}
      ${KK.map(k => {
        const c = t.k[k], pend = c.st === 'Pending', cek = c.st === 'Cek', f = `task|${esc(t.kode)}|k.${k}.`;
        const over = pend && c.target && c.target < d;
        return `<div class="kr${cek ? ' cek' : ''}">
          <div class="kh"><span>${KAT[k]}${cek ? ' <span class="chip warn">belum dicek</span>' : ''}</span>
            <span class="seg">
              <button type="button" data-act="kst" data-kode="${esc(t.kode)}" data-k="${k}" data-v="OK" class="${c.st === 'OK' ? 'on-ok' : ''}">OK</button>
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
      <input type="text" style="margin-top:8px" placeholder="Catatan kendala" value="${esc(t.catatanKendala)}" data-f="task|${esc(t.kode)}|catatanKendala">`}
    </div>`;
  }).join('') || '<div class="empty">Belum ada task yang disepakati. Buka tab <a href="#komit">Komitmen</a>.</div>'}`;
}

/* kartu satu janji di dalam header task (Harian) */
// = kolom Cek di Daily Log
const cek = (x, t) => !t ? '⚠ Kode tidak ada' : stK(t) !== 'READY' ? '⚠ Task belum READY' : !x.target || x.bobot == null ? '⚠ Isi target & bobot'
  : x.bobot > sisaTask(x) + 0.01 ? '⚠ Bobot > sisa task' : x.capaian == null ? 'Menunggu laporan sore'
  : tercapai(x) === 'Tidak' && !x.penyebab ? '⚠ Isi penyebab' : tercapai(x) === 'Tidak' && !wasteL(x).length ? '⚠ Isi jenis waste' : 'OK';
// pilihan TIMWOODS di kartu janji: wajib bila tidak tercapai, opsional ("Ada pemborosan?") bila tercapai
function wasteBox(x, f, wajib) {
  const on = wasteL(x);
  if (!wajib && !on.length && !UI.wBuka[x.id]) return `<button type="button" class="btn sm" data-act="waste-buka" data-id="${x.id}" style="margin-bottom:8px">+ Ada pemborosan? (TIMWOODS)</button>`;
  return `<div class="f wst"><span>Jenis waste (TIMWOODS)${wajib ? '' : ' · opsional, walau target tercapai'}</span>
    <div class="wst-c">${Object.keys(WASTE).map(k => `<button type="button" class="${on.includes(k) ? 'on' : ''}${wajib && !on.length ? ' need' : ''}" data-act="waste" data-id="${x.id}" data-v="${k}" title="${esc(WASTE[k][1])}">${esc(wNama(k))}</button>`).join('')}</div>
    ${on.length ? `<p class="muted" style="font-size:12px;margin:0 0 6px">${on.map(k => esc(WASTE[k][0] + ': ' + WASTE[k][1])).join('<br>')}</p>` : ''}
    ${isWaste(x) ? `<div class="frow">
      <label class="f">Jam hilang<input type="number" min="0" max="24" step="any" inputmode="decimal" value="${x.jamHilang ?? ''}" data-f="${f}jamHilang" placeholder="mis. 2"></label>
      <label class="f">Orang terdampak<input type="number" min="0" max="500" step="1" inputmode="numeric" value="${x.orang ?? ''}" data-f="${f}orang" placeholder="mis. 6"></label></div>
      ${jamOrang(x) != null ? `<p class="muted" style="font-size:12px;margin:-4px 0 8px">≈ <b>${r1(jamOrang(x))}</b> jam-orang terbuang</p>` : ''}` : ''}</div>`;
}
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
      ${tercapai(x) === 'Tidak' ? `<label class="f">Penyebab tidak tercapai<select data-f="${f}penyebab" class="${x.penyebab ? '' : 'need'}">${opt('', x.penyebab, '– pilih –')}${D.penyebab.map(p => opt(p, x.penyebab)).join('')}</select></label>
      ${wasteBox(x, f, true)}
      <label class="f">Detail masalah (log)<textarea rows="2" data-f="${f}masalah" placeholder="Apa yang terjadi, di mana, berapa lama, dampaknya">${esc(x.masalah)}</textarea></label>
      <label class="f">Tindak lanjut / rencana perbaikan<input type="text" value="${esc(x.tindakan)}" data-f="${f}tindakan" placeholder="mis. tambah 4 orang besok, minta material dikirim pagi"></label>` : tercapai(x) === 'Ya' ? wasteBox(x, f, false) : ''}
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
const hTabs = () => `<div class="prio-f">${[['', 'Daily commitment'], ['log', 'Log masalah']].map(([k, l]) => `<button type="button" data-act="h-tab" data-v="${k}" class="${UI.hTab === k ? 'on' : ''}">${l}</button>`).join('')}</div>`;
function vHarian() {
  if (UI.hTab === 'log') return vLog();
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
  ${hTabs()}
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

/* ---------- LOG MASALAH HARIAN: riwayat janji yang tidak tercapai (dan opsional belum dilaporkan) per tanggal ---------- */
const lgOpt = () => ({ from: UI.lgFrom || addDays(tgl(), -13), to: UI.lgTo || tgl(), peny: UI.lgPeny, dis: UI.lgDis, pic: UI.lgPic, q: UI.lgQ, all: UI.lgAll, waste: UI.lgWaste });
function logRows(o) {
  const byK = Object.fromEntries(D.tasks.map(t => [t.kode, t])), q = String(o.q || '').toLowerCase().trim();
  return D.daily.filter(x => x.tgl >= o.from && x.tgl <= o.to && x.tgl <= tgl())
    .filter(x => tercapai(x) === 'Tidak' || (x.capaian != null && wasteL(x).length) || (o.all && x.capaian == null)) // + janji tercapai yang mencatat pemborosan
    .map(x => ({ x, t: byK[x.kode] }))
    .filter(({ x, t }) => (!o.peny || (x.penyebab || '(belum diisi)') === o.peny) && (!o.dis || t?.disiplin === o.dis) && (!o.pic || (t?.pic || '') === o.pic) && (!o.waste || wasteL(x).includes(o.waste))
      && (!q || [x.kode, t?.uraian, t?.lokasi, t?.pekerjaan, x.target, x.penyebab, x.masalah, x.tindakan, x.catatan, ...wasteL(x).map(k => WASTE[k][0])].join(' ').toLowerCase().includes(q)))
    .sort((a, b) => b.x.tgl.localeCompare(a.x.tgl) || a.x.kode.localeCompare(b.x.kode));
}
const logStat = o => { const n = D.daily.filter(x => x.tgl >= o.from && x.tgl <= o.to && x.tgl <= tgl()), lap = n.filter(x => tercapai(x));
  return { n: n.length, tidak: lap.filter(x => tercapai(x) === 'Tidak').length, lap: lap.length, belum: n.length - lap.length }; };
// TIMWOODS per periode: kejadian & jam-orang (jam-orang dibagi rata bila 1 janji punya >1 waste)
const wastePar = rows => { const p = {}; rows.forEach(({ x }) => { const L = wasteL(x); L.forEach(k => { const a = p[k] || (p[k] = { n: 0, jo: 0 }); a.n++; a.jo += (jamOrang(x) || 0) / L.length; }); });
  return Object.entries(p).sort((a, b) => (a[0] === 'X') - (b[0] === 'X') || b[1].jo - a[1].jo || b[1].n - a[1].n); };
const wasteParH = par => { if (!par.length) return '<p class="muted">Belum ada waste tercatat.</p>';
  const mj = Math.max(...par.map(([, a]) => a.jo)), mn = Math.max(...par.map(([, a]) => a.n));
  return par.map(([k, a]) => `<div style="margin:6px 0"><div style="display:flex;justify-content:space-between;font-size:13px"><span title="${esc(WASTE[k][1])}">${esc(wNama(k))}</span><span class="muted">${a.n}×${a.jo ? ` · ${r1(a.jo)} jam-orang` : ''}</span></div><div class="bar-h"><i style="width:${mj ? a.jo / mj * 100 : a.n / mn * 100}%;background:${k === 'X' ? 'var(--muted)' : 'var(--warn)'}"></i></div></div>`).join(''); };
// matriks penyebab × waste: sumber pemborosan terbesar
function wasteMat(rows) {
  const R = rows.filter(r => wasteL(r.x).length), ks = Object.keys(WASTE).filter(k => R.some(r => wasteL(r.x).includes(k)));
  if (!R.length) return '';
  const pk = r => r.x.penyebab || (tercapai(r.x) === 'Ya' ? '(tercapai)' : '(tanpa penyebab)'), ps = [...new Set(R.map(pk))];
  return `<div class="scroll"><table class="tbl wst-m"><tr><th>Penyebab \\ waste</th>${ks.map(k => `<th class="n" title="${esc(WASTE[k][0])}">${k}</th>`).join('')}</tr>
    ${ps.map(p => `<tr><td>${esc(p)}</td>${ks.map(k => { const n = R.filter(r => pk(r) === p && wasteL(r.x).includes(k)).length; return `<td class="n">${n || ''}</td>`; }).join('')}</tr>`).join('')}</table></div>`;
}
const pareto = rows => { const p = {}; rows.filter(r => tercapai(r.x) === 'Tidak').forEach(({ x }) => { const k = x.penyebab || '(belum diisi)'; p[k] = (p[k] || 0) + 1; }); return Object.entries(p).sort((a, b) => b[1] - a[1]); };
function vLog() {
  const o = lgOpt(), L = logRows(o), S = logStat(o), par = pareto(L), parT = par.reduce((s, x) => s + x[1], 0);
  const tdk = L.filter(r => tercapai(r.x) === 'Tidak'), tanpaDetail = tdk.filter(r => !r.x.masalah).length;
  const wp = wastePar(L), jo = L.reduce((s, r) => s + (isWaste(r.x) ? jamOrang(r.x) || 0 : 0), 0), tanpaW = tdk.filter(r => !wasteL(r.x).length).length;
  const bocor = tdk.filter(r => wasteL(r.x).includes('W') && r.t && stK(r.t) === 'READY').length; // Waiting padahal sudah READY = make ready bocor
  const bisaSaran = D.daily.filter(x => tercapai(x) === 'Tidak' && !wasteL(x).length && saranWaste(x.penyebab)).length;
  const penyL = [...new Set([...D.penyebab, ...L.map(r => r.x.penyebab || '(belum diisi)')])], picL = [...new Set(D.tasks.map(t => t.pic).filter(Boolean))].sort();
  let h = '', d = null;
  for (const { x, t } of L) {
    if (x.tgl !== d) { d = x.tgl; const n = L.filter(r => r.x.tgl === d).length; h += `<h3 class="mp-g">${fmt(d)} <small class="muted">W${weekOf(d)} · ${n} masalah</small></h3>`; }
    const belum = x.capaian == null;
    h += `<div class="item lg-i" data-act="log-goto" data-tgl="${x.tgl}" data-kode="${esc(x.kode)}">
      <div class="hd"><div><span class="code">${esc(x.kode)}</span> <span class="ttl">${esc(t?.uraian || '(task tidak ada)')}</span>
        <div class="sub">${esc([t?.lokasi, t?.disiplin, t?.pic && 'PIC ' + t.pic].filter(Boolean).join(' · '))}</div></div>
        <div class="chips">${belum ? '<span class="chip warn">belum dilaporkan</span>' : `<span class="chip ${tercapai(x) === 'Ya' ? 'ok' : 'bad'}">tercapai ${pc(x.capaian)}</span>`}${wasteL(x).map(k => `<span class="chip ${k === 'X' ? '' : 'warn'}" title="${esc(WASTE[k][1])}">${esc(wNama(k))}</span>`).join('')}${x.penyebab ? `<span class="chip">${esc(x.penyebab)}</span>` : belum || tercapai(x) === 'Ya' ? '' : '<span class="chip warn">penyebab belum diisi</span>'}</div></div>
      <div class="meta"><span>Target: <b>${esc(x.target || '–')}</b></span><span>Bobot ${pc(x.bobot)} → actual ${belum ? '–' : pc(aktual(x))}</span>${isWaste(x) && jamOrang(x) != null ? `<span>Terbuang <b>${r1(jamOrang(x))} jam-orang</b></span>` : ''}</div>
      ${x.masalah ? `<div class="lg-t"><b>Masalah:</b> ${esc(x.masalah)}</div>` : tercapai(x) === 'Tidak' ? '<div class="lg-t muted">Detail masalah belum diisi</div>' : ''}
      ${x.tindakan ? `<div class="lg-t"><b>Tindak lanjut:</b> ${esc(x.tindakan)}</div>` : ''}
      ${x.catatan ? `<div class="lg-t muted">Catatan: ${esc(x.catatan)}</div>` : ''}
    </div>`;
  }
  return `<div class="vh"><h2>Log masalah harian</h2><button class="btn sm" data-act="pdf-log">PDF / Excel</button></div>
  ${hTabs()}
  <div class="bar">
    <label class="muted" style="font-size:13px">Dari <input type="date" value="${o.from}" data-ui="lgFrom"></label>
    <label class="muted" style="font-size:13px">s/d <input type="date" value="${o.to}" data-ui="lgTo"></label>
    <select data-ui="lgPeny" style="width:auto">${opt('', o.peny, 'Semua penyebab')}${penyL.map(p => opt(p, o.peny)).join('')}</select>
    <select data-ui="lgDis" style="width:auto">${opt('', o.dis, 'Semua disiplin')}${D.disiplin.map(p => opt(p, o.dis)).join('')}</select>
    <select data-ui="lgPic" style="width:auto">${opt('', o.pic, 'Semua PIC')}${picL.map(p => opt(p, o.pic)).join('')}</select>
    <select data-ui="lgWaste" style="width:auto">${opt('', o.waste, 'Semua waste')}${Object.keys(WASTE).map(k => opt(k, o.waste, wNama(k))).join('')}</select>
    <input type="search" data-ui="lgQ" value="${esc(o.q)}" placeholder="Cari task / masalah / tindakan" style="flex:1;min-width:160px">
    <label class="muted" style="font-size:13px"><input type="checkbox" data-act="log-all" ${o.all ? 'checked' : ''} style="width:auto;min-height:0"> sertakan janji belum dilaporkan</label>
  </div>
  <div class="grid kpis">
    <div class="kpi"><b>${S.n}</b><span>Janji ${fmt(o.from)} – ${fmt(o.to)}</span></div>
    <div class="kpi bad"><b>${S.tidak}</b><span>Tidak tercapai${S.lap ? ' · ' + pct(S.tidak / S.lap) + ' dari yang dilaporkan' : ''}</span></div>
    <div class="kpi warn"><b>${S.belum}</b><span>Belum dilaporkan</span></div>
    <div class="kpi ${tanpaDetail ? 'warn' : 'ok'}"><b>${tanpaDetail}</b><span>Tidak tercapai tanpa detail masalah</span></div>
    <div class="kpi ${jo ? 'bad' : ''}"><b>${r1(jo)}</b><span>Jam-orang terbuang (TIMWOODS)</span></div>
    <div class="kpi ${bocor ? 'bad' : 'ok'}"><b>${bocor}</b><span>Waiting padahal task READY · make ready bocor</span></div>
    <div class="kpi ${tanpaW ? 'warn' : 'ok'}"><b>${tanpaW}</b><span>Tidak tercapai tanpa jenis waste</span></div>
  </div>
  ${bisaSaran ? `<div class="bar"><button class="btn sm" data-act="waste-saran">Isi jenis waste otomatis dari penyebab (${bisaSaran} janji lama)</button><span class="muted" style="font-size:12px">Saran awal – periksa & ubah di janjinya bila perlu</span></div>` : ''}
  <div class="two"><div>${h || '<div class="empty">Tidak ada masalah tercatat untuk filter ini.</div>'}</div>
    <div class="card"><h3>Pareto penyebab (periode ini)</h3>
    ${par.length ? (() => { let cum = 0; return par.map(([k, v]) => { cum += v; return `<div style="margin:6px 0"><div style="display:flex;justify-content:space-between;font-size:13px"><span>${esc(k)}</span><span class="muted">${v} · ${pct(v / parT)} · kum ${pct(cum / parT)}</span></div><div class="bar-h"><i style="width:${v / par[0][1] * 100}%;background:var(--bad)"></i></div></div>`; }).join(''); })() : '<p class="muted">Belum ada penyebab tercatat.</p>'}
    <h3 style="margin-top:14px">Pareto TIMWOODS (periode ini)</h3>
    ${wasteParH(wp)}
    ${wasteMat(L) ? `<h3 style="margin-top:14px">Penyebab × waste</h3>${wasteMat(L)}<p class="muted" style="font-size:12px">${Object.keys(WASTE).filter(k => k !== 'X').map(k => `${k} = ${WASTE[k][0]}`).join(' · ')}</p>` : ''}
    <p class="muted" style="font-size:12px">Ketuk satu masalah untuk membuka janjinya di Daily commitment (isi detail masalah, jenis waste & tindak lanjut di sana).</p></div></div>`;
}
function pdfLogForm() {
  const o = lgOpt(), penyL = [...new Set([...D.penyebab, ...D.daily.map(x => x.penyebab).filter(Boolean)])], picL = [...new Set(D.tasks.map(t => t.pic).filter(Boolean))].sort();
  return `<h3>PDF Log Masalah Harian</h3>
  <p class="muted" style="font-size:13px;margin-top:0">Terisi sesuai filter di layar; ubah bila perlu.</p>
  <div class="frow"><label class="f">Dari<input type="date" name="from" value="${o.from}" required></label><label class="f">Sampai<input type="date" name="to" value="${o.to}" required></label></div>
  <div class="frow">
    <label class="f">Penyebab<select name="peny">${opt('', o.peny, 'Semua penyebab')}${penyL.map(p => opt(p, o.peny)).join('')}${opt('(belum diisi)', o.peny)}</select></label>
    <label class="f">Disiplin<select name="dis">${opt('', o.dis, 'Semua disiplin')}${D.disiplin.map(p => opt(p, o.dis)).join('')}</select></label>
  </div>
  <label class="f">Jenis waste (TIMWOODS)<select name="waste">${opt('', o.waste, 'Semua waste')}${Object.keys(WASTE).map(k => opt(k, o.waste, wNama(k))).join('')}</select></label>
  <div class="frow">
    <label class="f">PIC<select name="pic">${opt('', o.pic, 'Semua PIC')}${picL.map(p => opt(p, o.pic)).join('')}</select></label>
    <label class="f">Susunan<select name="grp">${opt('tgl', 'tgl', 'Satu tabel, urut tanggal')}${opt('pic', 'tgl', '1 halaman per PIC')}</select></label>
  </div>
  <label class="f">Cari<input type="text" name="q" value="${esc(o.q)}" placeholder="task / masalah / tindakan"></label>
  <label class="f" style="flex-direction:row;align-items:center;gap:8px"><input type="checkbox" name="all" value="1" ${o.all ? 'checked' : ''} style="width:auto;min-height:0"> Sertakan janji yang belum dilaporkan</label>
  <p class="muted" style="font-size:13px">Kertas A4 mendatar. Di jendela cetak pilih <b>Simpan sebagai PDF</b>.</p>
  <div class="dlg-act"><button value="cancel" class="btn" formnovalidate>Batal</button><button value="xls" class="btn">Excel</button><button value="save" class="btn pri">Buat PDF</button></div>`;
}
function printLog(o) {
  if (o.to < o.from) [o.from, o.to] = [o.to, o.from];
  const L = logRows(o).reverse(); // PDF: urut tanggal naik
  if (!L.length) { toast('Tidak ada masalah tercatat untuk filter ini'); return; }
  const groups = o.grp === 'pic' ? [...new Set(L.map(r => r.t?.pic || '(tanpa PIC)'))].sort().map(p => [p, L.filter(r => (r.t?.pic || '(tanpa PIC)') === p)]) : [['', L]];
  const per = `${fmt(o.from)} – ${fmt(o.to)}`, filt = [o.peny, o.waste && 'Waste ' + WASTE[o.waste][0], o.dis && 'Disiplin ' + o.dis, o.pic && 'PIC ' + o.pic, o.q && `"${o.q}"`].filter(Boolean).join(' · ') || 'Semua masalah';
  const head = ['No', 'Tanggal', 'Kode', 'Task · lokasi', 'PIC', 'Target hari ini', 'Bobot', 'Actual', '% tercapai', 'Penyebab', 'Waste', 'Jam-orang', 'Detail masalah', 'Tindak lanjut', 'Catatan'];
  const S = logStat(o);
  const rows = rs => { let d = null; return rs.map(({ x, t }, i) => { const sh = x.tgl !== d ? `<tr class="sb"><td colspan="${head.length}">${fmt(x.tgl)} · W${weekOf(x.tgl)}</td></tr>` : ''; d = x.tgl; const b = x.capaian == null;
    return sh + `<tr><td>${i + 1}</td><td>${fmt(x.tgl)}</td><td>${esc(x.kode)}</td><td>${esc([t?.uraian, t?.lokasi].filter(Boolean).join(' · '))}</td><td>${esc(t?.pic || '')}</td><td>${esc(x.target)}</td>
      <td>${pc(x.bobot)}</td><td>${b ? '' : pc(aktual(x))}</td><td class="u${b || tercapai(x) === 'Ya' ? '' : ' bad'}">${b ? 'belum lapor' : pc(x.capaian)}</td><td>${esc(x.penyebab || '')}</td><td>${esc(wasteL(x).map(k => WASTE[k][0]).join(', '))}</td><td>${isWaste(x) && jamOrang(x) != null ? r1(jamOrang(x)) : ''}</td><td class="rm">${esc(x.masalah || '')}</td><td class="rm">${esc(x.tindakan || '')}</td><td class="rm">${esc(x.catatan || '')}</td></tr>`; }).join(''); };
  const ringkas = rs => { const par = pareto(rs), tot = par.reduce((s, x) => s + x[1], 0), wp = wastePar(rs);
    return `${rs.filter(r => tercapai(r.x) === 'Tidak').length} janji tidak tercapai${wp.length ? ' · TIMWOODS: ' + wp.map(([k, a]) => `${esc(WASTE[k][0])} ${a.n}×${a.jo ? ' / ' + r1(a.jo) + ' jam-orang' : ''}`).join(', ') : ''}${rs.some(r => r.x.capaian == null) ? ` · ${rs.filter(r => r.x.capaian == null).length} belum dilaporkan` : ''}${par.length ? ' · Penyebab: ' + par.map(([k, v]) => `${esc(k)} ${v} (${pct(v / tot)})`).join(', ') : ''}`; };
  $('#print').innerHTML = groups.map(([g, rs]) => `<section class="pg dok lg">
    <header><div><b>Log Masalah Harian</b><span>${esc(D.meta.project)}</span></div>
      <div class="pr-meta"><span>Periode: <b>${esc(per)}</b></span>${g ? `<span>PIC: <b>${esc(g)}</b></span>` : ''}<span>${esc(filt)}</span><span>Dicetak ${fmt(todayISO())}${D.sync.user ? ' oleh ' + esc(D.sync.user) : ''}</span></div></header>
    <table><thead><tr>${head.map(h => `<th>${esc(h)}</th>`).join('')}</tr></thead><tbody>${rows(rs)}</tbody></table>
    <p class="pr-n">${ringkas(rs)}${g ? '' : ` · Periode ini: ${S.n} janji, ${S.tidak} tidak tercapai${S.lap ? ' (' + pct(S.tidak / S.lap) + ' dari yang dilaporkan)' : ''}`}</p>
  </section>`).join('');
  const judul = document.title;
  document.title = ['Log_Masalah', o.from, o.to, o.pic || o.dis || ''].filter(Boolean).join('_').replace(/[\\/:*?"<>|\s]+/g, '-');
  addEventListener('afterprint', () => { document.title = judul; $('#print').innerHTML = ''; }, { once: true });
  selesaiCetak();
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
  return `<div class="ro-ok"><h3>Data & pengaturan</h3>
  <label class="f">Tanggal update<select name="tglMode">${opt('auto', s.tglMode, 'Otomatis = hari ini')}${opt('fixed', s.tglMode, 'Tanggal tertentu')}</select></label>
  <label class="f">Tanggal (bila tertentu)<input type="date" name="tglFixed" value="${esc(s.tglFixed || D.meta.tglUpdate || '')}"></label>
  ${RO ? '' : `<label class="f">Target PPC (%)<input type="number" name="targetPPC" min="0" max="100" value="${Math.round(D.meta.targetPPC * 100)}"></label>`}
  <fieldset style="border:1px solid var(--line);border-radius:10px;padding:10px 12px;margin:0 0 12px">
    <legend style="font-weight:600;font-size:14px;padding:0 4px">${RO ? 'Data tim – SPV (hanya lihat)' : 'Sinkron tim (Google Sheets)'}</legend>
    <label class="f">URL Web App (…/exec)<input type="text" name="syncUrl" inputmode="url" placeholder="https://script.google.com/macros/s/…/exec" value="${esc(D.sync.url)}"></label>
    <div class="frow">
      <label class="f">${RO ? 'Kode akses SPV' : 'Kode akses tim'}<input type="text" name="syncKey" autocomplete="off" value="${esc(D.sync.key)}"></label>
      <label class="f">Nama Anda<input type="text" name="syncUser" placeholder="mis. Agus SM" value="${esc(D.sync.user)}"></label>
    </div>
    <p class="muted" style="font-size:13px;margin:0 0 8px">${syncOn() ? `Status: ${esc($('#sync')?.title || '')}${D.sync.lastOk ? ' · terakhir ' + esc(new Date(D.sync.lastOk).toLocaleString('id-ID')) : ''}` : 'Belum terhubung. Data hanya di perangkat ini.'}</p>
    ${syncOn() ? '<div style="display:flex;gap:8px;flex-wrap:wrap"><button type="button" class="btn sm" data-act="sync-now">Sinkron sekarang</button><button type="button" class="btn sm danger" data-act="sync-off">Putuskan</button></div>' : ''}
  </fieldset>
  <p class="muted" style="font-size:13px">Sumber: <b>${esc(D.source || '-')}</b><br>Diekspor dari Excel: ${esc((D.exportedAt || '').replace('T', ' '))}<br>
  ${D.tasks.length} task · ${D.daily.length} daily log · ${D.pekerjaan.length} pekerjaan. ${syncOn() ? 'Disalin di perangkat ini & disinkron ke Google Sheet tim.' : 'Data tersimpan di perangkat ini.'}</p>
  ${RO ? '<p class="muted" style="font-size:13px">Mode SPV: semua data hanya bisa dilihat. Perubahan dilakukan tim di aplikasi utama.</p>' : ''}
  <div class="grid" style="grid-template-columns:1fr 1fr"${RO ? ' hidden' : ''}>
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
  <div class="dlg-act"><button value="cancel" class="btn" formnovalidate>Tutup</button><button value="save" class="btn pri">Simpan</button></div></div>`;
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
    h.push('Catatan Kendala', 'Progress (%)', 'Penyebab Terlambat', 'Status Kendala', 'Status Pelaksanaan', 'Prioritas', 'Koordinasi Disiplin', 'Catatan Follow-up', 'Induk (pekerjaan multi-minggu)', 'Porsi dari Total (%)', 'Progress Total (%)');
    const rows = D.tasks.map(t => [t.kode, t.minggu, t.disiplin, t.pekerjaan, t.area, t.uraian, t.lokasi, t.pic, t.mulai, t.selesai, t.disepakati, t.alasan,
      ...KK.flatMap(k => [t.k[k].st, t.k[k].issue, t.k[k].detail, t.k[k].pic, t.k[k].target]),
      t.catatanKendala, prog(t.kode), t.penyebabTerlambat, t.disepakati === 'Ya' ? stK(t) : '', t.disepakati === 'Ya' ? stP(t) : '', t.prio, t.koord, t.followup, rantai(t)?.root ?? '', rantai(t) ? porsiEff(t) : t.porsi, rantai(t)?.total ?? '']);
    download(`LPS-Input_${tgl()}.csv`, csv([h, ...rows]), 'text/csv');
  } else {
    const h = ['Tanggal', 'Kode Task', 'Target Hari Ini', 'Bobot Target (% dari task)', '% Tercapai Harian (actual ÷ bobot)', 'Tercapai?', 'Progress s/d Kemarin', 'Kontribusi Hari Ini', 'Progress Kumulatif (%)', 'Penyebab Tidak Tercapai', 'Catatan', 'Detail Masalah', 'Tindak Lanjut', 'Jenis Waste (TIMWOODS)', 'Jam Hilang', 'Orang Terdampak', 'Jam-Orang Terbuang'];
    const rows = [...D.daily].sort((a, b) => a.tgl.localeCompare(b.tgl)).map(x => [x.tgl, x.kode, x.target, x.bobot, x.capaian, tercapai(x), progKemarin(x), kontrib(x), progKum(x), x.penyebab, x.catatan, x.masalah, x.tindakan, wasteL(x).map(k => WASTE[k][0]).join(', '), x.jamHilang, x.orang, isWaste(x) ? jamOrang(x) : null]);
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
// menunggu siapa (dokumen Submitted/Open): owner bila CW sudah respon atau remarks "awaiting owner", selain itu CW; For Info & yang sudah dibalas owner tidak dihitung
const tungguDok = x => umurDok(x) == null || x.own || /for info/i.test(x.stAsli || '') ? '' : x.cw || /awaiting owner/i.test(x.issue || '') ? 'own' : 'cw';
const DOK_FX = { lewat3: x => umurDok(x) > 3, cw: x => tungguDok(x) === 'cw', own: x => tungguDok(x) === 'own' };
const dokOfTask = kode => D.dokumen.filter(x => String(x.kode).split(',').map(s => s.trim()).includes(kode));

/* ---------- Dokumen dari log Google Drive (HANYA BACA lewat link berbagi "Pelihat"; tidak pernah menulis ke Drive) ----------
   Tiap sheet "LOG LIST OF …" (kolom B–L: No, Document Number, Rev, Description, Estimated Submission, Actual Submission,
   Response CW, Response owner, Status, Remarks, PIC) dibaca via endpoint CSV Google (gviz). Link disimpan di master tim
   (push_master.py --dok-link) atau per perangkat; tidak ada di kode publik. Hasil di-cache lokal agar tetap tampil offline. */
const DOK_KEY = lsKunci('lps-dok', 'dok');
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
// hanya link http(s) yang dipakai (isi sel dari Drive tidak dipercaya mentah-mentah)
const urlOk = v => { v = String(v || '').trim(); return /^https:\/\/[^\s"'<>]+$/i.test(v) ? v : ''; };
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
    by.get(key).revs.push({ rev: +c(3) || 0, est: tglTeks(c(5)), sub: tglTeks(c(6)), cw: tglTeks(c(7)), own: tglTeks(c(8)), st: c(9), rem: c(10), pic: c(11), fl: urlOk(c(12)) });
  }
  const link = `https://docs.google.com/spreadsheets/d/${id}/edit#gid=${gid}`;
  return [...by.entries()].map(([key, d]) => {
    const revs = d.revs.sort((a, b) => a.rev - b.rev), x = revs.at(-1), status = stDrv(x.st || ''), fl = [...revs].reverse().find(v => v.fl);
    return { id: 'g:' + jenis + ':' + key, drive: true, jenis, scope: d.scope, sub: d.sub, nomor: d.no, nama: d.desc || d.no, rev: x.rev,
      status, stAsli: x.st || 'Belum submit', target: x.est, tglSubmit: x.sub, cw: x.cw, own: x.own, respon: x.own || x.cw || null, tglApprove: status === 'Approved' ? (x.own || x.cw || null) : null,
      issue: x.rem, pic: x.pic, kode: '', link, file: fl ? fl.fl : '', fileRev: fl ? fl.rev : null,
      riwayat: revs.length < 2 ? "" : revs.map(v => `Rev ${v.rev} · ${v.st || '-'}${v.sub ? ' · submit ' + fmt(v.sub) : ''}${v.own || v.cw ? ' · respon ' + fmt(v.own || v.cw) : ''}${v.rem ? '\n   ' + v.rem.replace(/\s*\n\s*/g, ' ') : ''}`).join('\n') };
  });
}
// daftar [nama, gid] sheet dari halaman htmlview file Drive (link berbagi "Pelihat")
async function drvTabs(base) {
  const html = await (await fetch(base + 'htmlview')).text();
  const sh = [...html.matchAll(/\{name: "((?:[^"\\]|\\.)*)", pageUrl: "[^"]*", gid: "(\d+)"/g)]
    .map(m => [m[1].replace(/\\x([0-9a-f]{2})/gi, (_, h) => String.fromCharCode(parseInt(h, 16))).replace(/\\(.)/g, '$1'), m[2]]);
  if (!sh.length) throw new Error('Daftar sheet tidak terbaca. Pastikan akses link: "Siapa saja yang memiliki link" (Pelihat).');
  return sh;
}
const drvCsv = async (base, gid, n) => { const r = await fetch(base + 'gviz/tq?tqx=out:csv&headers=0&gid=' + gid); if (!r.ok) throw new Error(`Sheet ${n}: HTTP ${r.status}`); return parseCsv(await r.text()); };
async function drvLoad(manual) {
  const id = drvId(dokSrc()); if (!id || drvBusy) return;
  drvBusy = true; if (UI.view === 'sd') render();
  const base = `https://docs.google.com/spreadsheets/d/${id}/`;
  try {
    const sh = await drvTabs(base);
    const docs = (await Promise.all(sh.map(async ([n, gid]) => drvSheet(n, gid, await drvCsv(base, gid, n), id)))).flat();
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

/* ---------- LONG LEAD MATERIAL dari log Google Drive (hanya baca, cara sama dengan Dokumen) ----------
   Sheet "LOG LIST OF LONG LEAD MATERIAL": kolom dicari dari judulnya (Material/Equipment, Lead Time, Required On Site, …).
   Link disimpan di master tim (push_master.py --ll-link) atau per perangkat; tidak ada di kode publik. */
const LL_KEY = lsKunci('lps-ll', 'll');
let LLD = (() => { try { return JSON.parse(localStorage.getItem(LL_KEY)) || {}; } catch { return {}; } })();
let llBusy = false;
const llSrc = () => D.settings.llSrc || D.llSrc || '';
const llItems = () => (LLD.src && LLD.src === drvId(llSrc()) && LLD.items) || [];
// tanggal dari teks bebas: "31 Jan 2027", "MOS Maks W1 February 2027", "Partial 1 - W3 Desember 2026, …" → tanggal paling awal
function llTgl(s) {
  s = String(s || ''); const out = [], p2 = v => String(v).padStart(2, '0'), mo = m => BLN_N[m.slice(0, 3).toLowerCase()];
  for (const m of s.matchAll(/\b(\d{4})-(\d{2})-(\d{2})\b/g)) out.push(m[0]);
  for (const m of s.matchAll(/\b(\d{1,2})[\s-]+([A-Za-z]{3,})\.?[\s-]+(\d{4})/g)) if (mo(m[2])) out.push(`${m[3]}-${p2(mo(m[2]))}-${p2(m[1])}`);
  for (const m of s.matchAll(/\bW\s*([1-5])\s+([A-Za-z]{3,})\s+(\d{4})/gi)) if (mo(m[2])) out.push(`${m[3]}-${p2(mo(m[2]))}-${p2(1 + (m[1] - 1) * 7)}`);
  return out.sort()[0] || null;
}
// lead time teks → hari (ambil yang terpanjang: "Ready Stock & Indent 3 Months" = 90)
function llDur(s) {
  s = String(s || '').toLowerCase(); let d = null;
  for (const m of s.matchAll(/(\d+(?:[.,]\d+)?)(?:\s*-\s*(\d+(?:[.,]\d+)?))?\s*(week|minggu|month|bulan|day|hari)/g)) {
    const n = parseFloat((m[2] || m[1]).replace(',', '.')), k = /week|minggu/.test(m[3]) ? 7 : /month|bulan/.test(m[3]) ? 30 : 1;
    d = Math.max(d ?? 0, Math.round(n * k));
  }
  return d ?? (/ready/.test(s) ? 0 : null);
}
function llSheet(rows) {
  const hi = rows.findIndex(r => r.some(c => /material\s*\/\s*equipment/i.test(c)));
  if (hi < 0) return [];
  const H = rows[hi].map(c => String(c).replace(/\s+/g, ' ').trim().toLowerCase()), m = H.findIndex(h => /material\s*\/\s*equipment/.test(h));
  // judul kolom angka (No, Qty, Planned PO Date, Progress) bisa kosong di CSV → posisi baku relatif kolom Material/Equipment
  const col = (re, off) => { const i = H.findIndex(h => re.test(h)); return i >= 0 ? i : m + off; };
  const C = { dis: col(/dicipline|discipline|disiplin/, 1), area: col(/^area/, 2), vendor: col(/vendor|supplier/, 3), eng: col(/engineering status/, 4), ref: col(/submittal ref|po\/submittal/, 5),
    qty: col(/^qty/, 6), unit: col(/^unit/, 7), lead: col(/lead time/, 8), ros: col(/required on site/, 9), ppo: col(/planned po/, 10), apo: col(/actual po/, 11),
    sub: col(/^submittal status/, 12), proc: col(/procurement/, 13), manu: col(/manufactur/, 14), del: col(/delivery status/, 15), parr: col(/planned arrival/, 16),
    farr: col(/forecast arrival/, 17), prog: col(/^progress/, 18), dev: col(/deviation/, 19), upd: col(/weekly/, 20), risk: col(/risk/, 21), act: col(/action/, 22),
    pic: col(/^pic/, 23), last: col(/last update/, 24), rem: col(/remark/, 25) };
  let n = 0;
  return rows.slice(hi + 1).map(r => {
    const c = i => String(r[i] ?? '').trim(), nama = c(m);
    if (!nama) return null;
    const x = Object.fromEntries(Object.entries(C).map(([k, i]) => [k, c(i)]));
    let pg = parseFloat(String(x.prog).replace(',', '.')); if (isNaN(pg)) pg = null; else if (!/%/.test(x.prog) && pg > 0 && pg <= 1) pg *= 100;
    return { id: 'll' + (++n), nama, ...x, prog: pg };
  }).filter(Boolean);
}
async function llLoad(manual) {
  const id = drvId(llSrc()); if (!id || llBusy) return;
  llBusy = true; if (UI.view === 'sd') render();
  const base = `https://docs.google.com/spreadsheets/d/${id}/`;
  try {
    const sh = await drvTabs(base);
    const items = (await Promise.all(sh.map(async ([n, gid]) => llSheet(await drvCsv(base, gid, n)).map(x => ({ ...x, link: base + 'edit#gid=' + gid }))))).flat();
    LLD = { src: id, at: new Date().toISOString(), items, err: items.length ? '' : 'Kolom "Material/Equipment" tidak ditemukan di file ini.' };
  } catch (e) {
    LLD = { src: id, items: LLD.src === id ? LLD.items : [], at: new Date().toISOString(), err: /fetch|network/i.test(e.message) ? 'Tidak bisa membaca Drive (offline, atau link belum dibagikan "Siapa saja yang memiliki link")' : e.message };
  }
  llBusy = false;
  try { localStorage.setItem(LL_KEY, JSON.stringify(LLD)); } catch { /* penyimpanan penuh */ }
  render();
  if (manual) toast(LLD.err || `${LLD.items.length} item long lead dibaca dari Drive`);
}
const llAuto = () => { const id = drvId(llSrc()); if (id && !llBusy && (!LLD.at || LLD.src !== id || Date.now() - new Date(LLD.at) > 6e5)) setTimeout(() => llLoad(false)); };
// status item terhadap HARI INI: tanggal butuh di site (ROS), batas PO = Planned PO atau ROS − lead time
function llInfo(x) {
  const d = tgl(), ros = llTgl(x.ros), lead = llDur(x.lead), poTgl = llTgl(x.apo), po = poTgl || (x.apo && !/^(not|belum|-)/i.test(x.apo) ? x.apo : '');
  const bPO = llTgl(x.ppo) || (ros && lead != null ? addDays(ros, -lead) : null), bHitung = !llTgl(x.ppo) && !!bPO;
  const tiba = /deliver|arriv|on ?site|received|complete|tiba/i.test(x.del) && !/not|belum|partial|on the way/i.test(x.del);
  const appr = /approv/i.test(x.eng) && !/not/i.test(x.eng), tibaT = llTgl(x.farr) || llTgl(x.parr);
  const sisa = ros ? days(d, ros) : null;
  const st = tiba ? 'tiba' : !ros ? 'cek' : ros < d ? 'lewat' : !po && bPO && bPO < d ? 'po' : (tibaT && tibaT > ros) ? 'telat'
    : sisa <= 30 || (!po && bPO && days(d, bPO) <= 14) ? 'kritis' : 'aman';
  return { ros, lead, po, poTgl, bPO, bHitung, tiba, appr, sisa, tibaT, st };
}
const LL_ST = { lewat: ['bad', 'Lewat tgl butuh di site'], po: ['bad', 'PO terlambat'], telat: ['bad', 'Tiba setelah tgl butuh'], kritis: ['warn', 'Kritis'], cek: ['', 'Tanggal belum jelas'], aman: ['ok', 'On track'], tiba: ['info', 'Sudah tiba'] };
const LL_URUT = ['lewat', 'po', 'telat', 'kritis', 'cek', 'aman', 'tiba'];
const LL_F = { '': 'Semua', aksi: 'Perlu tindakan', ma: 'Belum approve', nopo: 'Belum PO', tiba: 'Sudah tiba' };
const llCocok = (o, f) => !f || (f === 'aksi' ? ['lewat', 'po', 'telat', 'kritis'].includes(o.i.st) : f === 'ma' ? !o.i.appr : f === 'nopo' ? !o.i.po : f === 'tiba' ? o.i.tiba : true);
const llDisOf = x => x.dis || 'Lainnya';
// o = { f, dis, q } → [{ x, i }] terurut: paling mendesak dulu, lalu tanggal butuh di site
function llFilter(o) {
  const q = String(o.q || '').toLowerCase().trim();
  return llItems().map(x => ({ x, i: llInfo(x) }))
    .filter(r => (!o.dis || llDisOf(r.x) === o.dis) && (!q || [r.x.nama, r.x.vendor, r.x.area, r.x.dis, r.x.pic, r.x.risk, r.x.act].join(' ').toLowerCase().includes(q)) && llCocok(r, o.f))
    .sort((a, b) => LL_URUT.indexOf(a.i.st) - LL_URUT.indexOf(b.i.st) || String(a.i.ros || '9').localeCompare(String(b.i.ros || '9')));
}
function pdfLLForm() {
  const disL = [...new Set(llItems().map(llDisOf))];
  return `<h3>PDF Long Lead Material</h3>
  <p class="muted" style="font-size:13px;margin-top:0">Terisi sesuai filter di layar; ubah bila perlu.</p>
  <div class="frow">
    <label class="f">Status / filter<select name="f">${Object.entries(LL_F).map(([k, l]) => opt(k, UI.llF, l)).join('')}</select></label>
    <label class="f">Disiplin<select name="dis">${opt('', UI.llDis, 'Semua disiplin')}${disL.map(d => opt(d, UI.llDis)).join('')}</select></label>
  </div>
  <div class="frow">
    <label class="f">Cari<input type="text" name="q" value="${esc(UI.llQ)}" placeholder="material / vendor / area"></label>
    <label class="f">Susunan<select name="grp">${opt('dis', 'dis', '1 halaman per disiplin')}${opt('satu', 'dis', 'Satu tabel')}</select></label>
  </div>
  <label class="f" style="flex-direction:row;align-items:center;gap:8px"><input type="checkbox" name="pic" value="1" checked style="width:auto;min-height:0"> Sertakan kolom PIC / kontak vendor</label>
  <p class="muted" style="font-size:13px">Kertas A4 mendatar, urut dari yang paling mendesak. Di jendela cetak pilih <b>Simpan sebagai PDF</b>.</p>
  <div class="dlg-act"><button value="cancel" class="btn" formnovalidate>Batal</button><button value="xls" class="btn">Excel</button><button value="save" class="btn pri">Buat PDF</button></div>`;
}
function printLL(o) {
  const L = llFilter(o);
  if (!L.length) { toast('Tidak ada item untuk filter ini'); return; }
  const groups = o.grp === 'dis' ? [...new Set(L.map(r => llDisOf(r.x)))].map(d => [d, L.filter(r => llDisOf(r.x) === d)]) : [[o.dis || 'Semua disiplin', L]];
  const filt = [o.f && LL_F[o.f], o.q && `"${o.q}"`].filter(Boolean).join(' · ') || 'Semua item';
  const head = ['No', 'Material / equipment', 'Area · vendor', 'Qty', 'Lead time', 'Butuh di site', 'Batas / rencana PO', 'PO aktual', 'Perkiraan tiba', 'Material approval', 'Procurement · manufacturing · delivery', 'Progress', 'Status', 'Risk / action / update', ...(o.pic ? ['PIC / kontak'] : [])];
  const t = v => esc(String(v || '').replace(/\s+/g, ' ').trim());
  const rows = rs => rs.map(({ x, i }, n) => { const [cl, lb] = LL_ST[i.st];
    return `<tr><td>${n + 1}</td><td class="nm">${esc(x.nama.trim()).replace(/\n+/g, '<br>')}</td><td>${t([x.area, x.vendor].filter(Boolean).join(' · '))}</td><td>${t([x.qty, x.unit].filter(Boolean).join(' '))}</td>
      <td>${t(x.lead)}</td><td class="tg">${i.ros ? `<b>${fmt(i.ros)}</b>${i.tiba ? '' : `<br><small>${i.sisa < 0 ? 'lewat ' + -i.sisa : i.sisa} hr</small>`}` : t(x.ros)}${x.ros && !/^\d/.test(x.ros) && i.ros ? `<br><small>${t(x.ros)}</small>` : ''}</td>
      <td class="tg">${i.bPO ? `${fmt(i.bPO)}${i.bHitung ? '<br><small>ROS − lead</small>' : ''}` : ''}</td><td class="tg">${i.po ? t(i.poTgl ? fmt(i.poTgl) : i.po) : 'belum'}</td>
      <td class="tg">${i.tibaT ? fmt(i.tibaT) : t(x.farr || x.parr)}</td><td>${t(x.eng)}${x.sub ? `<br><small>Submittal: ${t(x.sub)}</small>` : ''}${x.ref ? `<br><small>${t(x.ref)}</small>` : ''}</td>
      <td>${[x.proc, x.manu, x.del].map(t).join(' · ')}</td><td class="c">${x.prog == null ? '' : pc(x.prog)}${x.dev ? `<br><small>dev ${t(x.dev)} hr</small>` : ''}</td>
      <td class="st ${cl}">${lb}</td><td class="rm">${[x.risk && '⚠ ' + x.risk, x.act && '→ ' + x.act, x.upd, x.rem].filter(Boolean).map(t).join('<br>')}</td>${o.pic ? `<td>${t(x.pic)}</td>` : ''}</tr>`; }).join('');
  const ring = rs => { const c = st => rs.filter(r => r.i.st === st).length;
    return `${rs.length} item · ${c('lewat')} lewat tgl butuh · ${c('po')} PO terlambat · ${c('telat')} tiba setelah tgl butuh · ${c('kritis')} kritis · ${rs.filter(r => !r.i.appr).length} belum approve · ${rs.filter(r => !r.i.po).length} belum PO · ${c('tiba')} sudah tiba`; };
  $('#print').innerHTML = groups.map(([g, rs]) => `<section class="pg dok ll">
    <header><div><b>Tracking Long Lead Material</b><span>${esc(D.meta.project)}</span></div>
      <div class="pr-meta"><span>Disiplin: <b>${esc(g)}</b></span><span>${esc(filt)} · status per ${fmt(tgl())}</span><span>Dicetak ${fmt(todayISO())}${D.sync.user ? ' oleh ' + esc(D.sync.user) : ''}</span></div></header>
    <table><thead><tr>${head.map(h => `<th>${esc(h)}</th>`).join('')}</tr></thead><tbody>${rows(rs)}</tbody></table>
    <p class="pr-n">${ring(rs)}. PO terlambat = belum PO dan lewat batas PO (rencana PO, atau tgl butuh di site − lead time); kritis = butuh di site ≤ 30 hari atau batas PO ≤ 14 hari. Sumber: log long lead Drive, dibaca ${LLD.at ? fmt(LLD.at.slice(0, 10)) : '–'}.</p>
  </section>`).join('');
  const judul = document.title;
  document.title = ['Long_Lead', o.dis || 'Semua', o.f ? LL_F[o.f].replace(/\W+/g, '_') : '', todayISO()].filter(Boolean).join('_');
  addEventListener('afterprint', () => { document.title = judul; $('#print').innerHTML = ''; }, { once: true });
  selesaiCetak();
}
function vLL() {
  llAuto();
  const id = drvId(llSrc());
  if (!id) return `<div class="card"><b>Tracking long lead material dari Google Drive</b><p class="muted" style="font-size:13px">Tempel link file log long lead (akses link: "Siapa saja yang memiliki link" – Pelihat). Aplikasi hanya membaca, tidak mengubah file.</p><button class="btn pri" data-act="ll-set">Atur link long lead</button></div>`;
  const q = UI.llQ.toLowerCase().trim(), all0 = llItems().map(x => ({ x, i: llInfo(x) }));
  const all = all0.filter(o => (!UI.llDis || llDisOf(o.x) === UI.llDis) && (!q || [o.x.nama, o.x.vendor, o.x.area, o.x.dis, o.x.pic, o.x.risk, o.x.act].join(' ').toLowerCase().includes(q)));
  const L = all.filter(o => llCocok(o, UI.llF)).sort((a, b) => LL_URUT.indexOf(a.i.st) - LL_URUT.indexOf(b.i.st) || String(a.i.ros || '9').localeCompare(String(b.i.ros || '9')));
  const disL = [...new Set(all0.map(o => llDisOf(o.x)))];
  const nF = f => all.filter(o => llCocok(o, f)).length;
  const tg = (l, v, cls = '') => v ? `<span${cls ? ` class="${cls}"` : ''}><small class="muted">${l}</small> ${v}</span>` : '';
  const chipS = (l, v) => v ? `<span class="chip ${/not|belum/i.test(v) ? '' : /approv|complete|deliver|done|selesai|issued|received/i.test(v) ? 'ok' : 'info'}">${esc(l)}: ${esc(v)}</span>` : '';
  return `<div class="card" style="padding:10px 14px"><div class="bar" style="margin:0">
      <span class="muted" style="font-size:13px">Log long lead Drive (hanya baca) · ${llBusy ? 'membaca…' : LLD.at && LLD.src === id ? `${llItems().length} item · dibaca ${new Date(LLD.at).toLocaleString('id-ID', { day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' })}` : 'belum dibaca'}</span>
      <span style="flex:1"></span><button class="btn sm" data-act="ll-load" ${llBusy ? 'disabled' : ''}>Muat ulang</button><button class="btn sm" data-act="ll-set">Link</button>${llItems()[0]?.link ? `<a class="btn sm" href="${esc(llItems()[0].link)}" target="_blank" rel="noopener">Buka log</a>` : ''}</div>
      ${LLD.err && LLD.src === id ? `<p class="muted" style="color:var(--bad);margin:6px 0 0;font-size:13px">${esc(LLD.err)}</p>` : ''}</div>
  <div class="grid kpis">
    <div class="kpi"><b>${all.length}</b><span>Item long lead</span></div>
    <div class="kpi ${nF('aksi') ? 'bad' : 'ok'}"><b>${nF('aksi')}</b><span>Perlu tindakan (kritis / terlambat)</span></div>
    <div class="kpi ${nF('ma') ? 'warn' : 'ok'}"><b>${nF('ma')}</b><span>Material approval belum approve</span></div>
    <div class="kpi ${nF('nopo') ? 'warn' : 'ok'}"><b>${nF('nopo')}</b><span>Belum PO</span></div>
    <div class="kpi info"><b>${nF('tiba')}</b><span>Sudah tiba di site</span></div>
  </div>
  <div class="prio-f">${Object.entries(LL_F).map(([k, l]) => `<button type="button" data-act="ll-f" data-v="${k}" class="${UI.llF === k ? 'on' : ''}">${l} <b>${nF(k)}</b></button>`).join('')}</div>
  <div class="prio-f kat-f"><button type="button" data-act="ll-dis" data-v="" class="${UI.llDis ? '' : 'on'}">Semua disiplin</button>${disL.map(d => `<button type="button" data-act="ll-dis" data-v="${esc(d)}" class="${UI.llDis === d ? 'on' : ''}">${esc(d)} <b>${all0.filter(o => llDisOf(o.x) === d).length}</b></button>`).join('')}</div>
  <div class="bar"><input type="search" data-ui="llQ" value="${esc(UI.llQ)}" placeholder="Cari material / vendor / area / PIC" style="flex:1;min-width:160px"><button type="button" class="btn" data-act="pdf-ll"${llItems().length ? '' : ' disabled'}>PDF / Excel</button></div>
  ${L.map(({ x, i }) => { const [cl, lb] = LL_ST[i.st], nm = x.nama.split('\n').map(s => s.trim()).filter(Boolean);
    return `<div class="item">
      <div class="hd"><div><span class="ttl">${esc(nm[0])}</span>${nm.length > 1 ? `<small class="muted" style="display:block">${esc(nm.slice(1).join(' · '))}</small>` : ''}
        <div class="sub">${esc([x.dis, x.area, x.vendor].filter(Boolean).join(' · '))}</div></div>
        <div class="chips"><span class="chip ${cl}">${lb}</span><span class="chip ${i.appr ? 'ok' : /submit/i.test(x.eng) ? 'info' : 'warn'}">MA: ${esc(x.eng || '–')}</span></div></div>
      <div class="meta">${tg('Qty', esc([x.qty, x.unit].filter(Boolean).join(' ')))}${tg('Lead time', esc(x.lead) + (i.lead ? ` (±${i.lead} hr)` : ''))}
        ${tg('Butuh di site', i.ros ? `<b>${fmt(i.ros)}</b>${i.tiba ? '' : ` · ${i.sisa < 0 ? `<span style="color:var(--bad)">lewat ${-i.sisa} hr</span>` : i.sisa + ' hr lagi'}`}${x.ros && !/^\d/.test(x.ros) ? ` <small class="muted">(${esc(x.ros.replace(/\s+/g, ' '))})</small>` : ''}` : esc(x.ros) || '<i>belum diisi</i>')}
        ${tg(i.bHitung ? 'Batas PO (ROS − lead)' : 'Rencana PO', i.bPO && !i.po ? `<b style="color:${i.bPO < tgl() ? 'var(--bad)' : 'inherit'}">${fmt(i.bPO)}</b>` : i.bPO ? fmt(i.bPO) : '')}
        ${tg('PO', i.po ? esc(i.poTgl ? fmt(i.poTgl) : i.po) : '<i>belum</i>')}${tg('Perkiraan tiba', i.tibaT ? fmt(i.tibaT) : esc(x.farr || x.parr).replace(/\n/g, ' '))}${x.ref ? tg('Ref', esc(x.ref)) : ''}</div>
      <div style="display:flex;align-items:center;gap:8px;margin:8px 0 0"><div class="bar-h" style="flex:1"><i style="width:${Math.min(x.prog || 0, 100)}%"></i></div><small class="muted">Progress ${x.prog == null ? '–' : pc(x.prog)}${x.dev ? ` · deviasi ${esc(x.dev)} hr` : ''}</small></div>
      <div class="chips" style="justify-content:flex-start;margin-top:6px">${chipS('Submittal', x.sub)}${chipS('Procurement', x.proc)}${chipS('Manufacturing', x.manu)}${chipS('Delivery', x.del)}</div>
      ${x.upd ? `<div class="fu">Update: ${esc(x.upd)}</div>` : ''}${x.risk ? `<div class="fu" style="border-color:var(--bad)">⚠ ${esc(x.risk)}</div>` : ''}${x.act ? `<div class="fu">→ ${esc(x.act)}</div>` : ''}
      ${x.pic || x.last || x.rem ? `<div class="muted" style="font-size:12px;margin-top:6px">${esc([x.pic && 'PIC/kontak: ' + x.pic, x.last && 'update ' + (llTgl(x.last) ? fmt(llTgl(x.last)) : x.last), x.rem].filter(Boolean).join(' · '))}</div>` : ''}
    </div>`; }).join('') || `<div class="empty">${llBusy ? 'Membaca log long lead dari Drive…' : 'Tidak ada item untuk filter ini.'}</div>`}
  <p class="muted" style="font-size:12px">Status dihitung terhadap tanggal hari ini: <b>PO terlambat</b> = belum PO dan sudah lewat rencana PO (atau tgl butuh di site − lead time); <b>Kritis</b> = butuh di site ≤ 30 hari atau batas PO ≤ 14 hari. "W1 Februari 2027" dibaca tanggal 1, W2 = tanggal 8, dst.</p>`;
}
function llSet() {
  openDlg(`<h3>Link log long lead (Google Drive)</h3>
    <label class="f">Link Google Sheet / file Excel di Drive<input type="text" name="src" inputmode="url" value="${esc(D.settings.llSrc || '')}" placeholder="${D.llSrc ? 'Dari Sheet tim (biarkan kosong)' : 'https://docs.google.com/spreadsheets/d/…'}"></label>
    <p class="muted" style="font-size:13px">Aplikasi hanya membaca file ini; tidak ada yang ditulis ke Drive. Akses link harus "Siapa saja yang memiliki link" (Pelihat).
    ${D.llSrc ? '<br>Link tim sudah diatur admin (push_master.py --ll-link). Isi di sini hanya bila ingin memakai link lain di perangkat ini.' : '<br>Agar berlaku untuk semua anggota tim, admin menjalankan push_master.py dengan --ll-link.'}</p>
    <div class="dlg-act"><span class="sp"></span><button value="cancel" class="btn" formnovalidate>Batal</button><button value="save" class="btn pri">Simpan</button></div>`, fd => {
    const s = String(fd.get('src') || '').trim();
    if (s && !drvId(s)) { toast('Link tidak dikenali – salin link dari tombol Bagikan'); return; }
    D.settings.llSrc = s; persist(); LLD = {}; render(); llLoad(true);
  });
}

function vSD() {
  return `<div class="vh"><h2>Sumber daya</h2></div>
  <div class="prio-f">${[['mp', 'Manpower'], ['dok', 'Dokumen'], ['ll', 'Long lead']].map(([k, l]) => `<button type="button" data-act="sd-tab" data-v="${k}" class="${UI.sd === k ? 'on' : ''}">${l}</button>`).join('')}</div>
  ${UI.sd === 'dok' ? vDok() : UI.sd === 'll' ? vLL() : vMp()}`;
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

const DOK_FD = { '': 'Semua', open: 'Belum approve', lewat3: 'Submit > 3 hari', cw: 'Awaiting feedback (CW)', own: 'Awaiting owner', lewat: 'Lewat target', Draft: 'Draft', Submitted: 'Submitted', Revisi: 'Revisi / Not Approved', Rejected: 'Rejected', Approved: 'Approved' };
const dokKat = x => x.jenis || 'Manual';
const dokCocok = (x, f) => !f || (DOK_FX[f] ? DOK_FX[f](x) : f === 'open' ? x.status !== 'Approved' : f === 'lewat' ? lewatDok(x) : x.status === f);
// daftar dokumen terfilter & terurut (dipakai tabel dan cetak PDF); o = { f, sc, jn, sb, q, pic }
function dokFilter(o) {
  const all0 = allDok(), q = String(o.q || '').toLowerCase().trim(), urut = new Map();
  all0.forEach(x => { const k = dokKat(x); if (!urut.has(k)) urut.set(k, urut.size); if (x.sub && !urut.has(k + '|' + x.sub)) urut.set(k + '|' + x.sub, urut.size); });
  return all0.filter(x => (!o.sc || x.scope === o.sc) && (!q || [x.nomor, x.nama, x.pic, x.sub, x.issue].join(' ').toLowerCase().includes(q)) && dokCocok(x, o.f)
      && (!o.jn || dokKat(x) === o.jn) && (!o.sb || x.sub === o.sb) && (!o.pic || (x.pic || '') === o.pic))
    .sort((a, b) => urut.get(dokKat(a)) - urut.get(dokKat(b)) || (urut.get(dokKat(a) + '|' + a.sub) ?? -1) - (urut.get(dokKat(b) + '|' + b.sub) ?? -1)
      || (lewatDok(b) - lewatDok(a)) || ((a.status === 'Approved') - (b.status === 'Approved')) || String(a.target || '9999').localeCompare(String(b.target || '9999')));
}
function vDok() {
  drvAuto();
  const f = UI.dokF, sc = UI.dokScope, jn = UI.dokJenis, sb = UI.dokSub, q = UI.dokQ.toLowerCase().trim(), all0 = allDok();
  const kat = x => x.jenis || 'Manual';
  const urut = new Map(); all0.forEach(x => { if (!urut.has(kat(x))) urut.set(kat(x), urut.size); if (x.sub && !urut.has(kat(x) + '|' + x.sub)) urut.set(kat(x) + '|' + x.sub, urut.size); });
  const scopeL = [...new Set([...D.disiplin, ...all0.map(x => x.scope).filter(Boolean)])];
  // filter selain kategori → dipakai untuk angka di chip kategori
  const base = all0.filter(x => (!sc || x.scope === sc) && (!q || [x.nomor, x.nama, x.pic, x.sub, x.issue].join(' ').toLowerCase().includes(q))
      && dokCocok(x, f));
  const all = jn ? all0.filter(x => kat(x) === jn) : all0; // KPI mengikuti kategori terpilih
  const subL = jn ? [...new Set(all.map(x => x.sub).filter(Boolean))] : [];
  const L = dokFilter({ f, sc, jn, sb, q });
  const katL = [...urut.keys()].filter(k => !k.includes('|'));
  const nKat = k => base.filter(x => kat(x) === k).length;
  const n = s => all.filter(x => x.status === s).length, lw = all.filter(lewatDok).length;
  const FD = DOK_FD;
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
    ${(() => { const a = all.filter(x => x.drive && x.status === 'Approved'), k = a.filter(x => !x.file).length; return a.length ? `<div class="kpi ${k ? 'warn' : 'ok'}"><b>${a.length - k}<small class="muted" style="font-size:13px"> / ${a.length}</small></b><span>Approved sudah ada link file</span></div>` : ''; })()}
  </div>
  <div class="prio-f">${Object.entries(FD).map(([k, l]) => `<button type="button" data-act="dok-f" data-v="${k}" class="${f === k ? 'on' : ''}">${l}${DOK_FX[k] ? ` <b>${all.filter(DOK_FX[k]).length}</b>` : ''}</button>`).join('')}</div>
  <div class="prio-f kat-f"><button type="button" data-act="dok-kat" data-v="" class="${jn ? '' : 'on'}">Semua kategori <b>${base.length}</b></button>${katL.map(k => `<button type="button" data-act="dok-kat" data-v="${esc(k)}" class="${jn === k ? 'on' : ''}">${esc(k)} <b>${nKat(k)}</b></button>`).join('')}</div>
  <div class="bar">${subL.length ? `<select data-ui="dokSub" style="width:auto">${opt('', sb, 'Semua sub-kategori')}${subL.map(d => opt(d, sb)).join('')}</select>` : ''}
    <select data-ui="dokScope" style="width:auto">${opt('', sc, 'Semua scope')}${scopeL.map(d => opt(d, sc)).join('')}</select>
    <input type="search" data-ui="dokQ" value="${esc(UI.dokQ)}" placeholder="Cari nomor / nama / PIC" style="flex:1;min-width:160px">
    <button type="button" class="btn" data-act="pdf-dok">PDF / Excel</button></div>
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
          x.drive ? [tl('CW', x.cw), tl('Owner', x.own)].filter(Boolean).join('<br>') : tl('Approve', x.tglApprove)].filter(Boolean).join('<br>');
        h += `<tr id="d-${esc(x.id)}" data-act="dok-info" data-id="${esc(x.id)}">
          <td><span class="code">${esc(x.nomor || '(tanpa nomor)')}${x.rev ? ' · Rev ' + x.rev : ''}</span>${x.file ? ' <span class="chip ok" title="Link dokumen tersimpan">📎 file</span>' : x.drive && x.status === 'Approved' ? ' <span class="chip warn" title="Approved tapi kolom Approved Doc. Link masih kosong">link file?</span>' : ''}<br>${esc(x.nama)}${x.issue ? `<small class="dok-rem">${esc(x.issue)}</small>` : ''}</td>
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
          + row('Respon Owner', x.own && `${fmt(x.own)} <small class="muted">· ${days(x.tglSubmit || x.own, x.own)} hari setelah submit</small>`)
          + (!x.cw && !x.own && x.respon ? row('Respon', fmt(x.respon)) : '')
        : row('Target approve', fmt(x.target)) + row('Submit', fmt(x.tglSubmit)) + row('Approved', x.tglApprove && fmt(x.tglApprove))}
      ${row('Remarks / issue', x.issue && esc(x.issue).replace(/\n/g, '<br>'))}
      ${row('Task LPS', tasks.map(k => `<a href="#ready" data-act="goto-ready" data-kode="${esc(k)}" class="chip info" style="text-decoration:none">${esc(k)}</a>`).join(' '))}
    </table>
    ${x.riwayat ? `<details open style="margin-top:8px"><summary class="muted" style="font-size:13px">Riwayat revisi</summary><pre class="riwayat">${esc(x.riwayat)}</pre></details>` : ''}
    ${x.drive ? `<p class="muted" style="font-size:12px">Dari log Google Drive (hanya baca). Perubahan dilakukan di file Drive.${x.file ? '' : ' Link file dokumen belum diisi di kolom <b>Approved Doc. Link</b>.'}</p>` : ''}
    <div class="dlg-act" style="flex-wrap:wrap">${next.map(([v, l]) => `<button type="button" class="btn sm ${v === 'Approved' ? 'pri' : v === 'Rejected' ? 'danger' : ''}" data-act="dok-st" data-id="${esc(x.id)}" data-v="${v}">${esc(l)}</button>`).join('')}
      <span class="sp"></span>${x.file ? `<a class="btn pri" href="${esc(x.file)}" target="_blank" rel="noopener">Buka dokumen${x.fileRev != null && x.fileRev !== x.rev ? ' (Rev ' + x.fileRev + ')' : ''}</a>` : ''}${x.link ? `<a class="btn" href="${esc(x.link)}" target="_blank" rel="noopener">${x.drive ? 'Buka log' : 'Buka file'}</a>` : ''}${x.drive ? '' : `<button type="button" class="btn" data-act="dok-edit" data-id="${esc(x.id)}">Ubah</button>`}
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
const vh = (title, r) => `<div class="vh"><h2>${esc(title)}</h2><button class="btn sm" data-act="pdf" data-r="${r}">PDF / Excel</button></div>`;
const picOf = s => String(s || '').trim() || '(tanpa PIC)';
const prioTxt = t => [t.prio ? t.prio.toUpperCase() : '', koordList(t).length ? '↔ ' + koordList(t).join(', ') : '', t.followup].filter(Boolean).join('\n');
const progTxt = t => { if (!t) return ''; const p = prog(t.kode), r = rantai(t), a = t.bl && tlById(t.bl), sl = tlSlip(t);
  return [`Mingguan ${pc(p || 0)}`, r ? `Total ${pc(r.total)} (minggu ${r.idx + 1}/${r.M.length}, porsi ${pc(r.eff[r.idx])})` : '',
    t.disepakati === 'Ya' && stP(t) ? stP(t) : '', a ? `Timeline ${fmt(a.mulai)} – ${fmt(a.selesai)}${sl > 0 ? ` (+${sl} hr)` : ''}` : ''].filter(Boolean).join('\n'); };
const rentang = t => t.mulai || t.selesai ? `${fmt(t.mulai)} – ${fmt(t.selesai)}` : '–';
/* Tiap laporan: judul, kolom, dan baris {pic, cells}. Make Ready / Komitmen / Harian = PIC task; Issue = PIC kendala. */
const REPORTS = {
  komit: { title: 'Komitmen Mingguan', week: true,
    head: ['Kode', 'Prioritas · koordinasi', 'Pekerjaan', 'Uraian', 'Lokasi / Area', 'Rencana', 'Disepakati', 'Progress', 'Alasan / catatan rapat'],
    rows: o => D.tasks.filter(inWR(o.week, o.week2)).filter(prioOK).sort((a, b) => (rangeOn(o.week, o.week2) || o.week === '' ? a.minggu - b.minggu : 0) || byPrio(a, b) || a.kode.localeCompare(b.kode)).map(t => ({ pic: picOf(t.pic),
      cells: [t.kode + (o.week === '' || rangeOn(o.week, o.week2) ? '\nW' + t.minggu : ''), prioTxt(t), t.pekerjaan, t.uraian, [t.lokasi, t.area].filter(Boolean).join(' · '), rentang(t), t.disepakati || '–', progTxt(t), t.alasan] })) },
  ready: { title: 'Make Ready – Cek Kendala', week: true,
    head: ['Kode', 'Prioritas · koordinasi', 'Uraian', 'Lokasi', 'Rencana', 'Status kendala', 'Kendala pending (issue · PIC · target)', 'Progress'],
    rows: o => D.tasks.filter(t => t.disepakati === 'Ya' && inWR(o.week, o.week2)(t) && prioOK(t)).sort((a, b) => (rangeOn(o.week, o.week2) || o.week === '' ? a.minggu - b.minggu : 0) || byPrio(a, b) || a.kode.localeCompare(b.kode)).map(t => ({ pic: picOf(t.pic),
      cells: [t.kode + (o.week === '' || rangeOn(o.week, o.week2) ? '\nW' + t.minggu : ''), prioTxt(t), t.uraian, t.lokasi, rentang(t), stK(t),
        [...KK.filter(k => t.k[k].st === 'Pending').map(k => { const c = t.k[k]; return `${KAT[k]}: ${c.issue || '?'}${c.detail ? ' – ' + c.detail : ''} · ${c.pic || 'PIC?'} · ${fmt(c.target)}`; }),
          ...(belumCek(t).length ? ['Belum dicek: ' + belumCek(t).map(k => KAT[k]).join(', ')] : [])].join('\n') || 'Semua OK',
        progTxt(t) || '–'] })) },
  harian: { title: 'Daily Commitment', date: true,
    head: ['Kode', 'Uraian / lokasi', 'Target hari ini', 'Bobot', 'Actual', '% tercapai', 'Tercapai?', 'Progress kumulatif', 'Progress task', 'Penyebab tidak tercapai', 'Catatan'],
    rows: o => D.daily.filter(x => x.tgl === o.tgl).map(x => { const t = D.tasks.find(t => t.kode === x.kode); return (!UI.fPrio || prioOK(t)) && { pic: picOf(t?.pic),
      cells: [x.kode + (t?.prio ? '\n' + t.prio.toUpperCase() : ''), [t?.uraian, t?.lokasi].filter(Boolean).join(' · '), x.target, pc(x.bobot), x.capaian == null ? '' : pc(aktual(x)), x.capaian == null ? '' : pc(x.capaian), tercapai(x),
        x.capaian == null ? '' : pc(progKum(x)), progTxt(t), x.penyebab, x.catatan] }; }) },
  issue: { title: 'Follow Up Issue / Kendala', week: true,
    head: ['Kode', 'Task', 'Kategori', 'Issue', 'Detail', 'Target', 'Sisa hari', 'Status'],
    rows: o => issueList(o.week, o.week2).filter(x => prioOK(x.t)).sort((a, b) => byPrio(a.t, b.t)).map(x => ({ pic: picOf(x.c.pic || x.t.pic),
      cells: [x.t.kode + (x.t.prio ? '\n' + x.t.prio.toUpperCase() : ''), [x.t.uraian, x.t.lokasi].filter(Boolean).join(' · '), KAT[x.k], x.c.issue, x.c.detail, fmt(x.c.target), x.sisa ?? '–', x.st] })) },
};
function pdfOpts(r) { return { week: UI.week, week2: UI.view === 'komit' && kRange() ? UI.kTo : '', tgl: UI.harianTgl || tgl(), pic: '' , r }; }
function pdfForm(r) {
  const R = REPORTS[r], o = pdfOpts(r);
  const pics = [...new Set(R.rows(o).filter(Boolean).map(x => x.pic))].sort();
  return `<h3>PDF ${esc(R.title)}</h3>
  ${R.week ? `<div class="frow"><label class="f">Minggu${''}<select name="week">${opt('', o.week, 'Semua minggu')}${D.meta.weeks.map(x => opt(x.w, o.week, `W${x.w} · ${x.label}`)).join('')}</select></label>
    <label class="f">s/d minggu<select name="week2">${opt('', o.week2, '— hanya minggu itu —')}${D.meta.weeks.map(x => opt(x.w, o.week2, `W${x.w} · ${x.label}`)).join('')}</select></label></div>` : ''}
  ${R.date ? `<label class="f">Tanggal<input type="date" name="tgl" value="${o.tgl}"></label>` : ''}
  <label class="f">PIC<select name="pic">${opt('', '', 'Semua PIC – 1 halaman per PIC')}${pics.map(p => opt(p, '')).join('')}</select></label>
  <p class="muted" style="font-size:13px">Daftar PIC mengikuti ${R.date ? 'tanggal' : 'minggu'} yang sedang dibuka. Di jendela cetak pilih <b>Simpan sebagai PDF</b>, lalu kirim ke PIC.${UI.fPrio ? ` Filter prioritas aktif: <b>${esc(FPRIO[UI.fPrio])}</b>.` : ''}${r === 'issue' ? ' Issue dikelompokkan per PIC kendala.' : ''}</p>
  <div class="dlg-act"><button value="cancel" class="btn" formnovalidate>Batal</button><button value="xls" class="btn">Excel</button><button value="save" class="btn pri">Buat PDF</button></div>`;
}
function printReport(r, o) {
  const R = REPORTS[r], rows = R.rows(o).filter(x => x && (!o.pic || x.pic === o.pic));
  if (!rows.length) { toast('Tidak ada data untuk dicetak'); return; }
  const groups = {};
  rows.forEach(x => (groups[x.pic] ||= []).push(x.cells));
  const wo = o.week === '' ? null : weekObj(o.week);
  const per = R.date ? fmt(o.tgl) + ` (W${weekOf(o.tgl)})` : perLabel(o.week, o.week2);
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
  document.title = [R.title.split(' ')[0], o.pic || 'Semua-PIC', R.date ? o.tgl : wo ? 'W' + wo.w + (rangeOn(o.week, o.week2) ? '-W' + o.week2 : '') : 'semua'].join('_').replace(/[\\/:*?"<>|\s]+/g, '-'); // nama file PDF
  addEventListener('afterprint', () => { document.title = judul; $('#print').innerHTML = ''; }, { once: true });
  selesaiCetak();
}

/* ---------- PDF Jadwal / Lookahead: 1 minggu atau beberapa minggu ke depan, bar per minggu, opsional task komitmen ---------- */
function pdfDokForm() {
  const all = allDok(), katL = [...new Set(all.map(dokKat))], scL = [...new Set([...D.disiplin, ...all.map(x => x.scope).filter(Boolean)])];
  const picL = [...new Set(all.map(x => x.pic).filter(Boolean))].sort();
  return `<h3>PDF Tracking Dokumen</h3>
  <p class="muted" style="font-size:13px;margin-top:0">Terisi sesuai filter di layar; ubah bila perlu.</p>
  <div class="frow">
    <label class="f">Status / filter<select name="f">${Object.entries(DOK_FD).map(([k, l]) => opt(k, UI.dokF, l)).join('')}</select></label>
    <label class="f">Kategori<select name="jn">${opt('', UI.dokJenis, 'Semua kategori')}${katL.map(k => opt(k, UI.dokJenis)).join('')}</select></label>
  </div>
  <div class="frow">
    <label class="f">Scope<select name="sc">${opt('', UI.dokScope, 'Semua scope')}${scL.map(k => opt(k, UI.dokScope)).join('')}</select></label>
    <label class="f">PIC<select name="pic">${opt('', '', 'Semua PIC')}${picL.map(k => opt(k, '')).join('')}</select></label>
  </div>
  <div class="frow">
    <label class="f">Cari<input type="text" name="q" value="${esc(UI.dokQ)}" placeholder="nomor / nama / remarks"></label>
    <label class="f">Susunan<select name="grp">${opt('kat', 'kat', '1 halaman per kategori')}${opt('satu', 'kat', 'Satu tabel')}</select></label>
  </div>
  <p class="muted" style="font-size:13px">Kertas A4 mendatar. Di jendela cetak pilih <b>Simpan sebagai PDF</b>.</p>
  <div class="dlg-act"><button value="cancel" class="btn" formnovalidate>Batal</button><button value="xls" class="btn">Excel</button><button value="save" class="btn pri">Buat PDF</button></div>`;
}
function printDok(o) {
  const L = dokFilter({ ...o, sb: o.jn && o.jn === UI.dokJenis ? UI.dokSub : '' });
  if (!L.length) { toast('Tidak ada dokumen untuk filter ini'); return; }
  const groups = o.grp === 'kat' ? [...new Set(L.map(dokKat))].map(k => [k, L.filter(x => dokKat(x) === k)]) : [[o.jn || 'Semua kategori', L]];
  const filt = [o.f && DOK_FD[o.f], o.sc && 'Scope ' + o.sc, o.pic && 'PIC ' + o.pic, o.q && `"${o.q}"`].filter(Boolean).join(' · ') || 'Semua dokumen';
  const head = ['No', 'Document number', 'Rev', 'Description', 'Rencana submit', 'Submit aktual', 'Respon CW', 'Respon Owner', 'Umur (hari)', 'Status', 'Remarks', 'PIC'];
  const rows = ds => { let n = 0, sub = null; return ds.map(x => {
    const u = umurDok(x), sh = x.sub && x.sub !== sub ? `<tr class="sb"><td colspan="${head.length}">${esc(x.sub)}</td></tr>` : ''; sub = x.sub || sub;
    return sh + `<tr><td>${++n}</td><td>${esc(x.nomor || '–')}</td><td>${x.rev ?? 0}</td><td>${esc(x.nama)}</td><td>${fmt(x.target) || ''}</td><td>${fmt(x.tglSubmit) || ''}</td>
      <td>${x.drive ? fmt(x.cw) || '' : ''}</td><td>${x.drive ? fmt(x.own) || '' : fmt(x.tglApprove) || ''}</td><td class="u${u > 14 ? ' bad' : u > 7 ? ' warn' : ''}">${u ?? ''}</td>
      <td>${esc(x.stAsli || x.status)}${lewatDok(x) ? ' <b>(lewat)</b>' : ''}</td><td class="rm">${esc(x.issue || '')}</td><td>${esc(x.pic || '')}</td></tr>`; }).join(''); };
  const ring = ds => { const c = f => ds.filter(DOK_FX[f]).length; return `${ds.length} dokumen · ${ds.filter(x => x.status === 'Approved').length} approved · ${c('cw')} awaiting CW · ${c('own')} awaiting owner · ${c('lewat3')} submit > 3 hari`; };
  $('#print').innerHTML = groups.map(([g, ds]) => `<section class="pg dok">
    <header><div><b>Tracking Dokumen</b><span>${esc(D.meta.project)}</span></div>
      <div class="pr-meta"><span>Kategori: <b>${esc(g)}</b></span><span>${esc(filt)}</span><span>Dicetak ${fmt(todayISO())}${D.sync.user ? ' oleh ' + esc(D.sync.user) : ''}</span></div></header>
    <table><thead><tr>${head.map(h => `<th>${esc(h)}</th>`).join('')}</tr></thead><tbody>${rows(ds)}</tbody></table>
    <p class="pr-n">${ring(ds)} · umur = hari sejak submit aktual untuk dokumen yang belum ada keputusan</p>
  </section>`).join('');
  const judul = document.title;
  document.title = ['Tracking_Dokumen', o.jn || 'Semua', o.f ? DOK_FD[o.f].replace(/\W+/g, '_') : '', todayISO()].filter(Boolean).join('_');
  addEventListener('afterprint', () => { document.title = judul; $('#print').innerHTML = ''; }, { once: true });
  selesaiCetak();
}
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
  <div class="dlg-act"><button value="cancel" class="btn" formnovalidate>Batal</button><button value="xls" class="btn">Excel</button><button value="save" class="btn pri">Buat PDF</button></div>`;
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
      <td>${fmt(t.mulai)}</td><td>${fmt(t.selesai)}</td><td>W${t.minggu}</td><td>${esc([t.disepakati === 'Ya' ? stP(t) : 'Belum sepakat', t.pic].filter(Boolean).join(' · '))}<br><small>${esc(progTxt(t).split('\n').filter(x => !/^(FINISH|ON PROGRESS|LATE)/.test(x)).join(' · '))}</small></td>
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
  selesaiCetak();
}

/* ---------- export: laporan yang disusun di #print dicetak (PDF) atau diubah jadi .xlsx (tombol Excel di dialog) ---------- */
let XLS = false;
function selesaiCetak() {
  if (!XLS) { setTimeout(() => print(), 60); return; } // beri waktu dialog tertutup
  try { xlsDariPrint(document.title); } finally { dispatchEvent(new Event('afterprint')); } // afterprint = kembalikan judul & kosongkan #print
}
const xNum = v => /^-?\d{1,12}(\.\d+)?$/.test(v) ? +v : v;
function xlsDariPrint(nama) {
  const teks = el => { const c = el.cloneNode(true); c.querySelectorAll('br').forEach(b => b.replaceWith('\n'));
    return c.textContent.replace(/[ \t\r]*\n[ \t\r]*/g, '\n').replace(/[ \t]{2,}/g, ' ').trim(); };
  const sheets = [...$('#print').querySelectorAll('section')].map((sec, i) => {
    const hd = sec.querySelector('header'), meta = [...(hd ? hd.querySelectorAll('.pr-meta span') : [])].map(teks);
    const nm = (meta.find(m => /^(PIC|Disiplin|Kategori):/i.test(m)) || '').replace(/^[^:]+:\s*/, '') || 'Sheet' + (i + 1);
    const rows = [{ s: 'title', c: [[hd?.querySelector('b')?.textContent, hd?.querySelector('div > span')?.textContent].filter(Boolean).join(' · ')] }, { s: 'meta', c: [meta.join(' · ')] }];
    let freeze = 0;
    sec.querySelectorAll('table tr').forEach(tr => {
      const tds = [...tr.children];
      if (tr.closest('thead')) { rows.push({ s: 'head', c: tds.map(teks) }); freeze = freeze || rows.length; return; }
      const grp = (tds.length === 1 && tds[0].colSpan > 1) || tr.classList.contains('pj') || tr.classList.contains('sb');
      rows.push({ s: grp ? 'grp' : 'cell', c: tds.map(td => td.classList.contains('g') ? (td.classList.contains('on') ? '■' : '') : xNum(teks(td))) });
    });
    sec.querySelectorAll('.pr-n').forEach(x => rows.push({ s: 'note', c: [teks(x)] }));
    return { name: nm, rows, freeze };
  });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(xlsxBlob(sheets)); a.download = (nama || 'LPS').replace(/[\\/:*?"<>|]+/g, '-') + '.xlsx';
  document.body.append(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 5000);
  toast(`Excel dibuat: ${sheets.length} sheet`);
}

/* ---------- dialog ---------- */
let dlgSubmit = null;
function openDlg(html, onSave) {
  $('#dlg-form').innerHTML = html;
  roLock($('#dlg-form'));
  dlgSubmit = onSave;
  if (!$('#dlg').open) $('#dlg').showModal();
}
$('#dlg').addEventListener('close', () => {
  const dlg = $('#dlg');
  if ((dlg.returnValue === 'save' || dlg.returnValue === 'xls') && dlgSubmit) { XLS = dlg.returnValue === 'xls'; try { dlgSubmit(new FormData($('#dlg-form'))); } finally { XLS = false; } }
  dlg.returnValue = ''; dlgSubmit = null;
});

function editTask(t, isNew) {
  openDlg(taskForm(t, isNew), fd => {
    const v = Object.fromEntries(fd.entries());
    v.minggu = +v.minggu;
    v.porsi = nOr(v.porsi); if (v.porsi != null) v.porsi = Math.max(0, Math.min(100, v.porsi));
    v.induk = v.induk === v.kode ? '' : (v.induk || '');
    v.koord = fd.getAll('koord').join(','); // checkbox: bisa lebih dari satu
    delete v.tpl; if (!v.bl) delete v.bl; // bl = id aktivitas timeline bila memakai template
    if (isNew && D.tasks.some(x => x.kode === v.kode)) { toast('Kode ' + v.kode + ' sudah dipakai'); return; }
    const p = D.pekerjaan.find(x => x.nama === v.pekerjaan);
    if (p && !v.disiplin) v.disiplin = p.disiplin;
    if (isNew) D.tasks.push(Object.assign(t, v));
    else { const o = D.tasks.find(x => x.kode === t.kode); if (v.disepakati === 'Ya' && o.disepakati !== 'Ya') wajibCek(o); Object.assign(o, v); }
    if (isNew && v.disepakati === 'Ya') wajibCek(t);
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
  if (RO && !RO_OK.has(a)) { e.preventDefault(); toast('Mode SPV: hanya lihat'); return; }
  // tombol di dialog detail dokumen: tutup detail dulu (dialog berikutnya memakai #dlg yang sama)
  if (['dok-st', 'dok-edit', 'goto-ready'].includes(a) && el.closest('#dlg')?.open) $('#dlg').close();
  switch (a) {
    case 'settings':
      openDlg(settingsForm(), fd => {
        D.settings.tglMode = fd.get('tglMode');
        D.settings.tglFixed = fd.get('tglFixed') || null;
        if (!RO) D.meta.targetPPC = Math.max(0, Math.min(100, +fd.get('targetPPC') || 80)) / 100;
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
    case 'll-set': llSet(); break;
    case 'll-load': llLoad(true); break;
    case 'pdf-ll': openDlg(pdfLLForm(), fd => printLL({ f: fd.get('f') || '', dis: fd.get('dis') || '', q: String(fd.get('q') || '').trim(), grp: fd.get('grp'), pic: fd.has('pic') })); break;
    case 'll-f': UI.llF = el.dataset.v; render(); break;
    case 'll-dis': UI.llDis = el.dataset.v; render(); break;
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
      openDlg(pdfForm(r), fd => printReport(r, { week: fd.has('week') ? fd.get('week') : UI.week, week2: fd.get('week2') || '', tgl: fd.get('tgl') || UI.harianTgl || tgl(), pic: fd.get('pic') || '' }));
      break;
    }
    case 'pdf-dok': openDlg(pdfDokForm(), fd => printDok({ f: fd.get('f') || '', jn: fd.get('jn') || '', sc: fd.get('sc') || '', pic: fd.get('pic') || '', q: String(fd.get('q') || '').trim(), grp: fd.get('grp') })); break;
    case 'pdf-jadwal': openDlg(pdfJadwalForm(), fd => printJadwal({ from: fd.get('from'), look: fd.get('look'), dis: fd.get('dis'), grp: fd.get('grp'), task: fd.has('task') })); break;
    case 'tl-pick': tlPick(el.dataset.w || tlWeek(), el.dataset.pek || ''); break;
    case 'tl-pek-all': e.preventDefault(); tlPick($('#tl-w').value); break;
    case 'tl-ubah': e.preventDefault(); tlUbah(el.dataset.id); break;
    case 'jd-ubah': jdUbah(el.dataset.pek); break;
    case 'jd-reset': D.jdEdit = D.jdEdit.filter(x => x.id !== el.dataset.pek); dlgSubmit = null; $('#dlg').close(); save(); render(); toast('Tanggal kembali sesuai master'); break;
    case 'tl-reset': D.tlEdit = D.tlEdit.filter(x => x.id !== el.dataset.id); dlgSubmit = null; $('#dlg').close(); save(); render(); toast('Tanggal kembali sesuai PDF'); break;
    case 'tl-add': { const dt = el.closest('details[data-pek]'); if (dt) UI.jdBuka[dt.dataset.pek] = true; const t = tlBuat([el.dataset.id], +el.dataset.w)[0]; if (t) { render(); editTask({ ...t }, false); } break; }
    case 'tl-all': document.querySelectorAll('#tl-pick .pick-i:not([hidden]) input:not(:disabled)').forEach(c => { c.checked = true; }); tlFilter(); break;
    case 'tl-none': document.querySelectorAll('#tl-pick input:checked').forEach(c => { c.checked = false; }); tlFilter(); break;
    case 'tl-manual': { const w = $('#tl-w')?.value; if (w != null) UI.week = w; newTask($('#tl-pick')?.dataset.pek || undefined); break; } // isi dialog diganti form task biasa
    case 'tl-goto': UI.week = el.dataset.w; UI.wPilih = true; if (UI.view === 'komit') render(); break;
    case 'rantai': e.preventDefault(); rantaiDlg(task); break;
    case 'kw': e.preventDefault(); UI.week = el.dataset.v; UI.kTo = ''; UI.kRgBuka = false; UI.wPilih = true; render(); break;
    case 'k-rg': { if (UI.week === '') UI.week = String(weekOf(tgl())); UI.kRgBuka = true; if (!kRange()) { const n = +UI.week + 3; UI.kTo = String(Math.min(n, D.meta.weeks.at(-1).w)); if (!kRange()) UI.kTo = ''; } UI.wPilih = true; render(); break; }
    case 'lanjut': { const n = lanjutkan(task); if (n) { save(); render(); toast(`${task.kode} dilanjutkan ke W${n.minggu} sebagai ${n.kode} – sepakati di rapat`); } break; }
    case 'lanjut-semua': {
      const w = +el.dataset.w, L = D.tasks.filter(t => +t.minggu === w && perluLanjut(t));
      if (!L.length || !confirm(`Buat ${L.length} task lanjutan di W${w + 1}?\nTask W${w} tetap tercatat apa adanya (PPC tidak berubah).`)) break;
      L.forEach(lanjutkan); UI.week = String(w + 1); UI.wPilih = true; save(); render();
      toast(`${L.length} task dilanjutkan ke W${w + 1} – status disepakati dikosongkan untuk rapat`); break;
    }
    case 'task-new': newTask(el.dataset.pek); break;
    case 'task-edit': editTask({ ...task }, false); break;
    case 'task-del':
      if (confirm(`Hapus task ${kode}? Daily log task ini juga dihapus.`)) {
        D.tasks = D.tasks.filter(t => t.kode !== kode);
        D.daily = D.daily.filter(x => x.kode !== kode);
        save(); $('#dlg').close(); render(); toast('Task dihapus');
      }
      break;
    case 'sepakat': {
      const lama = task.disepakati;
      task.disepakati = lama === el.dataset.v ? '' : el.dataset.v; task.inLPS = task.disepakati === 'Ya';
      if (task.disepakati === 'Ya' && lama !== 'Ya') { wajibCek(task); toast(`${task.kode} disepakati – cek 4 kendala di Make Ready`); }
      save(); render(); break;
    }
    case 'kst': task.k[el.dataset.k].st = el.dataset.v; save(); render(); break;
    case 'daily-new': newDaily(el.dataset.kode, el.dataset.tgl); break;
    case 'h-tab': UI.hTab = el.dataset.v; render(); break;
    case 'log-all': UI.lgAll = el.checked; render(); break;
    case 'log-buka': UI.hTab = 'log'; if (location.hash !== '#harian') location.hash = 'harian'; else render(); break;
    case 'waste-buka': UI.wBuka[el.dataset.id] = true; render(); break;
    case 'waste': { // pilih/lepas jenis waste; "Bukan waste" berdiri sendiri
      const x = D.daily.find(d => d.id === el.dataset.id), k = el.dataset.v; if (!x) break;
      let L = wasteL(x);
      L = k === 'X' ? (L.includes('X') ? [] : ['X']) : L.includes(k) ? L.filter(v => v !== k) : [...L.filter(v => v !== 'X'), k];
      x.waste = L.join(','); if (!L.length) UI.wBuka[x.id] = true;
      save(); render(); break;
    }
    case 'waste-saran': {
      let n = 0;
      D.daily.forEach(x => { if (tercapai(x) === 'Tidak' && !wasteL(x).length && saranWaste(x.penyebab)) { x.waste = saranWaste(x.penyebab); n++; } });
      save(); render(); toast(`${n} janji diisi jenis waste dari penyebabnya`); break;
    }
    case 'log-goto': UI.hTab = ''; UI.harianTgl = el.dataset.tgl; D.settings.hMode = '1'; persist(); UI.focus = `h-${el.dataset.tgl}-${el.dataset.kode}`; render(); break;
    case 'pdf-log': openDlg(pdfLogForm(), fd => printLog({ from: fd.get('from'), to: fd.get('to'), peny: fd.get('peny') || '', dis: fd.get('dis') || '', pic: fd.get('pic') || '', waste: fd.get('waste') || '', q: String(fd.get('q') || '').trim(), all: fd.has('all'), grp: fd.get('grp') })); break;
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
    case 'r-tgl': e.preventDefault(); UI.rBuka[kode] = !rBuka(task); render(); break;
    case 'r-all': planTasks().forEach(t => { UI.rBuka[t.kode] = el.dataset.v === '1'; }); render(); break;
    case 'goto-ready': UI.rBuka[kode] = true; UI.focus = 'r-' + kode; if (UI.view === 'ready') render(); break;
    case 'goto-komit': UI.focus = 'k-' + kode; if (UI.view === 'komit') render(); break;
    case 'exp-json': download(`LPS_backup_${tgl()}.json`, JSON.stringify({ ...D, sync: undefined }, null, 1), 'application/json'); break; // tanpa kode akses
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
  if (RO) return; // SPV: tidak ada isian yang mengubah data
  if (el.dataset.mp) { // aktual manpower per aktivitas per tanggal; kosong = hapus
    const d = UI.mpTgl || tgl(), id = mpId(d, el.dataset.mp), v = el.value === '' ? null : Math.max(0, Math.round(+el.value));
    const r = D.manpower.find(x => x.id === id);
    if (v == null) D.manpower = D.manpower.filter(x => x !== r);
    else if (r) r.aktual = v; else D.manpower.push({ id, tgl: d, item: el.dataset.mp, aktual: v, catatan: '' });
    save(); render(); return;
  }
  if (el.id === 'tl-w') { tlPick(el.value, $('#tl-pick')?.dataset.pek || ''); return; }
  if (el.id === 'tf-tpl') { tplFill(el.value); return; }
  if (el.name === 'minggu' && $('#tf-tpl') && el.closest('#dlg-form')) { // ganti minggu → daftar template & tanggal ikut minggu baru
    const cur = $('#tf-tpl').value; $('#tf-tpl').innerHTML = tplOpts(el.value, el.form.elements.pekerjaan.value, cur);
    tplFill($('#tf-tpl').value); return; }
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
    else { setPath(x, path, v); if (path === 'penyebab' && !wasteL(x).length) x.waste = saranWaste(v); } // saran TIMWOODS dari penyebab
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
  if (k !== 'jSearch' && k !== 'dokQ' && k !== 'lgQ' && k !== 'llQ') return;
  UI[k] = e.target.value;
  const pos = e.target.selectionStart;
  render();
  const s = document.querySelector(`[data-ui="${k}"]`); s.focus(); s.setSelectionRange(pos, pos);
});
document.addEventListener('toggle', e => { const d = e.target; if (d.matches?.('details.tl-jd')) UI.jdBuka[d.dataset.pek] = d.open; }, true);
$('#f-week').addEventListener('change', e => { UI.week = e.target.value; UI.kTo = ''; UI.kRgBuka = false; UI.wPilih = true; render(); });

function route() {
  const v = location.hash.slice(1);
  UI.view = VIEWS.includes(v) ? v : 'dash';
  // Komitmen dibuka pertama kali → otomatis minggu ini (selama minggu belum dipilih manual)
  if (UI.view === 'komit' && UI.week === '' && !UI.wPilih && D?.meta?.weeks?.length) { const cw = weekOf(tgl()); if (weekObj(cw)) UI.week = String(cw); }
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
