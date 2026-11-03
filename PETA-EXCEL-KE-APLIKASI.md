# Peta Excel → Aplikasi PWA "LPS Batik"

Sumber: `LPS - WEEKLY - BATIK PROJECT - INTERNAL.xlsx` (15 sheet, metode Last Planner System).
Tujuan: aplikasi ringan (PWA, bisa di-install di HP, jalan offline) untuk monitoring harian/mingguan di lapangan.

## 1. Alur bisnis (tetap sama dengan Excel)

```
MASTER (jadwal & bobot)
   └─► Daftar Pekerjaan (header x.x)  ──►  Lookahead (pekerjaan aktif per minggu)
            └─► Komitmen (task detail + "Disepakati" di rapat mingguan)
                     └─► LPS Input / Make Ready (cek 4 kendala → READY / PENDING)
                              └─► Daily Log (janji pagi, laporan sore, progress %)
                                       └─► PPC Harian / Mingguan, Pareto penyebab, Board 1–4, Follow Up Issue
```

Di aplikasi, sheet **input** menjadi layar/form, sheet **otomatis** (board, rekap) menjadi tampilan yang dihitung langsung di HP.

## 2. Peta sheet → modul aplikasi

| Sheet Excel | Jenis | Modul / layar di aplikasi | Keterangan |
|---|---|---|---|
| MASTER | referensi | `meta.weeks`, `meta.scurve`, bobot di `pekerjaan` | Hanya dibaca saat impor. Baris 8–10 = nomor & tanggal minggu, baris 2696–2697 = kurva-S rencana. |
| Area Gedung (`tblMapArea`) | referensi | `areaMap`, `areas` | Dropdown Area Gedung di form Komitmen. |
| Daftar Pekerjaan (`tblPekerjaan`) | referensi + otomatis | **Jadwal** | Daftar 126 header pekerjaan, status jadwal & jumlah task dihitung di app. |
| LookHead Plan | otomatis | **Jadwal** → filter "Minggu ke … + N minggu" | Tidak disimpan, dihitung dari tanggal mulai/selesai. |
| Komitmen (`tblKomit`) | **input** | **Komitmen** (form task + tombol Disepakati Ya/Tidak) | Kolom A–M. |
| LPS Input (`tblLPS`) | **input** | **Make Ready** (4 kartu kendala per task) | Digabung dengan Komitmen dalam satu objek `task` (kunci = Kode Task). |
| Daily Log (`tblDaily`) | **input** | **Harian** (janji pagi + laporan sore) | Satu baris = satu task per hari. |
| PPC Harian | otomatis | **Dashboard** + **Harian** | PPC harian, mingguan, pareto penyebab. |
| Board 1 – Monitoring | otomatis | **Dashboard** (Task Work Plan vs Final) | |
| Board 2 – Kendala | otomatis | **Make Ready** (kolom per jenis kendala) | |
| Board 3 – Pelaksanaan | otomatis | **Dashboard** (On Progress / Late / Finish) | |
| Board 4 – Harian | otomatis | **Harian** (Dijanjikan / Tercapai / Tidak / Belum lapor) | |
| Follow Up Issue | otomatis | **Issue** | Kendala Pending urut target, sisa hari, status lewat target. |
| Daftar Issue (`tblIssue`) | referensi | `issues` | Dropdown jenis issue per kategori. |
| Petunjuk LPS | dokumen | menu **Info** | Ringkasan aturan status. |

## 3. Peta kolom → field data (`seed.json` / penyimpanan lokal)

### 3.1 `meta`
| Field | Sumber Excel |
|---|---|
| `project` | MASTER!A2 |
| `tglUpdate` | LPS Input!C2 (di app bisa diganti; default hari ini) |
| `targetPPC` | PPC Harian!C4 (0,8) |
| `weeks[] {w, start, end, label}` | MASTER baris 8 (nomor), 9 (tanggal mulai), 10 (label "02 - 08"), kolom Y…BC |
| `scurve[] {w, plan, cum}` | MASTER baris 2696 (PROGRESS PLAN) & 2697 (CUMULATIVE) kolom yang sama |
| `progress[] {w, aktual}` | **Actual progress mingguan** (kumulatif %), diinput manual di Dashboard → *Input actual mingguan*; tab Sheet *Progress*. Dari Excel: MASTER baris 2700 `CUMULATIVE PROGRESS ACTUAL` (input, kolom Y–BC), baris 2701 `DEVIASI (ACTUAL - PLAN)` otomatis |

### 3.2 `pekerjaan[]` (Daftar Pekerjaan)
| Field | Kolom | Catatan |
|---|---|---|
| `kode` | D Kode Item | kunci, mis. `3.3` |
| `nama` | E Pekerjaan (Master Schedule) | dipakai sebagai relasi ke task |
| `disiplin` | B | PRELIM / CSA / ARS / MEP |
| `bill` | C Grup / Bill | |
| `mulai`, `selesai` | H, I | MIN/MAX tanggal MASTER baris F…G |
| `area` | S Area Gedung | |
| `bobot` | Σ MASTER!D baris F…G | % bobot terhadap kontrak |
| *(dihitung)* minggu mulai/selesai, status jadwal, jumlah task, disepakati, finish, % selesai | J–Q | tidak disimpan |

### 3.3 `tasks[]` (Komitmen + LPS Input, kunci `kode`)
| Field | Komitmen | LPS Input |
|---|---|---|
| `kode` | A Kode | B Kode Task |
| `minggu` | B Minggu Ke | (otomatis) |
| `disiplin` | C | (otomatis) |
| `pekerjaan` | D | (otomatis) |
| `area` | E Area Gedung | (otomatis) |
| `uraian` | F | (otomatis) |
| `lokasi` | G | (otomatis) |
| `pic` | H | (otomatis) |
| `mulai`, `selesai` | I, J | (otomatis) |
| `disepakati` | L (Ya/Tidak/kosong) | |
| `alasan` | M Alasan / Catatan Rapat | |
| `inLPS` | N "Sudah masuk" | baris ada di tblLPS |
| `k.eng {st, issue, detail, pic, target}` | | L, M, N, O, P |
| `k.mat {…}` | | Q, R, S, T, U |
| `k.mp {…}` | | V, W, X, Y, Z |
| `k.oth {…}` | | AA, AB, AC, AD, AE |
| `catatanKendala` | | AF |
| `penyebabTerlambat` | | AH |
| `prio` | N Prioritas (dropdown Urgent / Tinggi; kosong = Normal) | |
| `koord` | O Koordinasi Disiplin (mis. `MEP, ARS`) | |
| `followup` | P Catatan Follow-up | |

Sejak 03-Nov-26 tblKomit = A–Q (kolom N–P baru; *Status di LPS Input* pindah ke Q, sumber dropdown kode ke S).

### 3.4 `daily[]` (Daily Log, tabel `tblDaily`)
Kolom dibaca **berdasarkan judul** (bukan huruf kolom), jadi sisipan kolom di Excel tidak merusak ekspor.

| Field | Judul kolom Excel | Catatan |
|---|---|---|
| `id` | – | Dari Excel: `x<tanggal>_<kode>` (+`_2` bila task sama 2× sehari). Dari app: acak. |
| `tgl` | Tanggal | |
| `kode` | Kode Task | |
| `target` | Target Hari Ini | pagi |
| `bobot` | Bobot Target (% dari task) | pagi; Excel pecahan 0,2 → app 20 |
| *(input sore)* | Actual Hari Ini (% dari task) | kolom L (baru); di Excel & app yang diisi sore adalah actual |
| `capaian` | % Tercapai Harian (otomatis = Actual ÷ Bobot) | Excel 0,5 → app 50. Di app yang diisi sore adalah **actual progress hari ini (% dari task)**; capaian = actual ÷ bobot × 100 |
| `penyebab` | Penyebab Tidak Tercapai | |
| `catatan` | Catatan | |
| *(dihitung)* | Tercapai?, Progress s/d Kemarin, Kontribusi Hari Ini, Progress Kumulatif, Cek | tidak disimpan |

### 3.5 Referensi
| Field | Sumber |
|---|---|
| `areas[]` | Area Gedung!AL (daftar grup area) |
| `areaMap{}` | Area Gedung!A→B |
| `issues {eng, mat, mp, oth}` | Daftar Issue!A, B, C, D |
| `penyebab[]` | Daily Log!S |
| `disiplin[]` | CSA, MEP, ARS, PRELIM |

## 4. Rumus Excel → logika aplikasi

| Excel | Logika di app (`app.js`) |
|---|---|
| LPS Input!AI Status Kendala | `READY` bila tidak ada kendala `Pending`; selain itu `PENDING` |
| Daily Log Tercapai? | % tercapai ≥ 100 → Ya, selain itu Tidak (kosong = belum lapor) |
| Daily Log Kontribusi / Progress | kontribusi = bobot × % tercapai; s/d kemarin = Σ kontribusi task sama di tanggal sebelumnya; kumulatif = MIN(100, s/d kemarin + kontribusi) |
| *(app)* Beberapa janji / hari | satu task boleh punya beberapa janji di tanggal yang sama; janji sebelumnya di hari itu (urut id) ikut dihitung ke kumulatif & sisa task. Excel (SUMIFS tanggal "<") belum menghitung ini |
| LPS Input!AG Progress | MAXIFS progress kumulatif daily log task tsb |
| LPS Input!AJ Status Pelaksanaan | hanya untuk READY: progress ≥ 100 → `FINISH`; tglUpdate > rencana selesai → `LATE / DELAY`; selain itu `ON PROGRESS` |
| Daftar Pekerjaan!M Status Jadwal | mulai > tglUpdate → Belum mulai; selesai < tglUpdate → Selesai (jadwal); selain itu Berjalan |
| Daftar Pekerjaan!K/L Minggu | minggu terakhir yang `start ≤ tanggal` |
| Follow Up Issue | semua kendala Pending, urut target; sisa hari = target − tglUpdate; status: On track / Hari ini / ⚠ LEWAT TARGET / belum ada target |
| PPC Harian | tercapai ÷ dijanjikan per tanggal (belum lapor = belum tercapai); rata-rata % tercapai |
| PPC Harian bag. 2 PPC MINGGUAN | **PPC utama di Dashboard**: janji tercapai ÷ janji dari Daily Log per minggu (janji bertanggal setelah tanggal update belum dihitung) |
| Board 3 PPC (Finish ÷ Final) | KPI terpisah *Task Finish ÷ Final* |
| Daily Log Cek | task belum READY / isi target & bobot / bobot > sisa task / menunggu laporan sore / isi penyebab |

## 5. Arsitektur aplikasi

```
lps-pwa/
├─ PETA-EXCEL-KE-APLIKASI.md   ← dokumen ini
├─ gas/Code.gs, README.md      ← API sinkron tim (Google Apps Script + Google Sheets)
├─ tools/export_excel.py       ← Excel → app/data/seed.json (lokal saja, tidak di-commit)
├─ tools/push_master.py        ← seed.json → tab Master di Google Sheet tim
├─ tools/gas-dev-server.js     ← uji sinkron lokal tanpa Google
└─ app/                        ← folder PWA (statis, tanpa build)
   ├─ index.html, app.css, app.js
   ├─ manifest.webmanifest, sw.js (offline cache)
   ├─ icons/
   └─ data/seed.json
```

- **Tanpa framework/build** – HTML + JS murni (±60 KB), cepat di HP.
- **Offline** – service worker meng-cache aplikasi; data tersimpan di `localStorage` perangkat.
- **Sinkron data** – Excel tetap sumber jadwal: jalankan `export_excel.py`, lalu di app menu Data → "Muat ulang dari seed". Data lapangan bisa di-**Export JSON** (backup / kirim ke rekan) dan **Export CSV** (tempel balik ke Excel).
- **Prioritas** – task bisa ditandai *Urgent* / *Tinggi*, diberi disiplin lain yang perlu dikoordinasikan, dan catatan follow-up. Filter *Semua · Urgent · Urgent + tinggi · Koordinasi disiplin · Semua perlu follow up* ada di Komitmen, Make Ready, Harian, Issue (ikut ke PDF); task prioritas tampil paling atas dan di kartu Dashboard.
- **PDF per PIC** – tombol *PDF per PIC* di Komitmen, Make Ready, Harian, dan Issue: pilih minggu/tanggal dan PIC (atau semua PIC, 1 halaman per PIC), lalu *Simpan sebagai PDF* dari jendela cetak browser. Komitmen/Make Ready/Harian dikelompokkan per PIC task, Issue per PIC kendala.
- **Sinkron tim** – `tasks` dan `daily` disinkron ke Google Sheet (tab *Tasks*, *DailyLog*, *Log*) lewat Apps Script dengan satu kode akses tim. Pasang: [gas/README.md](gas/README.md).

### Peta data sinkron (Google Sheet)
| Tab Sheet | Isi | Kolom |
|---|---|---|
| Tasks | `tasks[]` (Komitmen + LPS Input) | field §3.3; kendala diratakan jadi `k.eng.st`, `k.eng.issue`, … `k.oth.target` |
| DailyLog | `daily[]` | field §3.4 |
| Master | jadwal & referensi dari Excel: `meta`, `pekerjaan`, `areas`, `areaMap`, `issues`, `penyebab`, `disiplin` | JSON dipotong per 40.000 karakter di kolom A; versi = `exportedAt` (Script Property `MASTER_VER`) |
| Log | riwayat | Waktu, User, Tabel, Kunci, Aksi (baru/ubah/hapus) |

Kolom teknis tiap record: `_u` waktu ubah (ISO, dari perangkat), `_by` nama pengubah, `_del` TRUE = dihapus, `_rev` nomor urut server.
Aturan: perubahan terakhir (`_u` terbesar) menang per record; perangkat menarik hanya baris dengan `_rev` > revisi terakhir yang ia terima.

## 6. Cara pakai

1. Update data dari Excel:
   `python lps-pwa/tools/export_excel.py "LPS - WEEKLY - BATIK PROJECT - INTERNAL.xlsx"`
2. Jalankan lokal: `python -m http.server 8080 -d lps-pwa/app` → buka `http://localhost:8080`
3. Jika file aplikasi (`app.js`, `app.css`, `index.html`) diubah, naikkan versi `CACHE` di `sw.js` (mis. `lps-batik-v2`) agar HP mengambil versi baru.
4. Hosting: repo GitHub `citrakd99/lps-monitoring`, folder `app/` terbit otomatis ke https://citrakd99.github.io/lps-monitoring/ (cabang `gh-pages` = isi folder `app/`, lihat README.md). `app/data/seed.json` sengaja tidak ikut: versi publik memuat jadwal dari tab Master setelah kode akses diisi, jadi jalankan `push_master.py` setiap jadwal Excel berubah.
