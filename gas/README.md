# Sinkron tim via Google Sheets – cara pasang (±10 menit, sekali saja)

## 1. Buat Google Sheet + script
1. Buka https://sheets.new (login akun Google kantor/proyek). Beri nama, mis. **LPS Batik – Data Tim**.
2. Menu **Extensions → Apps Script**. Hapus isi `Code.gs`, tempel seluruh isi file [`Code.gs`](Code.gs) ini, lalu **Save**.
3. Pilih fungsi **`setup`** di toolbar → **Run**. Izinkan akses (Review permissions → pilih akun → Advanced → Go to … → Allow).
4. Buka **Execution log**: muncul `Kode akses tim: XXXXXXXX`. Catat kode ini.
   (Bisa juga dilihat/diganti di ⚙ Project Settings → Script Properties → `TEAM_KEY`.)

Sheet akan berisi tab **Tasks**, **DailyLog**, **Master** (jadwal dari Excel), dan **Log** (riwayat siapa mengubah apa).
**Memperbarui script yang sudah terpasang** (mis. setelah daily log memakai bobot & % tercapai): tempel `Code.gs` terbaru → jalankan `setup()` sekali lagi (menambah tab **Master**, kolom `bobot`, `capaian` di DailyLog, `prio`, `koord`, `followup` di Tasks, tab **Progress** untuk actual mingguan, serta **Manpower** dan **Dokumen** untuk modul Resource – tab ini juga dibuat otomatis bila belum ada) → **Deploy → Manage deployments → Edit → Version: New version** (URL tetap sama). `push_master.py` menolak jalan bila server masih versi lama.

## 2. Deploy sebagai Web App
1. **Deploy → New deployment** → ikon roda → **Web app**.
2. *Execute as*: **Me** · *Who has access*: **Anyone**.
   (Siapa pun yang punya URL bisa memanggil API, tapi tanpa kode akses tim semua permintaan ditolak.)
3. **Deploy** → salin **Web app URL** (berakhiran `/exec`).

> Jika `Code.gs` diubah, pakai **Deploy → Manage deployments → Edit → Version: New version** agar URL tetap sama.

## 3. Unggah jadwal dari Excel (admin, di laptop)
Aplikasi di https://citrakd99.github.io/lps-monitoring/ tidak membawa data proyek; jadwal diambil dari tab **Master**.
```bash
python tools/export_excel.py "LPS - WEEKLY - BATIK PROJECT - INTERNAL.xlsx"
python tools/push_master.py <URL_WEB_APP_/exec> <KODE_AKSES> --nama "Admin"
```
Ulangi setiap Excel berubah. Task dari Excel hanya ditambahkan bila belum ada di Sheet. Daily log asal Excel ditambah atau diperbarui mengikuti Excel; daily log yang diisi lewat aplikasi tidak disentuh.

## 4. Hubungkan aplikasi (tiap HP/laptop)
Buka https://citrakd99.github.io/lps-monitoring/ di Chrome HP → **Isi kode akses**, lalu menu ⋮ → **Install app / Add to Home screen**.
Di aplikasi: ⚙ → **Sinkron tim** → isi **URL Web App**, **Kode akses tim**, **Nama Anda** → **Simpan**.

- Perangkat pertama yang terhubung mengunggah data awalnya ke Sheet.
- Perangkat berikutnya **mengambil data tim dulu** (data tim menang), lalu mengirim task yang hanya ada di perangkat itu.
- Setelah itu sinkron otomatis: 2,5 detik setelah ada perubahan, tiap 60 detik saat aplikasi dibuka, dan saat kembali online.
- Titik di header: hijau = tersinkron · kuning = ada perubahan belum terkirim · abu = offline · merah = gagal (ketuk untuk sinkron manual).

## Aplikasi SPV (hanya lihat)

- Alamat: `…/lps-monitoring/spv.html` – bisa dipasang ke layar utama sebagai aplikasi sendiri ("LPS SPV").
- Semua menu bisa dilihat; tidak ada tombol ubah, buat, hapus, cetak PDF, maupun ekspor.
- Kode akses SPV dibuat oleh `setup()` (lihat log *Execution log*: "Kode akses SPV (hanya lihat)"), atau lihat di
  Project Settings › Script Properties › `VIEW_KEY`. Ganti nilainya di situ untuk mencabut akses semua SPV.
- Server menolak semua kiriman tulis dari kode SPV. Aplikasi SPV hanya menerima kode SPV; aplikasi tim menolak kode SPV.
- Setelah memperbarui Code.gs: jalankan `setup()` sekali (kode tim tidak berubah), lalu Deploy › Manage deployments › Edit › New version.

## Aturan data
- Konflik (dua orang mengubah task yang sama): **perubahan paling akhir menang** untuk seluruh record task / baris daily log.
- Hapus task / daily log ikut terhapus di semua perangkat (baris di Sheet ditandai `_del = TRUE`, tidak dihapus fisik).
- Jadwal MASTER & daftar pekerjaan hanya dari Excel → `export_excel.py` → `push_master.py` (tab Master ditimpa utuh). Jangan mengedit tab Master manual.
- Jangan mengubah urutan/nama kolom baris 1 di tab Tasks/DailyLog. Mengedit isi sel langsung di Sheet tidak akan terkirim ke HP (tidak menaikkan `_rev`); ubah lewat aplikasi.
- Batas wajar: ribuan baris tetap cepat. Kuota Apps Script akun gratis ±20.000 panggilan/hari – cukup untuk tim ±30 orang.

## Uji lokal tanpa Google
```bash
node tools/gas-dev-server.js 8787 --tanpa-seed
```
`--tanpa-seed` meniru hosting publik (jadwal harus diunggah dulu dengan `push_master.py http://localhost:8787/exec DEV`). Menjalankan `Code.gs` yang sama dengan tiruan Google Sheets di memori. Aplikasi di http://localhost:8787/, URL sinkron `http://localhost:8787/exec`, kode `DEV`, data dump di `/dump`.
