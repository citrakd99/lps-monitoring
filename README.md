# LPS Monitoring – PWA Last Planner System

Aplikasi ringan (PWA, bisa di-install di HP, jalan offline) untuk monitoring Komitmen, Make Ready, Daily Log, dan PPC.

- Aplikasi: https://citrakd99.github.io/lps-monitoring/
- Data proyek **tidak** ada di repo ini. Jadwal dan data lapangan disimpan di Google Sheet tim dan baru dimuat setelah URL Web App + kode akses tim diisi di ⚙.

## Untuk admin
1. Pasang API sinkron sekali: [gas/README.md](gas/README.md).
2. Tiap jadwal Excel berubah:
   ```bash
   python tools/export_excel.py "LPS - WEEKLY - BATIK PROJECT - INTERNAL.xlsx"
   python tools/push_master.py <URL_WEB_APP_/exec> <KODE_AKSES> --nama "Admin"
   ```
   HP mengambil jadwal baru otomatis pada sinkron berikutnya.
   Log dokumen dibaca aplikasi langsung dari Google Sheet di Drive (hanya baca). Atur link-nya sekali untuk seluruh tim
   (akses link: "Siapa saja yang memiliki link" – Pelihat); link disimpan di Sheet tim, bukan di repo publik:
   ```bash
   python tools/push_master.py <URL_WEB_APP_/exec> <KODE_AKSES> --nama "Admin" --dok-link "https://docs.google.com/spreadsheets/d/…/edit"
   ```
3. Perubahan kode aplikasi: naikkan `CACHE` di `app/sw.js`, commit, lalu terbitkan ulang folder `app/` ke cabang `gh-pages`:
   ```bash
   git push origin main
   git push origin "$(git subtree split --prefix app main)":refs/heads/gh-pages --force
   ```

Peta Excel → aplikasi: [PETA-EXCEL-KE-APLIKASI.md](PETA-EXCEL-KE-APLIKASI.md).
