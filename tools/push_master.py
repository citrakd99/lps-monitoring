"""Unggah jadwal dari Excel (seed.json) ke Google Sheet tim, agar aplikasi yang di-hosting publik
bisa memuatnya setelah kode akses diisi.

Pemakaian (setelah export_excel.py):
    python push_master.py <URL_WEB_APP_/exec> <KODE_AKSES> [seed.json] [--nama "Admin"] [--dok-link "LINK_GOOGLE_SHEET_DOKUMEN"]

Yang dikirim:
  - master  : meta (minggu, kurva-S), pekerjaan, area, daftar issue, penyebab, timeline (export_timeline.py)  → tab Master (ditimpa)
  - tasks dari Excel yang BELUM ada di Sheet (task di Sheet tidak ditimpa)
  - daily log dari Excel: baris baru ditambah, baris dari Excel yang isinya berubah diperbarui (Excel menang
    untuk baris asal Excel; baris yang dibuat di aplikasi tidak disentuh). Baris ekspor model lama (id d1, d2, …) dihapus.
"""
import json
import re
import sys
import urllib.error
import urllib.request
from datetime import datetime, timezone
from pathlib import Path

DAILY_F = ["tgl", "kode", "target", "bobot", "capaian", "penyebab", "catatan"]
MASTER_F = ["mpPlan", "meta", "pekerjaan", "areas", "areaMap", "issues", "penyebab", "disiplin", "source", "exportedAt", "timeline"]


def post(url, body):
    # text/plain = "simple request" seperti di aplikasi. Apps Script membalas 302; urllib mengikutinya sebagai GET.
    req = urllib.request.Request(url, data=json.dumps(body).encode("utf-8"),
                                 headers={"Content-Type": "text/plain;charset=utf-8"})
    try:
        with urllib.request.urlopen(req, timeout=120) as r:
            text = r.read().decode("utf-8")
    except urllib.error.URLError as e:
        sys.exit(f"Tidak bisa menghubungi {url}: {e.reason} (cek internet/DNS/VPN dan URL Web App)")
    try:
        res = json.loads(text)
    except ValueError:
        sys.exit("Balasan bukan JSON – cek URL berakhiran /exec dan deploy Web App dengan akses 'Anyone'.\n" + text[:300])
    if not res.get("ok"):
        sys.exit("Ditolak server: " + str(res.get("error")))
    return res


def main():
    args = sys.argv[1:]
    nama = "push_master"
    if "--nama" in args:
        i = args.index("--nama")
        nama = args[i + 1]
        del args[i:i + 2]
    dok_link = None
    if "--dok-link" in args:  # link Google Sheet log dokumen (dibaca aplikasi, hanya baca); kosong "" = hapus
        i = args.index("--dok-link")
        dok_link = args[i + 1].strip()
        del args[i:i + 2]
    if len(args) < 2:
        sys.exit(__doc__)
    m = re.search(r"https?://\S+", args[0])  # toleran: spasi/teks lain ikut tertempel di argumen URL
    if not m:
        sys.exit(f"Argumen pertama harus URL Web App (https://script.google.com/macros/s/.../exec), bukan: {args[0]!r}")
    url, key = m.group(0).strip("\"'"), args[1].strip()
    seed_path = Path(args[2]) if len(args) > 2 else Path(__file__).resolve().parent.parent / "app" / "data" / "seed.json"
    seed = json.loads(seed_path.read_text(encoding="utf-8"))

    team = post(url, {"key": key, "user": nama, "since": 0, "masterVer": "-"})
    now = datetime.now(timezone.utc).isoformat(timespec="milliseconds").replace("+00:00", "Z")
    body = {"key": key, "user": nama, "since": team["rev"], "master": {f: seed[f] for f in MASTER_F if f in seed}}
    # link log dokumen Drive: tetap pakai yang sudah ada di Sheet tim bila tidak diberikan
    old_link = (team.get("master") or {}).get("dokSrc") or ""
    body["master"]["dokSrc"] = old_link if dok_link is None else dok_link
    if body["master"]["dokSrc"] != old_link:  # versi master = exportedAt; ganti agar perangkat tim menarik master baru
        body["master"]["exportedAt"] = now
    known_dok = {str(r["id"]) for r in team.get("dokumen", [])}
    body["dokumen"] = [{**r, "_u": now, "_by": nama} for r in seed.get("dokumen", []) if str(r["id"]) not in known_dok]
    known = {str(r["kode"]) for r in team.get("tasks", [])}
    body["tasks"] = [{**r, "_u": now, "_by": nama} for r in seed.get("tasks", []) if str(r["kode"]) not in known]

    sheet = {str(r["id"]): r for r in team.get("daily", [])}
    if (sheet and not any("bobot" in r for r in sheet.values())) or \
            (team.get("tasks") and not any("prio" in r for r in team["tasks"])) or "progress" not in team or "dokumen" not in team:
        sys.exit("Code.gs di Google masih versi lama (belum ada kolom bobot/capaian/prioritas/tabel Progress). Tempel Code.gs terbaru, "
                 "jalankan setup(), lalu Deploy > Manage deployments > Edit > Version: New version.")
    if seed.get("timeline") and ("tlEdit" not in team or (team.get("tasks") and not any("bl" in r for r in team["tasks"]))):
        sys.exit("Code.gs di Google belum punya kolom 'bl' / tab TimelineEdit (Timeline Schedule). Tempel Code.gs terbaru, "
                 "jalankan setup(), lalu Deploy > Manage deployments > Edit > Version: New version.")
    def same(a, b):
        return all(str(a.get(f) if a.get(f) is not None else "") == str(b.get(f) if b.get(f) is not None else "")
                   or (f in ("bobot", "capaian") and a.get(f) not in (None, "") and b.get(f) not in (None, "")
                       and abs(float(a[f]) - float(b[f])) < 1e-6)
                   for f in DAILY_F)
    body["daily"] = [{**r, "_u": now, "_by": nama} for r in seed.get("daily", [])
                     if str(r["id"]) not in sheet or sheet[str(r["id"])].get("_del") == "TRUE" or not same(r, sheet[str(r["id"])])]
    old = [{"id": i, "_del": True, "_u": now, "_by": nama} for i, r in sheet.items()
           if re.fullmatch(r"d\d+", i) and r.get("_del") != "TRUE"]
    body["daily"] += old
    # actual mingguan dari Excel (baris CUMULATIVE PROGRESS ACTUAL di MASTER, bila ada): Excel menang bila nilainya beda
    prog_sheet = {str(r["w"]): r for r in team.get("progress", [])}
    body["progress"] = [{**p, "_u": now, "_by": nama} for p in seed.get("progress", [])
                        if str(p["w"]) not in prog_sheet or prog_sheet[str(p["w"])].get("_del") == "TRUE"
                        or prog_sheet[str(p["w"])].get("aktual") in (None, "")
                        or abs(float(prog_sheet[str(p["w"])]["aktual"]) - p["aktual"]) > 1e-6]
    res = post(url, body)
    print(f"Jadwal terunggah: {len(seed['meta']['weeks'])} minggu, {len(seed['pekerjaan'])} pekerjaan (versi {res['masterVer']}).")
    n_new = sum(1 for r in body["daily"] if not r.get("_del") and str(r["id"]) not in sheet)
    print(f"Task baru dari Excel: {len(body['tasks'])} (task yang sudah ada di Sheet tidak diubah).")
    print(f"Daily log: {n_new} baru, {len(body['daily']) - n_new - len(old)} diperbarui dari Excel, {len(old)} baris model lama dihapus.")
    if body["progress"]:
        print(f"Actual progress mingguan dari Excel: {len(body['progress'])} minggu diperbarui.")
    if body["master"].get("timeline"):
        print(f"Timeline Schedule: {len(body['master']['timeline'])} aktivitas (bank task komitmen).")
    if body["master"]["dokSrc"]:
        print("Link log dokumen Drive: " + ("diperbarui" if dok_link is not None else "tetap") + " (aplikasi membacanya langsung, hanya baca).")
    if res.get("skipped"):
        print(f"Catatan: {res['skipped']} record dilewati karena versi di Sheet lebih baru.")


if __name__ == "__main__":
    main()
