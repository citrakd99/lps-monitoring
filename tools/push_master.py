"""Unggah jadwal dari Excel (seed.json) ke Google Sheet tim, agar aplikasi yang di-hosting publik
bisa memuatnya setelah kode akses diisi.

Pemakaian (setelah export_excel.py):
    python push_master.py <URL_WEB_APP_/exec> <KODE_AKSES> [seed.json] [--nama "Admin"]

Yang dikirim:
  - master  : meta (minggu, kurva-S), pekerjaan, area, daftar issue, penyebab  → tab Master (ditimpa)
  - tasks / daily dari Excel yang BELUM ada di Sheet (data lapangan di Sheet tidak ditimpa)
"""
import json
import sys
import urllib.request
from datetime import datetime, timezone
from pathlib import Path

MASTER_F = ["meta", "pekerjaan", "areas", "areaMap", "issues", "penyebab", "disiplin", "source", "exportedAt"]


def post(url, body):
    # text/plain = "simple request" seperti di aplikasi. Apps Script membalas 302; urllib mengikutinya sebagai GET.
    req = urllib.request.Request(url, data=json.dumps(body).encode("utf-8"),
                                 headers={"Content-Type": "text/plain;charset=utf-8"})
    with urllib.request.urlopen(req, timeout=120) as r:
        res = json.loads(r.read().decode("utf-8"))
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
    if len(args) < 2:
        sys.exit(__doc__)
    url, key = args[0], args[1]
    seed_path = Path(args[2]) if len(args) > 2 else Path(__file__).resolve().parent.parent / "app" / "data" / "seed.json"
    seed = json.loads(seed_path.read_text(encoding="utf-8"))

    team = post(url, {"key": key, "user": nama, "since": 0, "masterVer": "-"})
    now = datetime.now(timezone.utc).isoformat(timespec="milliseconds").replace("+00:00", "Z")
    body = {"key": key, "user": nama, "since": team["rev"], "master": {f: seed[f] for f in MASTER_F if f in seed}}
    for t, k in (("tasks", "kode"), ("daily", "id")):
        known = {str(r[k]) for r in team.get(t, [])}
        body[t] = [{**r, "_u": now, "_by": nama} for r in seed.get(t, []) if str(r[k]) not in known]
    res = post(url, body)
    print(f"Jadwal terunggah: {len(seed['meta']['weeks'])} minggu, {len(seed['pekerjaan'])} pekerjaan (versi {res['masterVer']}).")
    print(f"Task baru dari Excel: {len(body['tasks'])}, daily log baru: {len(body['daily'])} (yang sudah ada di Sheet tidak diubah).")


if __name__ == "__main__":
    main()
