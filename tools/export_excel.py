"""Ekspor workbook LPS (Excel) ke app/data/seed.json untuk PWA.

Pemakaian:
    python export_excel.py "LPS - WEEKLY - BATIK PROJECT - INTERNAL.xlsx" [output.json]

Peta kolom lengkap ada di ../PETA-EXCEL-KE-APLIKASI.md
"""
import json
import sys
import warnings
from datetime import date, datetime
from pathlib import Path

import openpyxl
from openpyxl.utils import column_index_from_string as col

warnings.filterwarnings("ignore")

KAT = ["eng", "mat", "mp", "oth"]
# Kolom awal (Status) tiap kategori kendala di LPS Input: Status, Issue, Detail, PIC, Target
KAT_COL = {"eng": "L", "mat": "Q", "mp": "V", "oth": "AA"}


def iso(v):
    if isinstance(v, (datetime, date)):
        return v.strftime("%Y-%m-%d")
    return None


def txt(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def num(v, default=0):
    try:
        return float(v)
    except (TypeError, ValueError):
        return default


def rows(ws, first, key_col):
    """Baris data mulai `first` selama kolom kunci terisi (melewati baris kosong di tengah)."""
    k = col(key_col)
    for r in range(first, ws.max_row + 1):
        if txt(ws.cell(r, k).value):
            yield r


def export(xlsx, out):
    wb = openpyxl.load_workbook(xlsx, data_only=True)
    m = wb["MASTER"]
    g = lambda ws, ref: ws[ref].value  # noqa: E731

    # --- Minggu & kurva-S (MASTER baris 8,9,10 / 2696,2697) ---
    scurve_row = next(r for r in range(m.max_row, 12, -1) if txt(m.cell(r, 1).value) == "PROGRESS PLAN")
    weeks, scurve = [], []
    for c in range(col("Y"), col("BC") + 1):
        start = m.cell(9, c).value
        if not isinstance(start, datetime):
            continue
        w = int(num(m.cell(8, c).value, 0))
        nxt = m.cell(9, c + 1).value
        end = (nxt.toordinal() - 1) if isinstance(nxt, datetime) else start.toordinal() + 6
        weeks.append({"w": w, "start": iso(start), "end": date.fromordinal(end).isoformat(),
                      "label": txt(m.cell(10, c).value)})
        scurve.append({"w": w, "plan": round(num(m.cell(scurve_row, c).value), 4),
                       "cum": round(num(m.cell(scurve_row + 1, c).value), 4)})

    # --- Area Gedung ---
    ag = wb["Area Gedung"]
    area_map = {txt(ag.cell(r, 1).value): txt(ag.cell(r, 2).value) for r in rows(ag, 5, "A")
                if txt(ag.cell(r, 2).value)}
    areas = [txt(ag.cell(r, col("AL")).value) for r in rows(ag, 5, "AL")]

    # --- Daftar Pekerjaan (+ bobot dari MASTER kolom D) ---
    dp = wb["Daftar Pekerjaan"]
    pekerjaan = []
    for r in rows(dp, 5, "D"):
        r0, r1 = int(num(g(dp, f"F{r}"))), int(num(g(dp, f"G{r}")))
        bobot = sum(num(m.cell(i, 4).value) for i in range(r0 + 1, r1 + 1)) if r0 and r1 >= r0 else 0
        pekerjaan.append({
            "kode": txt(g(dp, f"D{r}")), "nama": txt(g(dp, f"E{r}")),
            "disiplin": txt(g(dp, f"B{r}")), "bill": txt(g(dp, f"C{r}")),
            "mulai": iso(g(dp, f"H{r}")), "selesai": iso(g(dp, f"I{r}")),
            "area": txt(g(dp, f"S{r}")), "ket": txt(g(dp, f"R{r}")),
            "bobot": round(bobot, 4),
        })

    # --- Komitmen + LPS Input → tasks ---
    km, lp = wb["Komitmen"], wb["LPS Input"]
    tasks = {}
    for r in rows(km, 6, "A"):
        kode = txt(g(km, f"A{r}"))
        tasks[kode] = {
            "kode": kode, "minggu": int(num(g(km, f"B{r}"), 0)), "disiplin": txt(g(km, f"C{r}")),
            "pekerjaan": txt(g(km, f"D{r}")), "area": txt(g(km, f"E{r}")), "uraian": txt(g(km, f"F{r}")),
            "lokasi": txt(g(km, f"G{r}")), "pic": txt(g(km, f"H{r}")),
            "mulai": iso(g(km, f"I{r}")), "selesai": iso(g(km, f"J{r}")),
            "disepakati": txt(g(km, f"L{r}")), "alasan": txt(g(km, f"M{r}")),
            "inLPS": False, "k": {k: {"st": "OK", "issue": "", "detail": "", "pic": "", "target": None} for k in KAT},
            "catatanKendala": "", "penyebabTerlambat": "",
        }
    for r in rows(lp, 5, "B"):
        t = tasks.get(txt(g(lp, f"B{r}")))
        if not t:
            continue
        t["inLPS"] = True
        for k, c0 in KAT_COL.items():
            c = col(c0)
            t["k"][k] = {"st": txt(lp.cell(r, c).value) or "OK", "issue": txt(lp.cell(r, c + 1).value),
                         "detail": txt(lp.cell(r, c + 2).value), "pic": txt(lp.cell(r, c + 3).value),
                         "target": iso(lp.cell(r, c + 4).value)}
        t["catatanKendala"] = txt(g(lp, f"AF{r}"))
        t["penyebabTerlambat"] = txt(g(lp, f"AH{r}"))

    # --- Daily Log (tblDaily) – kolom dicari dari judulnya, jadi tahan bila kolom disisipkan ---
    dl = wb["Daily Log"]
    hr = int("".join(ch for ch in dl.tables["tblDaily"].ref.split(":")[0] if ch.isdigit()))  # baris judul tabel
    hdr = {}
    for c in dl[hr]:  # kemunculan pertama = kolom tabel (judul yang sama di kanan = sumber dropdown)
        if txt(c.value):
            hdr.setdefault(txt(c.value), c.column)

    def dv(r, name):
        if name not in hdr:
            sys.exit(f"Kolom '{name}' tidak ada di Daily Log (judul baris {hr})")
        return dl.cell(r, hdr[name]).value

    def pct100(v):  # Excel menyimpan persen sebagai pecahan (1 = 100%)
        return None if v in (None, "") else round(num(v) * 100, 4)

    daily, seen = [], {}
    for r in rows(dl, hr + 1, "B"):
        tgl_, kode_ = iso(dv(r, "Tanggal")), txt(dv(r, "Kode Task"))
        n = seen[(tgl_, kode_)] = seen.get((tgl_, kode_), 0) + 1
        daily.append({
            # id stabil dari tanggal + kode (bukan nomor baris) agar baris sisipan tidak menggeser id
            "id": f"x{tgl_}_{kode_}" + (f"_{n}" if n > 1 else ""), "tgl": tgl_, "kode": kode_,
            "target": txt(dv(r, "Target Hari Ini")), "bobot": pct100(dv(r, "Bobot Target (% dari task)")),
            "capaian": pct100(dv(r, "% Tercapai Harian")),
            "penyebab": txt(dv(r, "Penyebab Tidak Tercapai")), "catatan": txt(dv(r, "Catatan")),
        })

    di = wb["Daftar Issue"]
    issues = {k: [txt(di.cell(r, i + 1).value) for r in range(5, di.max_row + 1) if txt(di.cell(r, i + 1).value)]
              for i, k in enumerate(KAT)}
    # daftar penyebab = kolom berjudul "Penyebab Tidak Tercapai" di luar tabel (sumber dropdown)
    pc_ = next((c.column for c in dl[hr] if txt(c.value) == "Penyebab Tidak Tercapai" and c.column != hdr["Penyebab Tidak Tercapai"]), None)
    penyebab = [txt(dl.cell(r, pc_).value) for r in range(hr + 1, hr + 60) if pc_ and txt(dl.cell(r, pc_).value)]

    seed = {
        "version": 1, "exportedAt": datetime.now().isoformat(timespec="seconds"), "source": Path(xlsx).name,
        "meta": {"project": txt(g(m, "A2")), "tglUpdate": iso(g(lp, "C2")),
                 "targetPPC": num(g(wb["PPC Harian"], "C4"), 0.8), "weeks": weeks, "scurve": scurve},
        "disiplin": ["CSA", "MEP", "ARS", "PRELIM"], "areas": areas, "areaMap": area_map,
        "issues": issues, "penyebab": penyebab,
        "pekerjaan": pekerjaan, "tasks": list(tasks.values()), "daily": daily,
    }
    Path(out).parent.mkdir(parents=True, exist_ok=True)
    Path(out).write_text(json.dumps(seed, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"OK → {out}: {len(weeks)} minggu, {len(pekerjaan)} pekerjaan, {len(tasks)} task, {len(daily)} daily log")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    default_out = Path(__file__).resolve().parent.parent / "app" / "data" / "seed.json"
    export(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else default_out)
