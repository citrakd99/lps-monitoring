"""Timeline Schedule (PDF dari MS Project) → bank aktivitas untuk task komitmen di aplikasi.

Pemakaian (setelah export_excel.py, karena butuh daftar pekerjaan & minggu di seed.json):
    python tools/export_timeline.py "Timeline Schedule - ....pdf" [--peta "MAPPING Timeline PDF ke Task Komitmen.xlsx"] [seed.json]

Yang dilakukan:
  1. Membaca semua baris jadwal di PDF (ID, nama, durasi, mulai, selesai) + tingkat indentasi → induk/WBS.
  2. Aktivitas terbawah (bukan judul grup) dipetakan otomatis ke area, disiplin, dan pekerjaan BOQ.
  3. Bila file --peta ada: kolom "Koreksi pekerjaan" (dan "Koreksi area"/"Koreksi disiplin" bila diisi) dipakai
     menggantikan usulan otomatis; file peta lalu ditulis ulang (koreksi & catatan tetap) untuk ditinjau.
  4. Hasil ditulis ke seed.json sebagai field `timeline` → dikirim ke Sheet tim lewat push_master.py (tidak ke repo).
Butuh: pip install pdfplumber openpyxl
"""
import json
import re
import sys
from datetime import datetime
from pathlib import Path

import pdfplumber

BULAN = {'Jan': 1, 'Feb': 2, 'Mar': 3, 'Apr': 4, 'May': 5, 'Jun': 6, 'Jul': 7, 'Aug': 8, 'Sep': 9, 'Oct': 10, 'Nov': 11, 'Dec': 12}
RX = re.compile(r"^(\d+)\s+(.*?)\s+(\d+(?:\.\d+)?)\s+days?\s+(\d{2} \w{3} '\d{2})\s+(\d{2} \w{3} '\d{2})")
GENERIK = {'CSA Work', 'CSA Works', 'CSA', 'MEP', 'MEPHVAC Work', 'MEP HVAC Work'}
ATAS = {'BATIK PROJECT ( CONSTRUCTION )', 'CSA MEP HVAC', 'Milestone', 'Project'}  # induk paling atas: tidak ditulis di WBS


def tgl(s):
    d, m, y = s.replace("'", '').split()
    return f"20{y}-{BULAN[m]:02d}-{int(d):02d}"


def baca_pdf(path):
    rows = []
    with pdfplumber.open(path) as pdf:
        for pg in pdf.pages:
            for ln in pg.extract_text_lines():
                m = RX.match(ln['text'])
                if not m:
                    continue
                ch = [c for c in ln['chars'] if c['text'].strip()]
                rows.append(dict(id=int(m[1]), nama=m[2].strip(), dur=float(m[3]), mulai=tgl(m[4]), selesai=tgl(m[5]), x=ch[len(m[1])]['x0']))
    # indentasi → tingkat (toleransi 4pt: teks tebal sedikit bergeser)
    lv, xs = {}, []
    for x in sorted({r['x'] for r in rows}):
        if xs and x - xs[-1][-1] < 4: xs[-1].append(x)
        else: xs.append([x])
        lv[x] = len(xs) - 1
    st = []
    for r in rows:
        r['lvl'] = lv[r['x']]
        while st and st[-1]['lvl'] >= r['lvl']:
            st.pop()
        r['path'] = [s['nama'] for s in st]
        r['summary'] = False
        if st:
            st[-1]['summary'] = True
        st.append(r)
    ids = {r['id'] for r in rows}
    hilang = sorted(set(range(1, max(ids) + 1)) - ids)
    return rows, hilang


def rentang(ids):
    out, a = [], None
    for i, x in enumerate(ids):
        if a is None: a = x
        if i + 1 == len(ids) or ids[i + 1] != x + 1:
            out.append(f"{a}–{x}" if x != a else str(a)); a = None
    return ', '.join(out)


P, W = [], []  # diisi main(): pekerjaan & minggu dari seed.json


def pk(prefix, bill=None):
    c = [p for p in P if p['nama'].startswith(prefix + ' ') and (not bill or bill in p['bill'])]
    return c[0]['nama'] if c else None


def week(d):
    return next((w['w'] for w in W if w['start'] <= d <= w['end']), None)


# ---------- aturan pemetaan otomatis (aktivitas PDF → area, disiplin, pekerjaan BOQ, keyakinan) ----------
AREA = [('Preparation Work', 'Preliminaries'), ('Building 2', 'Gedung 2 (Factory 2)'), ('Main Office', 'Gedung 4 (Factory 4)'), ('Fasade', 'Gedung 4 (Factory 4)'),
        ('Building 4', 'Gedung 4 (Factory 4)'), ('Water Tank', 'RWT / Pump Room / Mosque / Canteen'), ('Fire Tank', 'RWT / Pump Room / Mosque / Canteen'),
        ('Gate', 'Gate & Perimeter'), ('Automatic Sliding Door', 'Gate & Perimeter'), ('Hazardous', 'Hazardous Area'), ('Chemical Storage', 'Chemical & Compressor Room'),
        ('Canteen', 'RWT / Pump Room / Mosque / Canteen'), ('Mosque', 'RWT / Pump Room / Mosque / Canteen'), ('Pump Room', 'RWT / Pump Room / Mosque / Canteen'),
        ('Cable Utility', 'Siteplan'), ('Toilet', 'Toilet Outside'), ('Genset', 'Genset Room'), ('Trafo', 'Trafo Room / Power House'), ('External Canopy', 'Siteplan'),
        ('Drainage U-Ditch', 'Siteplan'), ('Test & Commisioning', 'Gedung 2 & 4 (Bersama)'), ('PUNCHLIST', 'Gedung 2 & 4 (Bersama)'), ('Milestone', 'Preliminaries')]


def area(path):
    for k, v in AREA:
        if any(k in x for x in path):
            return v
    return 'Other Area'


FIN = {'Gedung 2 (Factory 2)': ('7.1', '8.1', '10.1'), 'Gedung 4 (Factory 4)': ('7.2', '8.2', '10.2'), 'Chemical & Compressor Room': ('7.3', None, '10.6'),
       'Toilet Outside': ('7.4', '8.3', '10.9'), 'Genset Room': ('7.5', '8.4', '10.12'), 'Trafo Room / Power House': ('7.6', '8.5', '10.11')}


def fin(a, path, i):
    if a == 'RWT / Pump Room / Mosque / Canteen':
        m = {'Mosque': ('7.7', '8.6', '10.10'), 'Canteen': ('7.8', '8.7', '10.7'), 'Pump Room GWT': ('7.14', None, '10.13'), 'Pump Room Hydrant': ('7.15', None, '10.14')}
        for k, v in m.items():
            if any(k in x for x in path):
                return v[i]
        return None
    return FIN.get(a, (None,) * 3)[i]


MEPK = ('MEP', 'Elelctrikal', 'MECHANICAL', 'FIRE PROTECTION', 'Lighting', 'PLUMBING', 'Panel', 'Cable Contaiment', 'Power Outlet', 'VRF', 'Split Duct', 'Exhoust', 'Axial',
        'Centrifugal', 'Propeller', 'Ceilling Cassette', 'Ceiling Fan', 'Outlaying Mechanikal', 'Wall Mount', 'Exhaust Air Duct', 'Support for Hydrant', 'Support For Hydrant')


def rule(r):
    n = r['nama']; path = r['path']; a = area(path); ctx = ' / '.join(path[3:] + [n]); L = n.lower(); C = ctx.lower()
    mep = any(k in x for x in path if x != 'CSA MEP HVAC' for k in MEPK)
    if 'Milestone' in path: return a, 'PRELIM', None, 'milestone'
    if 'Preparation Work' in path: return a, 'PRELIM', pk('1.6'), 'tinggi'
    if 'Test & Commisioning' in path: return a, 'MEP', pk('1.7'), 'sedang'
    if 'PUNCHLIST & DOCUMENTATION' in path: return a, 'PRELIM', pk('1.7'), 'tinggi'
    if mep:
        if re.search(r'fire|hydrant|sprinkler', C): return a, 'MEP', pk('7.17' if 'Building' in ctx else '7.16', 'MEP'), 'sedang'
        if re.search(r'plumb|clean water|waste water|sewage|rain', C): return a, 'MEP', pk('7.22', 'MEP'), 'rendah'
        if re.search(r'compress', C): return a, 'MEP', pk('7.13', 'MEP'), 'rendah'
        if re.search(r'lighting|lamp|downlight|led|tl ', C): return a, 'MEP', pk('7.25', 'MEP'), 'sedang'
        if re.search(r'vrf|split|fan|duct|cassette|ef-|if-|iu|ou-|fac|mechanical', C): return a, 'MEP', pk('7.15', 'MEP'), 'rendah'
        if re.search(r'alarm|sound|electronic', C): return a, 'MEP', pk('7.7', 'MEP'), 'rendah'
        return a, 'MEP', pk('7.3' if 'Gedung 4' in a else '7.2', 'MEP'), 'rendah'
    if 'Floor Overlay' in ctx or 'Overlay Floor' in ctx: return a, 'CSA', pk('3.3'), 'tinggi'
    if 'PIT' in path or 'Pit' in n: return a, 'CSA', pk('3.5'), 'tinggi'
    if 'Water Tank Foundation' in ctx: return a, 'CSA', pk('3.1'), 'tinggi'
    if 'Fire Tank Foundation' in ctx: return a, 'CSA', pk('3.2'), 'tinggi'
    if 'Fan Foundation' in ctx: return a, 'CSA', pk('3.4'), 'tinggi'
    if re.search(r'jack ?ro|purlin', C): return a, 'CSA', pk('4.1'), 'sedang'
    if re.search(r'repainting st|repainting existing stee|steel repainting|repainting stell', C): return a, 'CSA', pk('4.2'), 'tinggi'
    if re.search(r'stengthening|strengtening|reinforcement structure|repair pedestal', C): return a, 'CSA', pk('4.4'), 'tinggi'
    if re.search(r'walkable|tangga acces|vertical cage', C): return a, 'CSA', pk('4.5'), 'sedang'
    if re.search(r'main gate|loading dock gate|security gate|fix fence', C):
        for k, c in (('main gate automatic', '6.1'), ('loading dock', '6.2'), ('security gate automatic sliding', '6.3'), ('swing', '6.4'), ('fix fence', '6.5'), ('manual', '6.6')):
            if k in L: return a, 'ARS', pk(c, 'AUTOMATIC'), 'tinggi'
    if 'Hazardous Waste Storage' in path: return a, 'CSA', pk('6.1', 'BILL 6'), 'sedang'
    if 'Cable Utility' in ctx: return a, 'CSA', pk('6.2', 'BILL 6'), 'sedang'
    if 'Pedestrian Canopy' in n: return a, 'CSA', pk('6.3', 'BILL 6'), 'tinggi'
    if 'Road Crossing Canopy 1' in n: return a, 'CSA', pk('6.4', 'BILL 6'), 'tinggi'
    if 'Road Crossing Canopy 2' in n: return a, 'CSA', pk('6.5', 'BILL 6'), 'tinggi'
    if 'Canteen - Mosque Refurbisment' in n: return a, 'CSA', pk('6.6', 'BILL 6'), 'tinggi'
    if re.search(r'roof|gutter|flashing|fascia|facia|rain ?water|roof drain|shadowline', C) and 'demol' not in L:
        c = fin(a, path, 2); return a, 'ARS', pk(c) if c else None, 'sedang' if c else 'rendah'
    if re.search(r'epoxy|floor tile|ceramik|tile skriting|floor finish', C):
        c = fin(a, path, 0); return a, 'ARS', pk(c, 'FLOOR') if c else None, 'sedang' if c else 'rendah'
    if re.search(r'ceiling|gypsum|plafon', C):
        c = fin(a, path, 1); return a, 'ARS', pk(c, 'CEILING') if c else None, 'sedang' if c else 'rendah'
    if re.search(r'sanitary|cubicle|toilet wall tile', C): return a, 'ARS', pk('9.1' if 'Gedung 4' in a else '9.2'), 'rendah'
    if re.search(r'door|window|jendela|louv|bouvenlight', C): return a, 'ARS', pk('5.5', 'BILL 5'), 'sedang'
    if re.search(r'scrapp', C): return a, 'ARS', pk('5.4', 'BILL 5'), 'tinggi'
    if re.search(r'paint|pengecatan', C): return a, 'ARS', pk('5.3', 'BILL 5'), 'sedang'
    if re.search(r'partition|partisi|internal wall|calsiboard', C): return a, 'ARS', pk('5.2', 'BILL 5'), 'sedang'
    if re.search(r'wall|plester|plaster|render|cladding|brick|bandwall|acp|fasade|kaca', C): return a, 'ARS', pk('5.1', 'BILL 5'), 'sedang'
    if re.search(r'demol|bobok|chipping|dismantl', C): return a, 'CSA', pk('4.1'), 'rendah'
    if re.search(r'concrete|rebar|formwork|pembesian|foundation|excavation|galian|compaction|sand|lean|beam|column|pile|soil|u-ditch|drainage|plinth|slab|anti t|limestone|footplate|tie beam|wiremesh', C):
        return a, 'CSA', None, 'rendah'
    return a, 'ARS', None, 'rendah'




KOLOM = ['ID PDF', 'WBS (induk)', 'Aktivitas (uraian task)', 'Mulai', 'Selesai', 'Durasi (hari)', 'Minggu dari', 'Minggu s/d', 'Area aplikasi', 'Disiplin',
         'Pekerjaan BOQ (usulan otomatis)', 'Keyakinan', 'Koreksi pekerjaan', 'Koreksi area', 'Koreksi disiplin', 'Catatan']


def baca_peta(path):
    """{id: {pek, area, dis, catatan}} dari kolom koreksi file peta (bila ada)."""
    if not path or not Path(path).exists():
        return {}
    from openpyxl import load_workbook
    ws = load_workbook(path, read_only=True)['Mapping aktivitas']
    it = ws.iter_rows(values_only=True)
    head = [str(h or '').strip() for h in next(it)]
    col = lambda *names: next((head.index(n) for n in names if n in head), None)
    cid, cp, ca, cd, cc = col('ID PDF'), col('Koreksi pekerjaan', 'Koreksi pekerjaan (isi bila salah)'), col('Koreksi area'), col('Koreksi disiplin'), col('Catatan')
    out = {}
    for r in it:
        if r[cid] is None:
            continue
        g = lambda c: str(r[c]).strip() if c is not None and r[c] not in (None, '') else ''
        out[int(r[cid])] = dict(pek=g(cp), area=g(ca), dis=g(cd), catatan=g(cc))
    return out


TAMBAH_KOL = ['ID induk PDF', 'ID PDF', 'Sub-grup', 'Aktivitas', 'Mulai', 'Selesai']


def baca_tambahan(path):
    """Sheet 'Tambahan aktivitas' di file peta: isi grup yang tertutup saat PDF dicetak (diketik dari MS Project)."""
    if not path or not Path(path).exists():
        return []
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True)
    if 'Tambahan aktivitas' not in wb.sheetnames:
        return []
    out = []
    for r in list(wb['Tambahan aktivitas'].iter_rows(values_only=True))[1:]:
        if not r or r[0] in (None, '') or not r[3]:
            continue
        iso = lambda v: v.date().isoformat() if hasattr(v, 'date') else str(v).strip()[:10]
        out.append(dict(induk=int(r[0]), id=int(r[1]), grup=str(r[2] or '').strip(), nama=str(r[3]).strip(), mulai=iso(r[4]), selesai=iso(r[5])))
    return out


def gabung_tambahan(rows, tambah):
    """Grup tertutup di PDF (terbaca sebagai 1 aktivitas) diganti rincian dari sheet Tambahan."""
    if not tambah:
        return rows, set()
    by = {r['id']: r for r in rows}
    induk = {t['induk'] for t in tambah if t['induk'] in by}
    out = []
    for r in rows:
        out.append(r)
        if r['id'] in induk:
            r['summary'] = True
            for t in (x for x in tambah if x['induk'] == r['id']):
                path = r['path'] + [r['nama']] + ([t['grup']] if t['grup'] else [])
                out.append(dict(id=t['id'], nama=t['nama'], dur=1, mulai=t['mulai'], selesai=t['selesai'], path=path, summary=False, tambahan=True))
    for t in tambah:
        if t['induk'] not in by:
            print(f"  ⚠ Tambahan ID {t['id']}: induk ID {t['induk']} tidak ada di PDF – dilewati")
    return out, {t['id'] for t in tambah if t['induk'] in by}


def tulis_peta(path, recs, P, weeks, hilang, tambah=()):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.worksheet.datavalidation import DataValidation
    H, HF = Font(bold=True, color='FFFFFF'), PatternFill('solid', fgColor='163A5F')
    CF = {'tinggi': 'D9F2E1', 'sedang': 'FFF4CC', 'rendah': 'FCE1D6', 'belum': 'F4CCCC', 'milestone': 'DDE6F3', 'koreksi': 'CFE2F3'}
    wb = Workbook(); ws = wb.active; ws.title = 'Mapping aktivitas'
    ws.append(KOLOM)
    for i, c in enumerate(ws[1]):
        c.font, c.fill, c.alignment = H, HF, Alignment(wrap_text=True, vertical='center')
        ws.column_dimensions[c.column_letter].width = [7, 46, 38, 11, 11, 8, 8, 8, 26, 9, 46, 10, 40, 22, 10, 24][i]
    ws.freeze_panes = 'A2'
    for o in recs:
        k = o['_kor']
        ws.append([o['id'], o['wbs'], o['nama'], o['mulai'], o['selesai'], o['_dur'], o['_w1'], o['_w2'], o['_area0'], o['_dis0'], o['_pek0'], o['_conf'],
                   k.get('pek', ''), k.get('area', ''), k.get('dis', ''), k.get('catatan', '')])
        ws.cell(ws.max_row, 12).fill = PatternFill('solid', fgColor=CF['koreksi' if any(k.get(f) for f in ('pek', 'area', 'dis')) else o['_conf']])
    ws.auto_filter.ref = ws.dimensions
    names = sorted({p['nama'] for p in P}, key=lambda n: [int(x) if x.isdigit() else 0 for x in n.split(' ')[0].split('.')])
    lst = wb.create_sheet('_daftar')
    for i, n in enumerate(names, 1): lst.cell(i, 1, n)
    areas = sorted({o['_area0'] for o in recs} | {a for p in P for a in str(p.get('area') or '').split(', ') if a})
    for i, n in enumerate(areas, 1): lst.cell(i, 2, n)
    for i, n in enumerate(['CSA', 'ARS', 'MEP', 'PRELIM'], 1): lst.cell(i, 3, n)
    lst.sheet_state = 'hidden'
    for rng, f in ((f'M2:M{ws.max_row}', f'=_daftar!$A$1:$A${len(names)}'), (f'N2:N{ws.max_row}', f'=_daftar!$B$1:$B${len(areas)}'), (f'O2:O{ws.max_row}', '=_daftar!$C$1:$C$4')):
        dv = DataValidation(type='list', formula1=f, allow_blank=True); ws.add_data_validation(dv); dv.add(rng)
    wt = wb.create_sheet('Tambahan aktivitas', 1)
    wt.append(TAMBAH_KOL)
    for c, wd in zip(wt[1], [12, 9, 34, 48, 12, 12]):
        c.font, c.fill = H, HF
        wt.column_dimensions[c.column_letter].width = wd
    for t in tambah:
        wt.append([t['induk'], t['id'], t['grup'], t['nama'], t['mulai'], t['selesai']])
    wt.freeze_panes = 'A2'
    w4 = wb.create_sheet('Catatan'); w4.column_dimensions['A'].width = 140
    for line in [
        f'{len(recs)} aktivitas terbawah dari PDF. Keyakinan: tinggi = cocok langsung; sedang = per jenis + area; rendah = tebakan per jenis; belum = pilih manual; biru = sudah dikoreksi.',
        'Isi kolom Koreksi (pilih dari daftar) bila usulan salah, simpan, lalu jalankan ulang export_timeline.py dengan --peta file ini, kemudian push_master.py.',
        f'Baris ID {rentang(hilang)} tidak ada di PDF (grup ditutup saat cetak) → ketik di sheet "Tambahan aktivitas" (ID induk = ID grup di PDF).' if hilang else 'Semua ID baris PDF lengkap (termasuk dari sheet Tambahan aktivitas).',
        'Sheet "Tambahan aktivitas": rincian grup yang tertutup di PDF. Grup induknya diganti rincian ini saat export_timeline.py dijalankan.',
        'File ini berisi data proyek internal: simpan di folder proyek, jangan di-commit ke repo publik.',
    ]:
        w4.append([line])
    wb.save(path)


def main():
    if hasattr(sys.stdout, 'reconfigure'): sys.stdout.reconfigure(encoding='utf-8', errors='replace')  # konsol Windows
    args = sys.argv[1:]
    peta = None
    if '--peta' in args:
        i = args.index('--peta'); peta = args[i + 1]; del args[i:i + 2]
    if not args:
        sys.exit(__doc__)
    seed_path = Path(args[1]) if len(args) > 1 else Path(__file__).resolve().parent.parent / 'app' / 'data' / 'seed.json'
    global P, W
    seed = json.loads(seed_path.read_text(encoding='utf-8'))
    P, W = seed['pekerjaan'], seed['meta']['weeks']
    rows, hilang = baca_pdf(args[0])
    tambah = baca_tambahan(peta)
    rows, ada_tambah = gabung_tambahan(rows, tambah)
    sisa = [i for i in hilang if i not in ada_tambah and i not in {t['induk'] for t in tambah}]
    hilang = [i for i in sisa if not any(t['induk'] < i < max(x['id'] for x in tambah if x['induk'] == t['induk']) for t in tambah)] if tambah else hilang
    kor = baca_peta(peta)
    names = {p['nama'] for p in P}
    recs = []
    for r in rows:
        if r['summary']:
            continue
        if 'Milestone' in r['path'] and r['dur'] > 0:  # PO Release, Kick Off Meeting: kegiatan administrasi, bukan milestone/task lapangan
            continue
        a, dis, pek, conf = rule(r)
        k = kor.get(r['id'], {})
        if k.get('pek') and k['pek'] not in names:
            print(f"  ⚠ ID {r['id']}: koreksi pekerjaan '{k['pek']}' tidak ada di daftar pekerjaan – diabaikan")
            k['pek'] = ''
        wbs = [x for x in r['path'][1:] if x not in ATAS]
        o = dict(id=r['id'], wbs=' › '.join(wbs), nama=r['nama'], mulai=r['mulai'], selesai=r['selesai'],
                 area=k.get('area') or a, dis=k.get('dis') or dis, pek=k.get('pek') or pek or '')
        lok = [x for x in wbs if x not in GENERIK]
        o['lok'] = ' › '.join(lok if len(lok) <= 3 else [lok[0]] + lok[-2:])  # gedung/area + 2 tingkat terbawah
        if 'Milestone' in r['path'] or r['dur'] == 0:
            o['ms'] = True
        o.update(_kor=k, _dur=int(r['dur']), _w1=week(r['mulai']), _w2=week(r['selesai']), _area0=a, _dis0=dis, _pek0=pek or '',
                 _conf=conf if pek or conf == 'milestone' else 'belum')
        if r.get('tambahan'):
            o['_dur'] = (datetime.fromisoformat(r['selesai']) - datetime.fromisoformat(r['mulai'])).days + 1
        recs.append(o)
    seed['exportedAt'] = datetime.now().isoformat(timespec='seconds')  # versi master baru → HP tim menarik timeline
    seed['timeline'] = [{k: v for k, v in o.items() if not k.startswith('_') and k != 'wbs' and v != ''} for o in recs]  # ringkas: WBS cukup di file peta
    seed_path.write_text(json.dumps(seed, ensure_ascii=False, indent=1), encoding='utf-8')
    tanpa = sum(1 for o in recs if not o['pek'] and not o.get('ms'))
    print(f"Timeline: {len(rows)} baris PDF → {len(recs)} aktivitas ({sum(1 for o in recs if o.get('ms'))} milestone, {tanpa} belum punya pekerjaan BOQ, "
          f"{sum(1 for o in recs if any(o['_kor'].get(f) for f in ('pek', 'area', 'dis')))} dikoreksi dari peta, {len(ada_tambah)} dari sheet Tambahan) → {seed_path}")
    if hilang:
        print(f"  ⚠ ID {rentang(hilang)} tidak ada di PDF (grup ditutup saat cetak).")
    if peta:
        tulis_peta(peta, recs, P, W, hilang, tambah)
        print(f"Peta untuk ditinjau: {peta}")
    print('Lanjut: python tools/push_master.py <URL_WEB_APP_/exec> <KODE_AKSES> --nama "Admin"')


if __name__ == '__main__':
    main()
