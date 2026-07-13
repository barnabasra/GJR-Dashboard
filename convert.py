# ============================================================
# convert.py - Auto convert GJR Excel template -> data.json
# Dipakai oleh GitHub Actions (otomatis) atau manual.
# ============================================================
import openpyxl, json, re, os, glob
from datetime import datetime

EXCEL_FILE = "GJR_Dashboard_v11_Upload_Template.xlsx"
OUTPUT = "data.json"

# Auto-detect: kalau file default tidak ada, ambil .xlsx pertama di folder
if not os.path.exists(EXCEL_FILE):
    xs = [f for f in glob.glob("*.xlsx") if not f.startswith("~$")]
    if xs:
        EXCEL_FILE = xs[0]
        print("Pakai file:", EXCEL_FILE)

wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)

def rows(s): return [[c.value for c in r] for r in wb[s].iter_rows()]
def hdr(rs, kws):
    for i, r in enumerate(rs):
        c = [str(x or '').lower() for x in r]
        if all(any(k in cc for cc in c) for k in kws): return i
    return 0
def idx(h, pat):
    for j, x in enumerate(h):
        if re.search(pat, str(x or ''), re.I): return j
    return None
def norm(s): return re.sub(r'\s+', '', str(s or '')).upper()

# Dashboard Summary
rs = rows("Dashboard Summary"); hi = hdr(rs, ["head of area"]); h = rs[hi]
c = {k: idx(h, p) for k, p in {"hoa":"head of area","area":"sales area","sal":"jml salesman",
     "tcb":"target cb","oa":"oa actual","plan":"plan call","ac":"actual call","cc":"compliance",
     "ec":"effective call","ecr":"eff call rate","jp":"jp outlet"}.items()}
HOA = []
for r in rs[hi+1:]:
    n = r[c["hoa"]]
    if not n or "total" in str(n).lower(): continue
    g = lambda k: (r[c[k]] if c[k] is not None and r[c[k]] not in (None,"") else 0)
    HOA.append([str(n).strip(), g("area"), g("sal"), g("tcb"), g("oa"),
                g("plan"), g("ac"), g("cc"), g("ec"), g("ecr"), g("jp")])

# Salesman Detail
rs = rows("Salesman Detail"); hi = hdr(rs, ["uniquecode"]); h = rs[hi]
c = {k: idx(h, p) for k, p in {"code":"uniquecode","name":"salesman name","typ":"salesman type",
     "hoa":"head of area","area":"sales area","tcb":"target cb","oa":"oa actual","ach":"ach",
     "plan":"plan call","ac":"actual call","ec":"effective call"}.items()}
SM = []
for r in rs[hi+1:]:
    cd = r[c["code"]]
    if not cd: continue
    g = lambda k: (r[c[k]] if c[k] is not None and r[c[k]] not in (None,"") else 0)
    SM.append([cd, g("name"), g("typ"), str(g("hoa")).strip(), g("area"),
               g("tcb"), g("oa"), g("ach"), g("plan"), g("ac"), g("ec")])

# JP
rs = rows("JP"); hi = hdr(rs, ["outlet code"]); h = rs[hi]
c = {k: idx(h, p) for k, p in {"reg":"region","area":"sales area","dist":"shipto|distributor",
     "code":"outlet code","out":"outlet-name|outlet name","sls":"slsman","ee":"^ee$|eagle"}.items()}
jpc = set(); BE = []
for r in rs[hi+1:]:
    cd = r[c["code"]]
    if not cd: continue
    jpc.add(norm(cd))
    ee = str(r[c["ee"]]).strip().lower() if c["ee"] is not None else ""
    g = lambda k: (r[c[k]] if c[k] is not None and r[c[k]] not in (None,"") else "")
    if "eagle" not in ee:
        BE.append([g("reg") or "Grtr Jakarta", g("area"), g("dist"), cd, g("out"), g("sls"), "", 0])

# Registrasi per Area
rs = rows("Registrasi per Area"); hi = hdr(rs, ["sales area", "outlet active ee"]); h = rs[hi]
c = {k: idx(h, p) for k, p in {"area":"sales area","yd":"outlet ytd",
     "aee":"outlet active ee","at":"outlet active total"}.items()}
EE = []; AT = {}
for r in rs[hi+1:]:
    a = r[c["area"]]
    if not a or "total" in str(a).lower() or "hipos" in str(a).lower(): break
    g = lambda k: (r[c[k]] if c[k] is not None and r[c[k]] not in (None,"") else 0)
    EE.append([a, g("aee"), g("yd")]); AT[a] = g("at")

# area -> first HoA
a2h = {}
for x in HOA: a2h.setdefault(x[1], x[0])

# EE Hipos List
rs = rows("EE Hipos List"); hi = hdr(rs, ["source", "outlet code"]); h = rs[hi]
c = {k: idx(h, p) for k, p in {"src":"source","area":"sales area","dist":"shipto|distributor",
     "code":"outlet code","out":"outlet-name|outlet name"}.items()}
BJ = []
for r in rs[hi+1:]:
    cd = r[c["code"]]
    if not cd: continue
    if norm(cd) in jpc: continue
    g = lambda k: (r[c[k]] if c[k] is not None and r[c[k]] not in (None,"") else "")
    BJ.append([g("src"), g("area"), g("dist"), cd, g("out"), a2h.get(g("area"), "-")])

DATA = {"label": "GJR Data " + datetime.now().strftime("%d %b %Y"),
        "hoa": HOA, "salesman": SM, "ee": EE, "at": AT, "belumEE": BE, "belumJP": BJ}

with open(OUTPUT, "w", encoding="utf-8") as f:
    json.dump(DATA, f, ensure_ascii=False, indent=2)

print(f"OK data.json | HoA {len(HOA)} | Salesman {len(SM)} | EE {len(EE)} | BelumEE {len(BE)} | BelumJP {len(BJ)}")
