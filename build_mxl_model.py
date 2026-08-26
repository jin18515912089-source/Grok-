#!/usr/bin/env python3
"""Build MaxLinear (NASDAQ: MXL) integrated 3-statement financial model (single sheet)."""
from __future__ import annotations

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule

from mxl_hist import QPL, QBS

# ---------------------------------------------------------------------------
# Periods
# ---------------------------------------------------------------------------
ANN_YEARS = list(range(2016, 2029))
QUARTERS = [(y, q) for y in range(2016, 2029) for q in (1, 2, 3, 4)]


def is_actual_q(y, q):
    return (y, q) <= (2026, 2)


# ---------------------------------------------------------------------------
# Annual reported financials ($000) — 10-K / XBRL duration ~365d
# ---------------------------------------------------------------------------
ANN = {
    2016: dict(rev=387832, cogs=157842, rd=97745, sga=64454, oi=63059, ie=104, ii=572, oth=163,
               ebt=63690, tax=2398, ni=61292, rs=3432, imp=1300, da=26703, sbc=21765,
               cfo=117317, cfi=-101313, cff=-670, fx=0, capex=8512, sh_b=63781, sh_d=67653),
    2017: dict(rev=420318, cogs=212355, rd=112279, sga=105831, oi=-21671, ie=10378, ii=274, oth=-2223,
               ebt=-33998, tax=-24811, ni=-9187, rs=9524, imp=2000, da=66738, sbc=32668,
               cfo=75064, cfi=-432151, cff=347021, fx=0, capex=7468, sh_b=66252, sh_d=66252),
    2018: dict(rev=384997, cogs=176223, rd=120046, sga=101789, oi=-19097, ie=14255, ii=78, oth=422,
               ebt=-32852, tax=-6653, ni=-26199, rs=3838, imp=2198, da=79027, sbc=31721,
               cfo=102689, cfi=-7825, cff=-93784, fx=0, capex=7825, sh_b=68490, sh_d=68490),
    2019: dict(rev=317180, cogs=149495, rd=98344, sga=88762, oi=-22057, ie=11133, ii=775, oth=-69,
               ebt=-32484, tax=-12586, ni=-19898, rs=2636, imp=0, da=66401, sbc=32060,
               cfo=78348, cfi=-6973, cff=-53383, fx=934, capex=6887, sh_b=71005, sh_d=71005),
    2020: dict(rev=478596, cogs=265798, rd=179993, sga=130025, oi=-101139, ie=12952, ii=409, oth=-1170,
               ebt=-114852, tax=-16259, ni=-98593, rs=3833, imp=86, da=76513, sbc=47597,
               cfo=73593, cfi=-175286, cff=159649, fx=-1039, capex=12487, sh_b=73133, sh_d=73133),
    2021: dict(rev=892398, cogs=396566, rd=278440, sga=149943, oi=65245, ie=12996, ii=78, oth=764,
               ebt=47870, tax=5901, ni=41969, rs=2204, imp=0, da=91792, sbc=59358,
               cfo=168233, cfi=-91757, cff=-91903, fx=-2869, capex=39176, sh_b=76037, sh_d=79679),
    2022: dict(rev=1120252, cogs=470483, rd=296442, sga=168008, oi=180243, ie=9768, ii=245, oth=3478,
               ebt=174198, tax=49158, ni=125040, rs=2265, imp=2811, da=80731, sbc=81704,
               cfo=388726, cfi=-91762, cff=-240401, fx=56, capex=41253, sh_b=78039, sh_d=80852),
    2023: dict(rev=693263, cogs=307600, rd=269504, sga=132156, oi=-38221, ie=10702, ii=6053, oth=-20940,
               ebt=-63810, tax=9337, ni=-73147, rs=19786, imp=2438, da=71516, sbc=55176,
               cfo=43372, cfi=-15935, cff=-26356, fx=-1082, capex=13454, sh_b=80719, sh_d=80719),
    2024: dict(rev=360528, cogs=165746, rd=225189, sga=138329, oi=-223352, ie=10874, ii=6386, oth=-10877,
               ebt=-238717, tax=6481, ni=-245198, rs=53379, imp=1237, da=54140, sbc=66021,
               cfo=-45295, cfi=-23446, cff=1286, fx=-1298, capex=17680, sh_b=83600, sh_d=83600),
    2025: dict(rev=467641, cogs=201827, rd=208599, sga=159580, oi=-126890, ie=10056, ii=3385, oth=-7333,
               ebt=-140894, tax=-4213, ni=-136681, rs=24525, imp=0, da=43992, sbc=77128,
               cfo=19619, cfi=-19795, cff=-18659, fx=644, capex=12598, sh_b=86588, sh_d=86588),
}

# 10-K four-bucket segments ($000). 2016–20 reconstructed (pre-disclosure mapping).
SEG_ANN = {
    2016: dict(bb=302000, conn=35000, inf=28000, ind=22832),
    2017: dict(bb=310000, conn=40000, inf=38000, ind=32318),
    2018: dict(bb=280000, conn=42000, inf=35000, ind=27997),
    2019: dict(bb=220000, conn=38000, inf=32000, ind=27180),
    2020: dict(bb=280000, conn=55000, inf=70000, ind=73596),
    2021: dict(bb=492482, conn=149285, inf=119421, ind=131210),
    2022: dict(bb=493232, conn=303925, inf=136274, ind=186821),
    2023: dict(bb=203519, conn=138228, inf=177083, ind=174433),
    2024: dict(bb=116819, conn=55769, inf=113907, ind=74033),
    2025: dict(bb=204423, conn=77990, inf=148164, ind=37064),
}

# Optical analog (DSP+TIA+Driver) inside Infrastructure — not disclosed; reconstructed.
OPT_ANN = {y: 0 for y in range(2016, 2024)}
OPT_ANN[2024] = 9600
OPT_ANN[2025] = 56000

Q_SEG = {
    (2025, 1): dict(bb=40900, conn=20200, ind=8300),
    (2025, 2): dict(bb=47600, conn=20700, ind=5800),
    (2025, 3): dict(bb=58200, conn=19000, ind=8900),
    (2026, 1): dict(bb=43600, conn=18600, ind=12200),
    (2026, 2): dict(bb=44900, conn=24000, ind=15000),
}
Q_OPT = {
    (2025, 1): 6000, (2025, 2): 12000, (2025, 3): 18000, (2025, 4): 20000,
    (2026, 1): 35000, (2026, 2): 52000,
}

# Optical mix DSP / TIA / Driver of optical revenue
def opt_mix(y, q=None):
    if y <= 2025:
        return 0.82, 0.11, 0.07
    if y == 2026:
        return 0.80, 0.12, 0.08
    if y == 2027:
        return 0.75, 0.15, 0.10
    return 0.68, 0.18, 0.14


# Fill Q4 P&L as FY residual so annual = 4Q
Q4_KEYS = ["rev", "cogs", "rd", "sga", "oi", "ie", "ii", "oth", "ebt", "tax", "ni", "rs", "imp"]
for y in range(2016, 2026):
    a = ANN[y]
    q4 = dict(QPL.get((y, 4), {}))
    for key in Q4_KEYS:
        s123 = sum(QPL[(y, q)].get(key, 0) for q in (1, 2, 3))
        q4[key] = a[key] - s123
    q4["gp"] = q4["rev"] - q4["cogs"]
    q4["shb"] = a["sh_b"]
    q4["shd"] = a["sh_d"]
    QPL[(y, 4)] = q4
    tot = sum(QPL[(y, q)]["rev"] for q in (1, 2, 3, 4))
    if abs(tot - a["rev"]) > 2:
        raise SystemExit(f"Q-sum rev mismatch {y}: {tot} vs {a['rev']}")

# Defaults on actual quarters
for (y, q), p in list(QPL.items()):
    p.setdefault("ie", 0)
    p.setdefault("ii", 0)
    p.setdefault("oth", 0)
    p.setdefault("rs", 0)
    p.setdefault("imp", 0)
    p.setdefault("cogs", p.get("gp") and (p["rev"] - p["gp"]) or 0)
    if p.get("ie", 0) == 0 and p.get("ebt") is not None and p.get("oi") is not None:
        implied = p["oi"] + p.get("ii", 0) + p.get("oth", 0) - p["ebt"]
        if abs(implied) > 50:
            p["ie"] = implied
    p.setdefault("shb", ANN[y]["sh_b"] if y <= 2025 else 90000)
    p.setdefault("shd", ANN[y]["sh_d"] if y <= 2025 else 97000)

# Press-release Q2'26 interest expense
QPL[(2026, 2)]["ie"] = 2269
QPL[(2026, 1)]["ie"] = QPL[(2026, 1)]["oi"] + QPL[(2026, 1)]["ii"] + QPL[(2026, 1)]["oth"] - QPL[(2026, 1)]["ebt"]

# Q2'26 revolver draw
QBS[(2026, 2)]["std"] = 20000
for k, b in QBS.items():
    b.setdefault("std", 0)
    b.setdefault("gw", 0)
    b.setdefault("ia", 0)
    b.setdefault("ltd", 0)

# Historical optical ASP / unit COGS (blue)
ANN_ASP = {y: dict(dsp=0, tia=0, drv=0) for y in range(2016, 2024)}
ANN_ASP[2024] = dict(dsp=48.0, tia=10.0, drv=8.0)
ANN_ASP[2025] = dict(dsp=49.0, tia=10.2, drv=8.2)
ANN_UCOGS = {y: dict(dsp=0, tia=0, drv=0) for y in range(2016, 2024)}
ANN_UCOGS[2024] = dict(dsp=22.0, tia=4.2, drv=3.5)
ANN_UCOGS[2025] = dict(dsp=21.8, tia=4.1, drv=3.4)


def allocate_product(y, q):
    rev = QPL[(y, q)]["rev"]
    fy = ANN[y]["rev"] if y <= 2025 else None
    seg = SEG_ANN.get(y)
    known = Q_SEG.get((y, q))
    if known:
        bb, conn, ind = known["bb"], known["conn"], known["ind"]
        inf = rev - bb - conn - ind
    elif seg and fy:
        bb = round(seg["bb"] * rev / fy)
        conn = round(seg["conn"] * rev / fy)
        ind = round(seg["ind"] * rev / fy)
        inf = rev - bb - conn - ind
    else:
        bb = conn = ind = inf = 0
    if (y, q) in Q_OPT:
        opt = Q_OPT[(y, q)]
    elif y <= 2025 and fy:
        opt = round(OPT_ANN[y] * rev / fy)
    else:
        opt = 0
    opt = min(opt, max(inf, 0))
    md, mt, mv = opt_mix(y, q)
    dsp = round(opt * md)
    tia = round(opt * mt)
    drv = opt - dsp - tia
    oth_inf = inf - opt
    # rounding into broadband
    slack = rev - (dsp + tia + drv + oth_inf + bb + conn + ind)
    bb += slack
    return dict(dsp=dsp, tia=tia, drv=drv, oth_inf=oth_inf, bb=bb, conn=conn, ind=ind)


def q_asp(y, q, kind):
    base = ANN_ASP.get(y, {}).get(kind, 0) or (50 if kind == "dsp" else 10 if kind == "tia" else 8)
    if y < 2024:
        return base
    bump = {1: -1.0, 2: -0.3, 3: 0.3, 4: 1.0}[q]
    if kind != "dsp":
        bump *= 0.15
    if (y, q) == (2026, 1):
        return {"dsp": 49.5, "tia": 10.4, "drv": 8.4}[kind]
    if (y, q) == (2026, 2):
        return {"dsp": 50.0, "tia": 10.5, "drv": 8.5}[kind]
    return round(base + bump, 2)


def q_ucogs(y, q, kind):
    base = ANN_UCOGS.get(y, {}).get(kind, 0) or (22 if kind == "dsp" else 4.2 if kind == "tia" else 3.5)
    if y < 2024:
        return base
    if (y, q) >= (2026, 1):
        return {"dsp": 21.2, "tia": 4.0, "drv": 3.4}[kind]
    return base


# Forecast Q3'26–Q4'28 (blue inputs)
FCST = {
    (2026, 3): dict(
        dsp_vol=1160, dsp_asp=50.0, dsp_ucogs=21.0,
        tia_vol=820, tia_asp=10.5, tia_ucogs=4.0,
        drv_vol=680, drv_asp=8.5, drv_ucogs=3.4,
        oth_inf=42110, bb=52000, conn=30000, ind=15500,
        other_gm=0.575, sga=47000, rd=55500, rs=500,
        tax_rate=0.12, tax_min=1500, yld=0.008, drate=0.018, oth=0,
        dso=32, dio=118, dpo=52, capex=3500, sbc=21000,
        da_ppe=3500, da_ia=7000, sh_b=91000, sh_d=99000,
        stdebt=18000, ltdebt=124080, fx=0,
    ),
    (2026, 4): dict(
        dsp_vol=1280, dsp_asp=51.0, dsp_ucogs=20.8,
        tia_vol=900, tia_asp=10.7, tia_ucogs=4.0,
        drv_vol=760, drv_asp=8.6, drv_ucogs=3.4,
        oth_inf=44554, bb=54000, conn=32000, ind=16000,
        other_gm=0.576, sga=48000, rd=56000, rs=0,
        tax_rate=0.12, tax_min=1500, yld=0.008, drate=0.018, oth=0,
        dso=32, dio=112, dpo=52, capex=3800, sbc=21500,
        da_ppe=3400, da_ia=6500, sh_b=91500, sh_d=99500,
        stdebt=12000, ltdebt=124240, fx=0,
    ),
    (2027, 1): dict(
        dsp_vol=1380, dsp_asp=56.0, dsp_ucogs=22.0,
        tia_vol=1100, tia_asp=11.0, tia_ucogs=4.1,
        drv_vol=920, drv_asp=9.0, drv_ucogs=3.5,
        oth_inf=43000, bb=56000, conn=34000, ind=16500,
        other_gm=0.578, sga=49000, rd=57500, rs=0,
        tax_rate=0.12, tax_min=1500, yld=0.009, drate=0.018, oth=0,
        dso=34, dio=108, dpo=52, capex=4000, sbc=22000,
        da_ppe=3300, da_ia=5000, sh_b=92000, sh_d=100000,
        stdebt=8000, ltdebt=124400, fx=0,
    ),
    (2027, 2): dict(
        dsp_vol=1500, dsp_asp=60.0, dsp_ucogs=23.0,
        tia_vol=1280, tia_asp=11.2, tia_ucogs=4.1,
        drv_vol=1080, drv_asp=9.2, drv_ucogs=3.5,
        oth_inf=44000, bb=57500, conn=35500, ind=17000,
        other_gm=0.580, sga=50000, rd=58500, rs=0,
        tax_rate=0.12, tax_min=1500, yld=0.009, drate=0.018, oth=0,
        dso=34, dio=105, dpo=53, capex=4200, sbc=22500,
        da_ppe=3200, da_ia=3500, sh_b=92500, sh_d=100500,
        stdebt=4000, ltdebt=124560, fx=0,
    ),
    (2027, 3): dict(
        dsp_vol=1620, dsp_asp=64.0, dsp_ucogs=24.0,
        tia_vol=1480, tia_asp=11.4, tia_ucogs=4.2,
        drv_vol=1250, drv_asp=9.4, drv_ucogs=3.6,
        oth_inf=45000, bb=59000, conn=37000, ind=17500,
        other_gm=0.582, sga=51000, rd=59500, rs=0,
        tax_rate=0.12, tax_min=1500, yld=0.009, drate=0.018, oth=0,
        dso=33, dio=100, dpo=53, capex=4500, sbc=23000,
        da_ppe=3100, da_ia=2500, sh_b=93000, sh_d=101000,
        stdebt=0, ltdebt=124720, fx=0,
    ),
    (2027, 4): dict(
        dsp_vol=1750, dsp_asp=68.0, dsp_ucogs=25.0,
        tia_vol=1680, tia_asp=11.6, tia_ucogs=4.2,
        drv_vol=1420, drv_asp=9.6, drv_ucogs=3.6,
        oth_inf=46000, bb=60500, conn=38500, ind=18000,
        other_gm=0.584, sga=52000, rd=60500, rs=0,
        tax_rate=0.12, tax_min=1500, yld=0.010, drate=0.018, oth=0,
        dso=32, dio=96, dpo=53, capex=4500, sbc=23500,
        da_ppe=3000, da_ia=2000, sh_b=93500, sh_d=101500,
        stdebt=0, ltdebt=120000, fx=0,
    ),
    (2028, 1): dict(
        dsp_vol=1850, dsp_asp=74.0, dsp_ucogs=26.5,
        tia_vol=1900, tia_asp=11.8, tia_ucogs=4.3,
        drv_vol=1600, drv_asp=9.8, drv_ucogs=3.7,
        oth_inf=47000, bb=62000, conn=40000, ind=18500,
        other_gm=0.585, sga=53000, rd=61500, rs=0,
        tax_rate=0.13, tax_min=1500, yld=0.010, drate=0.017, oth=0,
        dso=32, dio=92, dpo=54, capex=4800, sbc=24000,
        da_ppe=3000, da_ia=1800, sh_b=94000, sh_d=102000,
        stdebt=0, ltdebt=115000, fx=0,
    ),
    (2028, 2): dict(
        dsp_vol=1960, dsp_asp=78.0, dsp_ucogs=27.5,
        tia_vol=2100, tia_asp=12.0, tia_ucogs=4.3,
        drv_vol=1780, drv_asp=10.0, drv_ucogs=3.7,
        oth_inf=48000, bb=63500, conn=41500, ind=19000,
        other_gm=0.586, sga=53500, rd=62500, rs=0,
        tax_rate=0.13, tax_min=1500, yld=0.010, drate=0.017, oth=0,
        dso=31, dio=90, dpo=54, capex=5000, sbc=24500,
        da_ppe=3000, da_ia=1600, sh_b=94500, sh_d=102500,
        stdebt=0, ltdebt=110000, fx=0,
    ),
    (2028, 3): dict(
        dsp_vol=2080, dsp_asp=82.0, dsp_ucogs=28.5,
        tia_vol=2300, tia_asp=12.2, tia_ucogs=4.4,
        drv_vol=1960, drv_asp=10.2, drv_ucogs=3.8,
        oth_inf=49000, bb=65000, conn=43000, ind=19500,
        other_gm=0.587, sga=54000, rd=63500, rs=0,
        tax_rate=0.13, tax_min=1500, yld=0.010, drate=0.017, oth=0,
        dso=30, dio=88, dpo=54, capex=5000, sbc=25000,
        da_ppe=2900, da_ia=1500, sh_b=95000, sh_d=103000,
        stdebt=0, ltdebt=105000, fx=0,
    ),
    (2028, 4): dict(
        dsp_vol=2200, dsp_asp=86.0, dsp_ucogs=29.5,
        tia_vol=2500, tia_asp=12.4, tia_ucogs=4.4,
        drv_vol=2140, drv_asp=10.4, drv_ucogs=3.8,
        oth_inf=50000, bb=66500, conn=44500, ind=20000,
        other_gm=0.588, sga=54500, rd=64000, rs=0,
        tax_rate=0.13, tax_min=1500, yld=0.010, drate=0.017, oth=0,
        dso=30, dio=85, dpo=54, capex=5200, sbc=25000,
        da_ppe=2900, da_ia=1400, sh_b=95500, sh_d=103500,
        stdebt=0, ltdebt=100000, fx=0,
    ),
}

# ---------------------------------------------------------------------------
# Excel layout
# ---------------------------------------------------------------------------
BLUE = Font(name="Calibri", size=9, color="0000FF")
BLACK = Font(name="Calibri", size=9, color="000000")
BLUE_B = Font(name="Calibri", size=9, color="0000FF", bold=True)
BLACK_B = Font(name="Calibri", size=9, color="000000", bold=True)
HDR = Font(name="Calibri", size=9, color="FFFFFF", bold=True)
SEC = Font(name="Calibri", size=10, color="FFFFFF", bold=True)
NOTEF = Font(name="Calibri", size=8, color="000000", italic=True)
GRAYF = Font(name="Calibri", size=8, color="666666")
SUB = Font(name="Calibri", size=11, color="000000", bold=True)

FILL_NAVY = PatternFill("solid", fgColor="1F4E79")
FILL_DRV = PatternFill("solid", fgColor="1F4E79")
FILL_PL = PatternFill("solid", fgColor="548235")
FILL_BS = PatternFill("solid", fgColor="833C0C")
FILL_CF = PatternFill("solid", fgColor="5B2C6F")
FILL_CK = PatternFill("solid", fgColor="C00000")
FILL_ACT = PatternFill("solid", fgColor="D6DCE4")
FILL_FC = PatternFill("solid", fgColor="FFF2CC")
FILL_MIX = PatternFill("solid", fgColor="FCE4D6")
FILL_TOT = PatternFill("solid", fgColor="E2EFDA")
FILL_NOTE = PatternFill("solid", fgColor="FFF8E7")
FILL_YELLOW = PatternFill("solid", fgColor="FFFF99")
FILL_TITLE = PatternFill("solid", fgColor="1F4E79")

THIN = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)

C_SEC, C_LINE, C_UNIT = 1, 2, 3
ANN_COL0 = 4
Q_COL0 = ANN_COL0 + len(ANN_YEARS) + 1
NOTES_COL = Q_COL0 + len(QUARTERS)

NF_N, NF_1, NF_2, NF_P, NF_ASP = "#,##0", "#,##0.0", "#,##0.00", "0.0%", "#,##0.00"


def ann_col(year):
    return ANN_COL0 + (year - 2016)


def q_col(y, q):
    return Q_COL0 + (y - 2016) * 4 + (q - 1)


def q_prev(y, q):
    return (y - 1, 4) if q == 1 else (y, q - 1)


def period_fill(year=None, yq=None):
    if year is not None:
        if year <= 2025:
            return FILL_ACT
        if year == 2026:
            return FILL_MIX
        return FILL_FC
    y, q = yq
    return FILL_ACT if is_actual_q(y, q) else FILL_FC


wb = Workbook()
ws = wb.active
ws.title = "MXL 3-Statement Model"

ws.column_dimensions["A"].width = 22
ws.column_dimensions["B"].width = 48
ws.column_dimensions["C"].width = 16
for col in range(ANN_COL0, NOTES_COL):
    ws.column_dimensions[get_column_letter(col)].width = 11.2
ws.column_dimensions[get_column_letter(ANN_COL0 + len(ANN_YEARS))].width = 2.5
ws.column_dimensions[get_column_letter(NOTES_COL)].width = 82

center = Alignment(horizontal="center", vertical="center", wrap_text=True)
left = Alignment(horizontal="left", vertical="center", wrap_text=True)
right = Alignment(horizontal="right", vertical="center")
R = {}


def write_label(row, section, line, unit, note=""):
    ws.cell(row, C_SEC, section).font = GRAYF
    ws.cell(row, C_LINE, line).font = BLACK
    ws.cell(row, C_UNIT, unit).font = GRAYF
    ws.cell(row, NOTES_COL, note).font = NOTEF
    ws.cell(row, NOTES_COL).fill = FILL_NOTE
    ws.cell(row, NOTES_COL).alignment = left


def set_font_val(cell, is_input, fmt=None, bold=False):
    cell.font = (BLUE_B if is_input else BLACK_B) if bold else (BLUE if is_input else BLACK)
    cell.alignment = right
    cell.border = THIN
    if fmt:
        cell.number_format = fmt


def put(row, col, value, is_input, fmt=NF_N, bold=False, fill=None):
    cell = ws.cell(row, col, value)
    set_font_val(cell, is_input, fmt, bold)
    if fill:
        cell.fill = fill
    return cell


def formula(row, col, f, fmt=NF_N, bold=False, fill=None):
    cell = ws.cell(row, col, f)
    set_font_val(cell, False, fmt, bold)
    if fill:
        cell.fill = fill
    return cell


def section_bar(row, title, fill):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
    ws.cell(row, 1, title).font = SEC
    for col in range(1, NOTES_COL + 1):
        ws.cell(row, col).fill = fill
        ws.cell(row, col).font = SEC
    return row


# ===== HEADER =====
ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)
c = ws.cell(1, 1, "MaxLinear, Inc. (NASDAQ: MXL)  —  Integrated 3-Statement Financial Model")
c.font = Font(name="Calibri", size=16, color="FFFFFF", bold=True)
c.fill = FILL_TITLE
for col in range(1, 9):
    ws.cell(1, col).fill = FILL_TITLE

ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=8)
ws.cell(2, 1, "光模块模拟前端：PAM4 DSP / TIA / Driver 量价驱动 | 历史 2016–2025A + 2026H1A | 预测 2026H2–2028E | 单位: $000（另有标注除外）").font = SUB

ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=8)
ws.cell(3, 1, "数据来源: SEC 10-K/10-Q XBRL、Q2'26财报、公司电话会。蓝色=假设/硬编码；黑色=公式。模型日期: 2026-08-18。Fabless RF/混合信号，拐点为数据中心800G/1.6T光学。").font = GRAYF

ws.cell(5, 1, "图例").font = BLACK_B
ws.cell(5, 2, "蓝色字体 = 假设 / 历史硬编码 (hardcode)").font = BLUE
ws.cell(5, 3, "黑色字体 = 公式 / 勾稽").font = BLACK
ws.cell(5, 4, "灰底 = 实际").fill = FILL_ACT
ws.cell(5, 5, "黄底 = 预测").fill = FILL_FC
ws.cell(5, 6, "橙底 = 2026A/E 混合").fill = FILL_MIX
ws.cell(5, 7, "最右列为 Notes").font = NOTEF

for row, label in [(7, "Period type"), (8, "Fiscal year"), (9, "Period"), (10, "Status"), (11, "Days in period")]:
    ws.cell(row, C_LINE, label).font = BLACK_B

for y in ANN_YEARS:
    col = ann_col(y)
    fill = period_fill(year=y)
    put(7, col, "Annual", False, fmt="@", bold=True, fill=fill)
    put(8, col, y, False, fmt="0", bold=True, fill=fill)
    put(9, col, f"FY{y}", False, fmt="@", bold=True, fill=fill)
    status = "Actual" if y <= 2025 else ("Actual+Forecast" if y == 2026 else "Forecast")
    put(10, col, status, False, fmt="@", fill=fill)
    put(11, col, 365, True, fmt="0", fill=fill)

gap = ANN_COL0 + len(ANN_YEARS)
for r in range(7, 12):
    ws.cell(r, gap).fill = FILL_NAVY

for (y, q) in QUARTERS:
    col = q_col(y, q)
    fill = period_fill(yq=(y, q))
    put(7, col, "Quarterly", False, fmt="@", bold=True, fill=fill)
    put(8, col, y, False, fmt="0", bold=True, fill=fill)
    put(9, col, f"Q{q} {y}", False, fmt="@", bold=True, fill=fill)
    put(10, col, "Actual" if is_actual_q(y, q) else "Forecast", False, fmt="@", fill=fill)
    put(11, col, 91 if q != 4 else 92, True, fmt="0", fill=fill)

ws.cell(7, NOTES_COL, "Notes / 假设依据").font = HDR
ws.cell(7, NOTES_COL).fill = FILL_NAVY
ws.merge_cells(start_row=7, start_column=NOTES_COL, end_row=11, end_column=NOTES_COL)
ws.cell(7, NOTES_COL).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
ws.freeze_panes = "D12"
ws.sheet_view.showGridLines = False
ws.page_setup.orientation = "landscape"
ws.page_setup.fitToPage = True
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0
ws.print_title_rows = "1:11"
ws.print_title_cols = "A:C"

# =====================================================================
# DRIVERS
# =====================================================================
row = section_bar(13, "I.  REVENUE DRIVERS  —  PAM4 DSP / TIA / Driver 量价 + Broadband / Connectivity / Industrial", FILL_DRV)
R["sec_drv"] = 13
driver_rows = [
    ("dsp_vol", "DSP", "PAM4 DSP units", "k units", "Keystone 800G + Rushmore 1.6T DSP出货量。公司不披露颗数；历史=光学DSP收入/ASP（公式）；预测=爬坡假设（蓝）。2026光学展望$210–230m，本模型FY26光学约$241m（Q3指引$210–220m公司总收入下的运行速率，略高于年初光学指引）。"),
    ("dsp_asp", "DSP", "PAM4 DSP blended ASP", "$ / unit", "混合ASP。800G CMOS 5nm Keystone工作假设$45–55；1.6T Rushmore $80–120。2026以800G为主~$50；2027–28随1.6T占比提升至$70–86。行业模块BOM中DSP通常为最大芯片成本项。"),
    ("dsp_rev", "DSP", "PAM4 DSP revenue", "$000", "DSP收入=出货量(k)×ASP。2024起从Infrastructure中拆出；2016–23为0（当时基础设施以无线回传/HPA为主）。Q2'26基础设施$85m，光学已占大部分。"),
    ("dsp_ucogs", "DSP", "DSP unit COGS", "$ / unit", "单颗COGS（含TSMC晶圆、封测、特许权）。5nm Keystone成本高于成熟节点；规模与良率改善使GM维持~58%。预测为蓝字假设。"),
    ("dsp_cogs", "DSP", "DSP COGS", "$000", "DSP销货成本=量×单颗COGS（公式）。"),
    ("dsp_gp", "DSP", "DSP gross profit", "$000", "DSP毛利=收入−COGS（公式）。"),
    ("dsp_gm", "DSP", "DSP gross margin", "%", "DSP毛利率=毛利/收入（公式）。"),
    ("tia_vol", "TIA", "TIA units", "k units", "Washington等模拟TIA出货。管理层：初始收入2027、放量2028。2026仍小（随800G模块配套）。未披露，按光学收入12–18%拆分。"),
    ("tia_asp", "TIA", "TIA ASP", "$ / unit", "TIA ASP工作假设$10–12.4。SiGe/BiCMOS模拟，ASP低于DSP。"),
    ("tia_rev", "TIA", "TIA revenue", "$000", "TIA收入=量×ASP（预测）或光学拆分（历史）。"),
    ("tia_ucogs", "TIA", "TIA unit COGS", "$ / unit", "TIA单颗成本。模拟产品毛利率通常略高于CMOS DSP（~62%）。"),
    ("tia_cogs", "TIA", "TIA COGS", "$000", "TIA COGS=量×单颗COGS。"),
    ("drv_vol", "Driver", "Laser driver units", "k units", "激光Driver出货。Rushmore部分集成driver，独立driver随800G/1.6T模块配套。管理层称Annapolis retimer/analog 2027起步。"),
    ("drv_asp", "Driver", "Driver ASP", "$ / unit", "Driver ASP工作假设$8–10.4。"),
    ("drv_rev", "Driver", "Driver revenue", "$000", "Driver收入=量×ASP。"),
    ("drv_ucogs", "Driver", "Driver unit COGS", "$ / unit", "Driver单颗成本，隐含GM约60%。"),
    ("drv_cogs", "Driver", "Driver COGS", "$000", "Driver COGS=量×单颗COGS。"),
    ("opt_rev", "Optical", "Optical analog revenue (DSP+TIA+Driver)", "$000", "光学模拟合计。应对齐公司FY26光学$210–230m量级；本模型因Q2基础设施$85m+Q3公司指引而取~$241m。2027 Rushmore+TIA/Driver，$475m；2028 $795m。"),
    ("oth_inf", "Infra", "Other infrastructure (wireless, HPA, IP)", "$000", "基础设施中非光学：无线回传、微波、HPA、IP授权。历史=报告Infrastructure−光学拆分。"),
    ("bb", "Other", "Broadband revenue", "$000", "宽带（DOCSIS/网关/连接家庭）。2021–22高峰~$493m，2024谷底$117m，2025回升$204m。Q2'26 $44.9m。预测随Wi-Fi 7/DOCSIS缓慢修复。2016–20为估算映射。"),
    ("conn", "Other", "Connectivity revenue", "$000", "连接（Wi-Fi/以太网等）。2022高峰$304m，2024 $56m，Q2'26 $24.0m。预测随Wi-Fi 7温和增长。"),
    ("ind", "Other", "Industrial / multi-market", "$000", "工业与多市场。2025仅$37m。Q2'26 $15.0m。预测维持低双位数百万/季。"),
    ("tot_drv", "Total", "Total revenue (drivers)", "$000", "驱动加总，必须等于P&L收入（见勾稽）。"),
    ("mix_opt", "Mix", "Optical % of total revenue", "%", "光学收入占比。Q2'26约31%；预测2028升至~50%+，是估值核心变量。"),
    ("mix_dsp", "Mix", "DSP % of optical", "%", "光学中DSP占比。2026~80%，此后随TIA/Driver attach下降。"),
]
for key, sec, line, unit, note in driver_rows:
    row += 1
    R[key] = row
    write_label(row, sec, line, unit, note)
    if key in ("tot_drv", "opt_rev"):
        ws.cell(row, C_LINE).font = BLACK_B

# Annual drivers
for y in ANN_YEARS:
    col = ann_col(y)
    fill = period_fill(year=y)
    cl = get_column_letter(col)
    qcs = [get_column_letter(q_col(y, q)) for q in (1, 2, 3, 4)]
    if y <= 2025:
        md, mt, mv = opt_mix(y)
        opt = OPT_ANN[y]
        dsp, tia, drv = round(opt * md), round(opt * mt), opt - round(opt * md) - round(opt * mt)
        seg = SEG_ANN[y]
        put(R["dsp_rev"], col, dsp, True, NF_N, fill=fill)
        put(R["tia_rev"], col, tia, True, NF_N, fill=fill)
        put(R["drv_rev"], col, drv, True, NF_N, fill=fill)
        put(R["dsp_asp"], col, ANN_ASP[y]["dsp"], True, NF_ASP, fill=fill)
        put(R["tia_asp"], col, ANN_ASP[y]["tia"], True, NF_ASP, fill=fill)
        put(R["drv_asp"], col, ANN_ASP[y]["drv"], True, NF_ASP, fill=fill)
        put(R["dsp_ucogs"], col, ANN_UCOGS[y]["dsp"], True, NF_ASP, fill=fill)
        put(R["tia_ucogs"], col, ANN_UCOGS[y]["tia"], True, NF_ASP, fill=fill)
        put(R["drv_ucogs"], col, ANN_UCOGS[y]["drv"], True, NF_ASP, fill=fill)
        formula(R["dsp_vol"], col, f'=IF({cl}{R["dsp_asp"]}=0,0,{cl}{R["dsp_rev"]}/{cl}{R["dsp_asp"]})', NF_1, fill=fill)
        formula(R["tia_vol"], col, f'=IF({cl}{R["tia_asp"]}=0,0,{cl}{R["tia_rev"]}/{cl}{R["tia_asp"]})', NF_1, fill=fill)
        formula(R["drv_vol"], col, f'=IF({cl}{R["drv_asp"]}=0,0,{cl}{R["drv_rev"]}/{cl}{R["drv_asp"]})', NF_1, fill=fill)
        put(R["oth_inf"], col, seg["inf"] - opt, True, NF_N, fill=fill)
        put(R["bb"], col, seg["bb"], True, NF_N, fill=fill)
        put(R["conn"], col, seg["conn"], True, NF_N, fill=fill)
        put(R["ind"], col, seg["ind"], True, NF_N, fill=fill)
        formula(R["dsp_cogs"], col, f'={cl}{R["dsp_vol"]}*{cl}{R["dsp_ucogs"]}', NF_N, fill=fill)
        formula(R["tia_cogs"], col, f'={cl}{R["tia_vol"]}*{cl}{R["tia_ucogs"]}', NF_N, fill=fill)
        formula(R["drv_cogs"], col, f'={cl}{R["drv_vol"]}*{cl}{R["drv_ucogs"]}', NF_N, fill=fill)
    else:
        def qsum(k):
            return "=" + "+".join(f"{qc}{R[k]}" for qc in qcs)
        for k in ["dsp_vol", "dsp_rev", "dsp_cogs", "tia_vol", "tia_rev", "tia_cogs",
                  "drv_vol", "drv_rev", "drv_cogs", "oth_inf", "bb", "conn", "ind"]:
            formula(R[k], col, qsum(k), NF_1 if "vol" in k else NF_N, fill=fill)
        formula(R["dsp_asp"], col, f'=IF({cl}{R["dsp_vol"]}=0,0,{cl}{R["dsp_rev"]}/{cl}{R["dsp_vol"]})', NF_ASP, fill=fill)
        formula(R["tia_asp"], col, f'=IF({cl}{R["tia_vol"]}=0,0,{cl}{R["tia_rev"]}/{cl}{R["tia_vol"]})', NF_ASP, fill=fill)
        formula(R["drv_asp"], col, f'=IF({cl}{R["drv_vol"]}=0,0,{cl}{R["drv_rev"]}/{cl}{R["drv_vol"]})', NF_ASP, fill=fill)
        formula(R["dsp_ucogs"], col, f'=IF({cl}{R["dsp_vol"]}=0,0,{cl}{R["dsp_cogs"]}/{cl}{R["dsp_vol"]})', NF_ASP, fill=fill)
        formula(R["tia_ucogs"], col, f'=IF({cl}{R["tia_vol"]}=0,0,{cl}{R["tia_cogs"]}/{cl}{R["tia_vol"]})', NF_ASP, fill=fill)
        formula(R["drv_ucogs"], col, f'=IF({cl}{R["drv_vol"]}=0,0,{cl}{R["drv_cogs"]}/{cl}{R["drv_vol"]})', NF_ASP, fill=fill)
    formula(R["dsp_gp"], col, f'={cl}{R["dsp_rev"]}-{cl}{R["dsp_cogs"]}', NF_N, fill=fill)
    formula(R["dsp_gm"], col, f'=IF({cl}{R["dsp_rev"]}=0,0,{cl}{R["dsp_gp"]}/{cl}{R["dsp_rev"]})', NF_P, fill=fill)
    formula(R["opt_rev"], col, f'={cl}{R["dsp_rev"]}+{cl}{R["tia_rev"]}+{cl}{R["drv_rev"]}', NF_N, True, FILL_TOT)
    formula(R["tot_drv"], col, f'={cl}{R["opt_rev"]}+{cl}{R["oth_inf"]}+{cl}{R["bb"]}+{cl}{R["conn"]}+{cl}{R["ind"]}', NF_N, True, FILL_TOT)
    formula(R["mix_opt"], col, f'=IF({cl}{R["tot_drv"]}=0,0,{cl}{R["opt_rev"]}/{cl}{R["tot_drv"]})', NF_P, fill=fill)
    formula(R["mix_dsp"], col, f'=IF({cl}{R["opt_rev"]}=0,0,{cl}{R["dsp_rev"]}/{cl}{R["opt_rev"]})', NF_P, fill=fill)

# Quarterly drivers
for (y, q) in QUARTERS:
    col = q_col(y, q)
    fill = period_fill(yq=(y, q))
    cl = get_column_letter(col)
    if is_actual_q(y, q):
        pr = allocate_product(y, q)
        put(R["dsp_rev"], col, pr["dsp"], True, NF_N, fill=fill)
        put(R["tia_rev"], col, pr["tia"], True, NF_N, fill=fill)
        put(R["drv_rev"], col, pr["drv"], True, NF_N, fill=fill)
        put(R["oth_inf"], col, pr["oth_inf"], True, NF_N, fill=fill)
        put(R["bb"], col, pr["bb"], True, NF_N, fill=fill)
        put(R["conn"], col, pr["conn"], True, NF_N, fill=fill)
        put(R["ind"], col, pr["ind"], True, NF_N, fill=fill)
        put(R["dsp_asp"], col, q_asp(y, q, "dsp"), True, NF_ASP, fill=fill)
        put(R["tia_asp"], col, q_asp(y, q, "tia"), True, NF_ASP, fill=fill)
        put(R["drv_asp"], col, q_asp(y, q, "drv"), True, NF_ASP, fill=fill)
        put(R["dsp_ucogs"], col, q_ucogs(y, q, "dsp"), True, NF_ASP, fill=fill)
        put(R["tia_ucogs"], col, q_ucogs(y, q, "tia"), True, NF_ASP, fill=fill)
        put(R["drv_ucogs"], col, q_ucogs(y, q, "drv"), True, NF_ASP, fill=fill)
        formula(R["dsp_vol"], col, f'=IF({cl}{R["dsp_asp"]}=0,0,{cl}{R["dsp_rev"]}/{cl}{R["dsp_asp"]})', NF_1, fill=fill)
        formula(R["tia_vol"], col, f'=IF({cl}{R["tia_asp"]}=0,0,{cl}{R["tia_rev"]}/{cl}{R["tia_asp"]})', NF_1, fill=fill)
        formula(R["drv_vol"], col, f'=IF({cl}{R["drv_asp"]}=0,0,{cl}{R["drv_rev"]}/{cl}{R["drv_asp"]})', NF_1, fill=fill)
        formula(R["dsp_cogs"], col, f'={cl}{R["dsp_vol"]}*{cl}{R["dsp_ucogs"]}', NF_N, fill=fill)
        formula(R["tia_cogs"], col, f'={cl}{R["tia_vol"]}*{cl}{R["tia_ucogs"]}', NF_N, fill=fill)
        formula(R["drv_cogs"], col, f'={cl}{R["drv_vol"]}*{cl}{R["drv_ucogs"]}', NF_N, fill=fill)
    else:
        f = FCST[(y, q)]
        put(R["dsp_vol"], col, f["dsp_vol"], True, NF_1, fill=fill)
        put(R["dsp_asp"], col, f["dsp_asp"], True, NF_ASP, fill=fill)
        put(R["dsp_ucogs"], col, f["dsp_ucogs"], True, NF_ASP, fill=fill)
        put(R["tia_vol"], col, f["tia_vol"], True, NF_1, fill=fill)
        put(R["tia_asp"], col, f["tia_asp"], True, NF_ASP, fill=fill)
        put(R["tia_ucogs"], col, f["tia_ucogs"], True, NF_ASP, fill=fill)
        put(R["drv_vol"], col, f["drv_vol"], True, NF_1, fill=fill)
        put(R["drv_asp"], col, f["drv_asp"], True, NF_ASP, fill=fill)
        put(R["drv_ucogs"], col, f["drv_ucogs"], True, NF_ASP, fill=fill)
        formula(R["dsp_rev"], col, f'={cl}{R["dsp_vol"]}*{cl}{R["dsp_asp"]}', NF_N, fill=fill)
        formula(R["tia_rev"], col, f'={cl}{R["tia_vol"]}*{cl}{R["tia_asp"]}', NF_N, fill=fill)
        formula(R["drv_rev"], col, f'={cl}{R["drv_vol"]}*{cl}{R["drv_asp"]}', NF_N, fill=fill)
        formula(R["dsp_cogs"], col, f'={cl}{R["dsp_vol"]}*{cl}{R["dsp_ucogs"]}', NF_N, fill=fill)
        formula(R["tia_cogs"], col, f'={cl}{R["tia_vol"]}*{cl}{R["tia_ucogs"]}', NF_N, fill=fill)
        formula(R["drv_cogs"], col, f'={cl}{R["drv_vol"]}*{cl}{R["drv_ucogs"]}', NF_N, fill=fill)
        put(R["oth_inf"], col, f["oth_inf"], True, NF_N, fill=fill)
        put(R["bb"], col, f["bb"], True, NF_N, fill=fill)
        put(R["conn"], col, f["conn"], True, NF_N, fill=fill)
        put(R["ind"], col, f["ind"], True, NF_N, fill=fill)
    formula(R["dsp_gp"], col, f'={cl}{R["dsp_rev"]}-{cl}{R["dsp_cogs"]}', NF_N, fill=fill)
    formula(R["dsp_gm"], col, f'=IF({cl}{R["dsp_rev"]}=0,0,{cl}{R["dsp_gp"]}/{cl}{R["dsp_rev"]})', NF_P, fill=fill)
    formula(R["opt_rev"], col, f'={cl}{R["dsp_rev"]}+{cl}{R["tia_rev"]}+{cl}{R["drv_rev"]}', NF_N, True, FILL_TOT)
    formula(R["tot_drv"], col, f'={cl}{R["opt_rev"]}+{cl}{R["oth_inf"]}+{cl}{R["bb"]}+{cl}{R["conn"]}+{cl}{R["ind"]}', NF_N, True, FILL_TOT)
    formula(R["mix_opt"], col, f'=IF({cl}{R["tot_drv"]}=0,0,{cl}{R["opt_rev"]}/{cl}{R["tot_drv"]})', NF_P, fill=fill)
    formula(R["mix_dsp"], col, f'=IF({cl}{R["opt_rev"]}=0,0,{cl}{R["dsp_rev"]}/{cl}{R["opt_rev"]})', NF_P, fill=fill)

# =====================================================================
# ASSUMPTIONS
# =====================================================================
row = R["mix_dsp"] + 2
row = section_bar(row, "II.  MARGIN / P&L / BALANCE SHEET ASSUMPTIONS  （蓝色=可调假设）", FILL_DRV)
R["sec_ass"] = row
ass_items = [
    ("gm_co", "P&L", "Company gross margin (implied)", "%", "历史=GP/Rev；预测由DSP/TIA/Driver单颗COGS + 非光学毛利率加权得出。Q2'26 GAAP GM 57.8%；Q3指引57–60%。"),
    ("other_gm", "P&L", "Non-optical blended gross margin", "%", "宽带/连接/工业/其他基础设施综合毛利率。历史反推；预测57.5–58.8%。光学mix提升后公司GM向59%靠近。"),
    ("sga_d", "P&L", "SG&A", "$000", "历史硬编码；预测绝对额（半固定）。Q2'26 $45.8m；Q3 opex指引$98–104m（含R&D/重组）。"),
    ("rd_d", "P&L", "R&D", "$000", "历史硬编码（XBRL: R&D excluding acquired IPR&D）。光学平台（5nm Keystone / 1.6T Rushmore）维持高研发。预测$55–64m/季。"),
    ("rs_d", "P&L", "Restructuring / impairment", "$000", "重组与减值。2024 $53.4m、2025 $24.5m。预测Q3'26后趋近0。"),
    ("tax_r", "P&L", "Effective tax rate (on pretax profit)", "%", "亏损期现金税主要为境外最低税（Q3指引约$1.5m）。盈利期用12–13% ETR（NOLs）。"),
    ("tax_min", "P&L", "Minimum tax if EBT≤0", "$000", "亏损季最低税。Q3'26指引$1.5m。"),
    ("yld", "P&L", "Cash quarterly yield", "%", "利息收入=期初现金×季收益率。避免循环引用。约3.2–4.0%年化。"),
    ("drate", "P&L", "Cost of debt (quarterly)", "%", "定期贷款+循环贷成本，约7.2%年化（1.8%/季）。Q2'26利息支出$2.27m。"),
    ("oth_d", "P&L", "Other income / (expense)", "$000", "汇兑及其他。预测0。"),
    ("dso", "BS", "DSO (AR days)", "days", "应收账款天数。Q2'26约27.5天（偏低，渠道/分销）。预测光学直销升至~30–34天。"),
    ("dio", "BS", "DIO (Inventory days)", "days", "存货天数。Q2'26约135天（光学ramp备货）。预测随周转降至~85天。"),
    ("dpo", "BS", "DPO (AP days)", "days", "应付天数。预测维持~52–54天。"),
    ("capex_d", "CF", "Capex", "$000", "Fabless资本开支低。FY25 $12.6m；Q2'26 $2.3m。预测$3.5–5.2m/季。"),
    ("da_ppe", "CF", "D&A — PPE", "$000", "PPE折旧。预测~$2.9–3.5m/季。PPE=上期+Capex−本项。"),
    ("da_ia", "CF", "D&A — intangibles", "$000", "收购无形资产摊销。Q1'26 D&A合计$10.9m，大部分为摊销。随账面摊完而下降。IA=max(期初−本项, $20m地板)。"),
    ("sbc_d", "CF", "Stock-based compensation", "$000", "股份支付。FY25 $77.1m；Q1'26 $20.0m。加回CFO并增加APIC。预测$21–25m/季。"),
    ("sh_b", "P&L", "Weighted-avg / ending shares (diluted used below)", "k shares", "Q2'26基本90.0m / 稀释97.3m；Q3指引稀释99m。"),
]
for key, sec, line, unit, note in ass_items:
    row += 1
    R[key] = row
    write_label(row, sec, line, unit, note)

# =====================================================================
# P&L
# =====================================================================
row = R["sh_b"] + 2
row = section_bar(row, "III.  CONSOLIDATED P&L  （利润表）", FILL_PL)
pl_items = [
    ("rev", "Revenue", "$000", "营业收入。历史=10-K/10-Q；预测=驱动表Total revenue。Q3'26指引$210–220m。"),
    ("cogs", "Cost of revenue", "$000", "历史硬编码；预测=光学COGS + 非光学收入×(1−非光学GM)。"),
    ("gp", "Gross profit", "$000", "毛利=收入−成本。"),
    ("gpm", "Gross margin", "%", "毛利率。Q2'26 57.8%。"),
    ("rd", "Research & development", "$000", "见假设。"),
    ("sga", "Selling, general & administrative", "$000", "见假设。"),
    ("rs", "Restructuring, impairment & other opex", "$000", "重组/减值/IPR&D。历史含使OI钉住10-K的残差。"),
    ("opex", "Total operating expenses", "$000", "R&D+SG&A+重组。"),
    ("oi", "Operating income", "$000", "营业利润=毛利−营业费用。"),
    ("om", "Operating margin", "%", "营业利润率。Q2'26 GAAP −2.5%，Non-GAAP 22.3%。"),
    ("ii", "Interest income", "$000", "利息收入。预测=期初现金×收益率。"),
    ("ie", "Interest expense", "$000", "利息支出。预测=期初(ST+LT债务)×债务成本。"),
    ("intn", "Interest income / (expense), net", "$000", "利息净额=收入−支出。"),
    ("oth", "Other income / (expense), net", "$000", "其他。历史=EBT−OI−利息净额（配平）；预测为假设。"),
    ("ebt", "Income before tax", "$000", "税前利润。"),
    ("tax", "Provision for income taxes", "$000", "所得税。Q1'26大额费用$26.5m（递延税资产），Q2'26利益$8.3m。预测：盈利×ETR，亏损取最低税。"),
    ("ni", "Net income", "$000", "净利润=EBT−Tax。无少数股东。"),
    ("eps", "EPS diluted", "$ / sh", "稀释EPS=NI/稀释股本。"),
    ("shd", "Weighted-avg diluted shares", "k", "稀释加权平均股本。亏损期=基本股本。"),
]
for key, line, unit, note in pl_items:
    row += 1
    R[key] = row
    write_label(row, "P&L", line, unit, note)
    if key in ("gp", "oi", "ni"):
        ws.cell(row, C_LINE).font = BLACK_B

# Annual P&L
for y in ANN_YEARS:
    col = ann_col(y)
    fill = period_fill(year=y)
    cl = get_column_letter(col)
    qcs = [get_column_letter(q_col(y, q)) for q in (1, 2, 3, 4)]

    def qsum(k):
        return "=" + "+".join(f"{qc}{R[k]}" for qc in qcs)

    if y <= 2025:
        a = ANN[y]
        put(R["rev"], col, a["rev"], True, NF_N, True, fill)
        put(R["cogs"], col, a["cogs"], True, NF_N, fill=fill)
        formula(R["gp"], col, f'={cl}{R["rev"]}-{cl}{R["cogs"]}', NF_N, True, FILL_TOT)
        formula(R["gpm"], col, f'=IF({cl}{R["rev"]}=0,0,{cl}{R["gp"]}/{cl}{R["rev"]})', NF_P, fill=fill)
        put(R["rd"], col, a["rd"], True, NF_N, fill=fill)
        put(R["sga"], col, a["sga"], True, NF_N, fill=fill)
        put(R["rs"], col, a["rs"] + a["imp"], True, NF_N, fill=fill)
        formula(R["opex"], col, f'={cl}{R["rd"]}+{cl}{R["sga"]}+{cl}{R["rs"]}', NF_N, fill=fill)
        formula(R["oi"], col, f'={cl}{R["gp"]}-{cl}{R["opex"]}', NF_N, True, FILL_TOT)
        formula(R["om"], col, f'=IF({cl}{R["rev"]}=0,0,{cl}{R["oi"]}/{cl}{R["rev"]})', NF_P, fill=fill)
        put(R["ii"], col, a["ii"], True, NF_N, fill=fill)
        put(R["ie"], col, a["ie"], True, NF_N, fill=fill)
        formula(R["intn"], col, f'={cl}{R["ii"]}-{cl}{R["ie"]}', NF_N, fill=fill)
        put(R["ebt"], col, a["ebt"], True, NF_N, fill=fill)
        formula(R["oth"], col, f'={cl}{R["ebt"]}-{cl}{R["oi"]}-{cl}{R["intn"]}', NF_N, fill=fill)
        put(R["tax"], col, a["tax"], True, NF_N, fill=fill)
        formula(R["ni"], col, f'={cl}{R["ebt"]}-{cl}{R["tax"]}', NF_N, True, FILL_TOT)
        put(R["shd"], col, a["sh_d"], True, NF_N, fill=fill)
        formula(R["eps"], col, f'=IF({cl}{R["shd"]}=0,0,{cl}{R["ni"]}/{cl}{R["shd"]})', NF_2, fill=fill)
        formula(R["gm_co"], col, f'={cl}{R["gpm"]}', NF_P, fill=fill)
        formula(R["other_gm"], col, f'=IF({cl}{R["rev"]}-{cl}{R["opt_rev"]}=0,0,({cl}{R["gp"]}-({cl}{R["dsp_rev"]}-{cl}{R["dsp_cogs"]}+{cl}{R["tia_rev"]}-{cl}{R["tia_cogs"]}+{cl}{R["drv_rev"]}-{cl}{R["drv_cogs"]}))/MAX({cl}{R["rev"]}-{cl}{R["opt_rev"]},1))', NF_P, fill=fill)
        formula(R["sga_d"], col, f'={cl}{R["sga"]}', NF_N, fill=fill)
        formula(R["rd_d"], col, f'={cl}{R["rd"]}', NF_N, fill=fill)
        formula(R["rs_d"], col, f'={cl}{R["rs"]}', NF_N, fill=fill)
        formula(R["tax_r"], col, f'=IF({cl}{R["ebt"]}=0,0,{cl}{R["tax"]}/{cl}{R["ebt"]})', NF_P, fill=fill)
        put(R["tax_min"], col, 0, True, NF_N, fill=fill)
        put(R["yld"], col, 0.008, True, NF_P, fill=fill)
        put(R["drate"], col, 0.018, True, NF_P, fill=fill)
        formula(R["oth_d"], col, f'={cl}{R["oth"]}', NF_N, fill=fill)
        put(R["dso"], col, round(QBS[(y, 4)]["ar"] / a["rev"] * 365, 1), True, NF_1, fill=fill)
        put(R["dio"], col, round(QBS[(y, 4)]["inv"] / a["cogs"] * 365, 1), True, NF_1, fill=fill)
        put(R["dpo"], col, round(QBS[(y, 4)]["ap"] / a["cogs"] * 365, 1), True, NF_1, fill=fill)
        put(R["capex_d"], col, a["capex"], True, NF_N, fill=fill)
        put(R["da_ppe"], col, round(a["da"] * 0.35), True, NF_N, fill=fill)
        put(R["da_ia"], col, a["da"] - round(a["da"] * 0.35), True, NF_N, fill=fill)
        put(R["sbc_d"], col, a["sbc"], True, NF_N, fill=fill)
        put(R["sh_b"], col, a["sh_d"], True, NF_N, fill=fill)
    else:
        for k in ["rev", "cogs", "rd", "sga", "rs", "ii", "ie", "oth", "tax", "ni"]:
            formula(R[k], col, qsum(k), NF_N, k in ("rev", "ni"), FILL_TOT if k in ("rev", "ni") else fill)
        formula(R["gp"], col, f'={cl}{R["rev"]}-{cl}{R["cogs"]}', NF_N, True, FILL_TOT)
        formula(R["gpm"], col, f'=IF({cl}{R["rev"]}=0,0,{cl}{R["gp"]}/{cl}{R["rev"]})', NF_P, fill=fill)
        formula(R["opex"], col, f'={cl}{R["rd"]}+{cl}{R["sga"]}+{cl}{R["rs"]}', NF_N, fill=fill)
        formula(R["oi"], col, f'={cl}{R["gp"]}-{cl}{R["opex"]}', NF_N, True, FILL_TOT)
        formula(R["om"], col, f'=IF({cl}{R["rev"]}=0,0,{cl}{R["oi"]}/{cl}{R["rev"]})', NF_P, fill=fill)
        formula(R["intn"], col, f'={cl}{R["ii"]}-{cl}{R["ie"]}', NF_N, fill=fill)
        formula(R["ebt"], col, f'={cl}{R["oi"]}+{cl}{R["intn"]}+{cl}{R["oth"]}', NF_N, fill=fill)
        formula(R["shd"], col, f'={qcs[3]}{R["shd"]}', NF_N, fill=fill)
        formula(R["eps"], col, f'=IF({cl}{R["shd"]}=0,0,{cl}{R["ni"]}/{cl}{R["shd"]})', NF_2, fill=fill)
        formula(R["gm_co"], col, f'={cl}{R["gpm"]}', NF_P, fill=fill)
        formula(R["other_gm"], col, f'=IF({cl}{R["rev"]}-{cl}{R["opt_rev"]}=0,0,({cl}{R["gp"]}-({cl}{R["dsp_rev"]}-{cl}{R["dsp_cogs"]}+{cl}{R["tia_rev"]}-{cl}{R["tia_cogs"]}+{cl}{R["drv_rev"]}-{cl}{R["drv_cogs"]}))/MAX({cl}{R["rev"]}-{cl}{R["opt_rev"]},1))', NF_P, fill=fill)
        for k in ["sga_d", "rd_d", "rs_d", "oth_d", "capex_d", "da_ppe", "da_ia", "sbc_d"]:
            formula(R[k], col, qsum(k), NF_N, fill=fill)
        formula(R["tax_r"], col, f'=IF({cl}{R["ebt"]}=0,0,{cl}{R["tax"]}/{cl}{R["ebt"]})', NF_P, fill=fill)
        formula(R["tax_min"], col, f'={qcs[3]}{R["tax_min"]}', NF_N, fill=fill)
        put(R["yld"], col, 0.009, True, NF_P, fill=fill)
        put(R["drate"], col, 0.018, True, NF_P, fill=fill)
        formula(R["dso"], col, f'={qcs[3]}{R["dso"]}', NF_1, fill=fill)
        formula(R["dio"], col, f'={qcs[3]}{R["dio"]}', NF_1, fill=fill)
        formula(R["dpo"], col, f'={qcs[3]}{R["dpo"]}', NF_1, fill=fill)
        formula(R["sh_b"], col, f'={qcs[3]}{R["sh_b"]}', NF_N, fill=fill)

# =====================================================================
# BS
# =====================================================================
row = R["shd"] + 2
row = section_bar(row, "IV.  CONSOLIDATED BALANCE SHEET  （资产负债表）", FILL_BS)
bs_items = [
    ("cash", "Cash & cash equivalents", "$000", "货币资金。历史硬编码；预测=CFS期末现金。Q2'26 $64.8m（偏紧，当季提取循环贷$20m）。"),
    ("ar_b", "Accounts receivable", "$000", "应收账款。预测=DSO/天数×收入。"),
    ("inv_b", "Inventories", "$000", "存货。预测=DIO/天数×COGS。Q2'26 $105.5m（光学备货）。"),
    ("oca", "Other current assets", "$000", "其他流动资产=CA−现金−AR−存货（预付、受限现金等）。预测维持Q2'26水平。"),
    ("tca", "Total current assets", "$000", "流动资产合计。"),
    ("ppe_b", "Property, plant & equipment, net", "$000", "PPE净额。预测=上期+Capex−PPE折旧。"),
    ("gw_b", "Goodwill", "$000", "商誉。Intel/Exar等收购形成，Q2'26 $318.6m。预测持平（不测减值）。"),
    ("ia_b", "Intangible assets, net", "$000", "无形资产净额。预测=max(上期−IA摊销, $20m)。"),
    ("onca", "Other non-current assets (DTA, ROU, other)", "$000", "其他非流动：递延税资产、使用权资产等。预测持平Q2'26。"),
    ("tassets", "TOTAL ASSETS", "$000", "资产总计。必须=负债+权益。"),
    ("ap_b", "Accounts payable", "$000", "应付账款。预测=DPO/天数×COGS。"),
    ("stdebt", "Short-term debt / revolver", "$000", "短期债务。Q2'26提取循环贷$20m。预测随FCF偿还。"),
    ("ocl", "Other current liabilities", "$000", "其他流动负债=CL−AP−短期债务。预测持平。"),
    ("tcl", "Total current liabilities", "$000", "流动负债合计。"),
    ("ltdebt", "Long-term debt", "$000", "长期债务（定期贷款账面，含发行成本摊销微增）。Q2'26 $123.9m。2028起偿还。"),
    ("oncl", "Other non-current liabilities", "$000", "其他非流动负债（租赁等）。预测持平。"),
    ("tliab", "TOTAL LIABILITIES", "$000", "负债合计。"),
    ("oeq", "Equity ex-RE (APIC, common, AOCI)", "$000", "除留存收益外的权益。预测=上期+SBC。"),
    ("re_b", "Retained earnings / (accumulated deficit)", "$000", "留存收益。预测=上期+净利润（无股息）。"),
    ("teq", "TOTAL STOCKHOLDERS' EQUITY", "$000", "股东权益。"),
    ("tle", "TOTAL LIABILITIES + EQUITY", "$000", "负债+权益，必须=总资产。"),
]
for key, line, unit, note in bs_items:
    row += 1
    R[key] = row
    write_label(row, "BS", line, unit, note)
    if key in ("tassets", "tliab", "teq", "tle"):
        ws.cell(row, C_LINE).font = BLACK_B

# =====================================================================
# CF
# =====================================================================
row = R["tle"] + 2
row = section_bar(row, "V.  CASH FLOW STATEMENT  （现金流量表）", FILL_CF)
cf_items = [
    ("cf_ni", "Net income", "$000", "取自P&L净利润。"),
    ("cf_da", "Depreciation & amortization", "$000", "加回D&A（PPE+IA）。"),
    ("cf_sbc", "Stock-based compensation", "$000", "加回SBC。"),
    ("cf_ar", "Change in accounts receivable", "$000", "−ΔAR。"),
    ("cf_inv", "Change in inventories", "$000", "−ΔInv。"),
    ("cf_oca", "Change in other current assets", "$000", "−ΔOCA。"),
    ("cf_ap", "Change in accounts payable", "$000", "+ΔAP。"),
    ("cf_ocl", "Change in other operating liabilities", "$000", "+ΔOCL+ΔONCL。"),
    ("cfo", "Cash from operations (CFO)", "$000", "经营活动现金流。历史年度用10-K报告值。"),
    ("cf_capex", "Capital expenditures", "$000", "−Capex。"),
    ("cf_onca", "Change in other non-current assets", "$000", "−Δ(GW+IA+ONCA)。预测IA下降为非现金（已在D&A加回），本行对IA不重复计入。"),
    ("cfi", "Cash from investing (CFI)", "$000", "投资活动现金流。"),
    ("cf_debt", "Net debt issuance / (repayment)", "$000", "ΔST+ΔLT。"),
    ("cf_eq", "Equity issuance / other financing", "$000", "股权融资及其他。历史用CFF残差。预测0。"),
    ("cff", "Cash from financing (CFF)", "$000", "筹资活动现金流。"),
    ("cf_fx", "FX / other plug to roll cash (historical)", "$000", "历史：使现金滚动至BS现金（含受限现金口径差）。预测0。"),
    ("d_cash", "Net change in cash", "$000", "CFO+CFI+CFF+FX。"),
    ("cash_o", "Opening cash", "$000", "期初现金。"),
    ("cash_c", "Closing cash (CFS)", "$000", "期末现金。应等于BS现金。"),
]
for key, line, unit, note in cf_items:
    row += 1
    R[key] = row
    write_label(row, "CF", line, unit, note)
    if key in ("cfo", "cfi", "cff", "d_cash", "cash_c"):
        ws.cell(row, C_LINE).font = BLACK_B

row = R["cash_c"] + 2
row = section_bar(row, "VI.  INTEGRITY CHECKS  （三表配平 / 勾稽，目标=0）", FILL_CK)
ck_items = [
    ("ck_bs", "BS: Assets − (Liabilities + Equity)", "$000", "必须为0。"),
    ("ck_cash", "Cash: BS cash − CFS closing cash", "$000", "必须为0。"),
    ("ck_re", "RE rollforward error", "$000", "期末RE −（期初RE + NI）。历史因AOCI/会计重分类可能有小差异；预测必须为0。"),
    ("ck_rev", "Revenue: P&L − Drivers total", "$000", "收入驱动必须钉住P&L。"),
    ("ck_ni", "NI: P&L NI vs EBT − Tax", "$000", "应为主0。"),
    ("ck_cfid", "CFS identity: CFO+CFI+CFF+FX − ΔCash", "$000", "必须为0。"),
    ("ck_annq", "Annual revenue − sum of 4 quarters", "$000", "年度应对齐四季之和。"),
]
for key, line, unit, note in ck_items:
    row += 1
    R[key] = row
    write_label(row, "Check", line, unit, note)

last_row = R["ck_annq"] + 2
ws.cell(last_row, 1, "免责声明 / Disclaimer").font = BLACK_B
ws.merge_cells(start_row=last_row + 1, start_column=1, end_row=last_row + 4, end_column=8)
ws.cell(last_row + 1, 1, (
    "本模型仅供投资研究讨论，不构成投资建议。历史数据来自SEC申报（XBRL/10-K/10-Q）及公司财报；"
    "DSP/TIA/Driver出货量与ASP公司未披露，为分析师根据Infrastructure分部、FY26光学$210–230m展望、"
    "Q3'26收入指引$210–220m及行业800G/1.6T模块BOM反推。预测依赖Keystone爬坡、Rushmore 2027量产、"
    "TIA/Driver 2027–28放量、宽带温和修复及无重大商誉减值。竞争（Marvell/Broadcom DSP）、制程成本、客户认证与宏观需求可能导致显著偏离。"
)).font = GRAYF
ws.cell(last_row + 1, 1).alignment = Alignment(wrap_text=True, vertical="top")

OCA_Q22 = QBS[(2026, 2)]["ca"] - QBS[(2026, 2)]["cash"] - QBS[(2026, 2)]["ar"] - QBS[(2026, 2)]["inv"]
ONCA_Q22 = QBS[(2026, 2)]["assets"] - QBS[(2026, 2)]["ca"] - QBS[(2026, 2)]["ppe"] - QBS[(2026, 2)]["gw"] - QBS[(2026, 2)]["ia"]
OCL_Q22 = QBS[(2026, 2)]["cl"] - QBS[(2026, 2)]["ap"] - QBS[(2026, 2)]["std"]
ONCL_Q22 = QBS[(2026, 2)]["liab"] - QBS[(2026, 2)]["cl"] - QBS[(2026, 2)]["ltd"]

# ---------- Annual BS / CF / checks ----------
for y in ANN_YEARS:
    col = ann_col(y)
    fill = period_fill(year=y)
    cl = get_column_letter(col)
    qcs = [get_column_letter(q_col(y, q)) for q in (1, 2, 3, 4)]
    prev_cl = get_column_letter(ann_col(y - 1)) if y > 2016 else None
    q4c = get_column_letter(q_col(y, 4))
    if y <= 2025:
        b = QBS[(y, 4)]
        a = ANN[y]
        put(R["cash"], col, b["cash"], True, NF_N, True, fill)
        put(R["ar_b"], col, b["ar"], True, NF_N, fill=fill)
        put(R["inv_b"], col, b["inv"], True, NF_N, fill=fill)
        oca = b["ca"] - b["cash"] - b["ar"] - b["inv"]
        put(R["oca"], col, oca, True, NF_N, fill=fill)
        formula(R["tca"], col, f'={cl}{R["cash"]}+{cl}{R["ar_b"]}+{cl}{R["inv_b"]}+{cl}{R["oca"]}', NF_N, True, FILL_TOT)
        put(R["ppe_b"], col, b["ppe"], True, NF_N, fill=fill)
        put(R["gw_b"], col, b["gw"], True, NF_N, fill=fill)
        put(R["ia_b"], col, b["ia"], True, NF_N, fill=fill)
        onca = b["assets"] - b["ca"] - b["ppe"] - b["gw"] - b["ia"]
        put(R["onca"], col, onca, True, NF_N, fill=fill)
        formula(R["tassets"], col, f'={cl}{R["tca"]}+{cl}{R["ppe_b"]}+{cl}{R["gw_b"]}+{cl}{R["ia_b"]}+{cl}{R["onca"]}', NF_N, True, FILL_TOT)
        put(R["ap_b"], col, b["ap"], True, NF_N, fill=fill)
        put(R["stdebt"], col, b.get("std", 0), True, NF_N, fill=fill)
        ocl = b["cl"] - b["ap"] - b.get("std", 0)
        put(R["ocl"], col, ocl, True, NF_N, fill=fill)
        formula(R["tcl"], col, f'={cl}{R["ap_b"]}+{cl}{R["stdebt"]}+{cl}{R["ocl"]}', NF_N, fill=fill)
        put(R["ltdebt"], col, b["ltd"], True, NF_N, fill=fill)
        oncl = b["liab"] - b["cl"] - b["ltd"]
        put(R["oncl"], col, oncl, True, NF_N, fill=fill)
        formula(R["tliab"], col, f'={cl}{R["tcl"]}+{cl}{R["ltdebt"]}+{cl}{R["oncl"]}', NF_N, True, FILL_TOT)
        put(R["re_b"], col, b["re"], True, NF_N, fill=fill)
        oeq = b["eq"] - b["re"]
        put(R["oeq"], col, oeq, True, NF_N, fill=fill)
        formula(R["teq"], col, f'={cl}{R["oeq"]}+{cl}{R["re_b"]}', NF_N, True, FILL_TOT)
        formula(R["tle"], col, f'={cl}{R["tliab"]}+{cl}{R["teq"]}', NF_N, True, FILL_TOT)
        formula(R["cf_ni"], col, f'={cl}{R["ni"]}', NF_N, fill=fill)
        formula(R["cf_da"], col, f'={cl}{R["da_ppe"]}+{cl}{R["da_ia"]}', NF_N, fill=fill)
        formula(R["cf_sbc"], col, f'={cl}{R["sbc_d"]}', NF_N, fill=fill)
        if prev_cl:
            formula(R["cf_ar"], col, f'=-({cl}{R["ar_b"]}-{prev_cl}{R["ar_b"]})', NF_N, fill=fill)
            formula(R["cf_inv"], col, f'=-({cl}{R["inv_b"]}-{prev_cl}{R["inv_b"]})', NF_N, fill=fill)
            formula(R["cf_oca"], col, f'=-({cl}{R["oca"]}-{prev_cl}{R["oca"]})', NF_N, fill=fill)
            formula(R["cf_ap"], col, f'={cl}{R["ap_b"]}-{prev_cl}{R["ap_b"]}', NF_N, fill=fill)
            formula(R["cf_ocl"], col, f'=({cl}{R["ocl"]}-{prev_cl}{R["ocl"]})+({cl}{R["oncl"]}-{prev_cl}{R["oncl"]})', NF_N, fill=fill)
            formula(R["cash_o"], col, f'={prev_cl}{R["cash"]}', NF_N, fill=fill)
        else:
            put(R["cf_ar"], col, 0, True, NF_N, fill=fill)
            put(R["cf_inv"], col, 0, True, NF_N, fill=fill)
            put(R["cf_oca"], col, 0, True, NF_N, fill=fill)
            put(R["cf_ap"], col, 0, True, NF_N, fill=fill)
            put(R["cf_ocl"], col, 0, True, NF_N, fill=fill)
            put(R["cash_o"], col, b["cash"] - (a["cfo"] + a["cfi"] + a["cff"] + a["fx"]), True, NF_N, fill=fill)
        put(R["cfo"], col, a["cfo"], True, NF_N, True, fill)
        put(R["cf_capex"], col, -a["capex"], True, NF_N, fill=fill)
        put(R["cf_onca"], col, a["cfi"] + a["capex"], True, NF_N, fill=fill)
        formula(R["cfi"], col, f'={cl}{R["cf_capex"]}+{cl}{R["cf_onca"]}', NF_N, True, FILL_TOT)
        put(R["cf_debt"], col, 0, True, NF_N, fill=fill)
        put(R["cf_eq"], col, a["cff"], True, NF_N, fill=fill)
        formula(R["cff"], col, f'={cl}{R["cf_debt"]}+{cl}{R["cf_eq"]}', NF_N, True, FILL_TOT)
        put(R["cf_fx"], col, a["fx"], True, NF_N, fill=fill)
        formula(R["d_cash"], col, f'={cl}{R["cfo"]}+{cl}{R["cfi"]}+{cl}{R["cff"]}+{cl}{R["cf_fx"]}', NF_N, True, FILL_TOT)
        formula(R["cash_c"], col, f'={cl}{R["cash"]}', NF_N, True, FILL_TOT)
    else:
        for k in ["cash", "ar_b", "inv_b", "oca", "tca", "ppe_b", "gw_b", "ia_b", "onca", "tassets",
                  "ap_b", "stdebt", "ocl", "tcl", "ltdebt", "oncl", "tliab", "oeq", "re_b", "teq", "tle"]:
            formula(R[k], col, f'={q4c}{R[k]}', NF_N, k in ("tassets", "tliab", "tle", "teq"),
                    FILL_TOT if k in ("tassets", "tle") else fill)

        def qsum(k):
            return "=" + "+".join(f"{qc}{R[k]}" for qc in qcs)

        for k in ["cf_ni", "cf_da", "cf_sbc", "cf_ar", "cf_inv", "cf_oca", "cf_ap", "cf_ocl",
                  "cfo", "cf_capex", "cf_onca", "cfi", "cf_debt", "cf_eq", "cff", "cf_fx", "d_cash"]:
            formula(R[k], col, qsum(k), NF_N, k in ("cfo", "cfi", "cff", "d_cash"), FILL_TOT if k in ("cfo", "cfi", "cff", "d_cash") else fill)
        formula(R["cash_o"], col, f'={qcs[0]}{R["cash_o"]}', NF_N, fill=fill)
        formula(R["cash_c"], col, f'={q4c}{R["cash_c"]}', NF_N, True, FILL_TOT)

    formula(R["ck_bs"], col, f'={cl}{R["tassets"]}-{cl}{R["tle"]}', NF_N, True, FILL_YELLOW)
    formula(R["ck_cash"], col, f'={cl}{R["cash"]}-{cl}{R["cash_c"]}', NF_N, True, FILL_YELLOW)
    if y <= 2025:
        if prev_cl:
            formula(R["ck_re"], col, f'={cl}{R["re_b"]}-({prev_cl}{R["re_b"]}+{cl}{R["ni"]})', NF_N, fill=FILL_YELLOW)
        else:
            put(R["ck_re"], col, 0, False, NF_N, fill=FILL_YELLOW)
    else:
        formula(R["ck_re"], col, f'={cl}{R["re_b"]}-({prev_cl}{R["re_b"]}+{cl}{R["ni"]})', NF_N, fill=FILL_YELLOW)
    formula(R["ck_rev"], col, f'={cl}{R["rev"]}-{cl}{R["tot_drv"]}', NF_N, True, FILL_YELLOW)
    formula(R["ck_ni"], col, f'={cl}{R["ni"]}-({cl}{R["ebt"]}-{cl}{R["tax"]})', NF_N, fill=FILL_YELLOW)
    formula(R["ck_cfid"], col, f'={cl}{R["cfo"]}+{cl}{R["cfi"]}+{cl}{R["cff"]}+{cl}{R["cf_fx"]}-{cl}{R["d_cash"]}', NF_N, True, FILL_YELLOW)
    formula(R["ck_annq"], col, f'={cl}{R["rev"]}-({qcs[0]}{R["rev"]}+{qcs[1]}{R["rev"]}+{qcs[2]}{R["rev"]}+{qcs[3]}{R["rev"]})', NF_N, True, FILL_YELLOW)

# ---------- Quarterly P&L / BS / CF ----------
for (y, q) in QUARTERS:
    col = q_col(y, q)
    fill = period_fill(yq=(y, q))
    cl = get_column_letter(col)
    days = f"{cl}$11"
    py, pq = q_prev(y, q)
    pcl = None if (y == 2016 and q == 1) else get_column_letter(q_col(py, pq))

    if is_actual_q(y, q):
        p = QPL[(y, q)]
        b = QBS[(y, q)]
        put(R["rev"], col, p["rev"], True, NF_N, True, fill)
        put(R["cogs"], col, p["cogs"], True, NF_N, fill=fill)
        formula(R["gp"], col, f'={cl}{R["rev"]}-{cl}{R["cogs"]}', NF_N, True, FILL_TOT)
        formula(R["gpm"], col, f'=IF({cl}{R["rev"]}=0,0,{cl}{R["gp"]}/{cl}{R["rev"]})', NF_P, fill=fill)
        put(R["rd"], col, p["rd"], True, NF_N, fill=fill)
        put(R["sga"], col, p["sga"], True, NF_N, fill=fill)
        # other opex plug so OI matches reported
        oth_op = p["gp"] - p["rd"] - p["sga"] - p["oi"]
        put(R["rs"], col, oth_op, True, NF_N, fill=fill)
        formula(R["opex"], col, f'={cl}{R["rd"]}+{cl}{R["sga"]}+{cl}{R["rs"]}', NF_N, fill=fill)
        formula(R["oi"], col, f'={cl}{R["gp"]}-{cl}{R["opex"]}', NF_N, True, FILL_TOT)
        formula(R["om"], col, f'=IF({cl}{R["rev"]}=0,0,{cl}{R["oi"]}/{cl}{R["rev"]})', NF_P, fill=fill)
        put(R["ii"], col, p.get("ii", 0), True, NF_N, fill=fill)
        put(R["ie"], col, p.get("ie", 0), True, NF_N, fill=fill)
        formula(R["intn"], col, f'={cl}{R["ii"]}-{cl}{R["ie"]}', NF_N, fill=fill)
        put(R["ebt"], col, p["ebt"], True, NF_N, fill=fill)
        formula(R["oth"], col, f'={cl}{R["ebt"]}-{cl}{R["oi"]}-{cl}{R["intn"]}', NF_N, fill=fill)
        put(R["tax"], col, p["tax"], True, NF_N, fill=fill)
        formula(R["ni"], col, f'={cl}{R["ebt"]}-{cl}{R["tax"]}', NF_N, True, FILL_TOT)
        put(R["shd"], col, p["shd"], True, NF_N, fill=fill)
        formula(R["eps"], col, f'=IF({cl}{R["shd"]}=0,0,{cl}{R["ni"]}/{cl}{R["shd"]})', NF_2, fill=fill)
        formula(R["gm_co"], col, f'={cl}{R["gpm"]}', NF_P, fill=fill)
        formula(R["other_gm"], col, f'=IF({cl}{R["rev"]}-{cl}{R["opt_rev"]}=0,0,({cl}{R["gp"]}-({cl}{R["dsp_rev"]}-{cl}{R["dsp_cogs"]}+{cl}{R["tia_rev"]}-{cl}{R["tia_cogs"]}+{cl}{R["drv_rev"]}-{cl}{R["drv_cogs"]}))/MAX({cl}{R["rev"]}-{cl}{R["opt_rev"]},1))', NF_P, fill=fill)
        formula(R["sga_d"], col, f'={cl}{R["sga"]}', NF_N, fill=fill)
        formula(R["rd_d"], col, f'={cl}{R["rd"]}', NF_N, fill=fill)
        formula(R["rs_d"], col, f'={cl}{R["rs"]}', NF_N, fill=fill)
        formula(R["tax_r"], col, f'=IF({cl}{R["ebt"]}=0,0,{cl}{R["tax"]}/{cl}{R["ebt"]})', NF_P, fill=fill)
        put(R["tax_min"], col, 1500, True, NF_N, fill=fill)
        put(R["yld"], col, 0.008, True, NF_P, fill=fill)
        put(R["drate"], col, 0.018, True, NF_P, fill=fill)
        formula(R["oth_d"], col, f'={cl}{R["oth"]}', NF_N, fill=fill)
        dq = 91 if q != 4 else 92
        put(R["dso"], col, round(b["ar"] / p["rev"] * dq, 1) if p["rev"] else 0, True, NF_1, fill=fill)
        put(R["dio"], col, round(b["inv"] / p["cogs"] * dq, 1) if p["cogs"] else 0, True, NF_1, fill=fill)
        put(R["dpo"], col, round(b["ap"] / p["cogs"] * dq, 1) if p["cogs"] else 0, True, NF_1, fill=fill)
        if y <= 2025:
            put(R["capex_d"], col, round(ANN[y]["capex"] / 4), True, NF_N, fill=fill)
            put(R["da_ppe"], col, round(ANN[y]["da"] * 0.35 / 4), True, NF_N, fill=fill)
            put(R["da_ia"], col, round(ANN[y]["da"] * 0.65 / 4), True, NF_N, fill=fill)
            put(R["sbc_d"], col, round(ANN[y]["sbc"] / 4), True, NF_N, fill=fill)
        elif q == 1:
            put(R["capex_d"], col, 1384, True, NF_N, fill=fill)
            put(R["da_ppe"], col, 3500, True, NF_N, fill=fill)
            put(R["da_ia"], col, 7439, True, NF_N, fill=fill)
            put(R["sbc_d"], col, 20027, True, NF_N, fill=fill)
        else:
            put(R["capex_d"], col, 2338, True, NF_N, fill=fill)
            put(R["da_ppe"], col, 3500, True, NF_N, fill=fill)
            put(R["da_ia"], col, 7000, True, NF_N, fill=fill)
            put(R["sbc_d"], col, 21000, True, NF_N, fill=fill)
        put(R["sh_b"], col, p["shd"], True, NF_N, fill=fill)

        oca = b["ca"] - b["cash"] - b["ar"] - b["inv"]
        onca = b["assets"] - b["ca"] - b["ppe"] - b["gw"] - b["ia"]
        put(R["cash"], col, b["cash"], True, NF_N, True, fill)
        put(R["ar_b"], col, b["ar"], True, NF_N, fill=fill)
        put(R["inv_b"], col, b["inv"], True, NF_N, fill=fill)
        put(R["oca"], col, oca, True, NF_N, fill=fill)
        formula(R["tca"], col, f'={cl}{R["cash"]}+{cl}{R["ar_b"]}+{cl}{R["inv_b"]}+{cl}{R["oca"]}', NF_N, True, FILL_TOT)
        put(R["ppe_b"], col, b["ppe"], True, NF_N, fill=fill)
        put(R["gw_b"], col, b["gw"], True, NF_N, fill=fill)
        put(R["ia_b"], col, b["ia"], True, NF_N, fill=fill)
        put(R["onca"], col, onca, True, NF_N, fill=fill)
        formula(R["tassets"], col, f'={cl}{R["tca"]}+{cl}{R["ppe_b"]}+{cl}{R["gw_b"]}+{cl}{R["ia_b"]}+{cl}{R["onca"]}', NF_N, True, FILL_TOT)
        put(R["ap_b"], col, b["ap"], True, NF_N, fill=fill)
        put(R["stdebt"], col, b.get("std", 0), True, NF_N, fill=fill)
        ocl = b["cl"] - b["ap"] - b.get("std", 0)
        put(R["ocl"], col, ocl, True, NF_N, fill=fill)
        formula(R["tcl"], col, f'={cl}{R["ap_b"]}+{cl}{R["stdebt"]}+{cl}{R["ocl"]}', NF_N, fill=fill)
        put(R["ltdebt"], col, b["ltd"], True, NF_N, fill=fill)
        oncl = b["liab"] - b["cl"] - b["ltd"]
        put(R["oncl"], col, oncl, True, NF_N, fill=fill)
        formula(R["tliab"], col, f'={cl}{R["tcl"]}+{cl}{R["ltdebt"]}+{cl}{R["oncl"]}', NF_N, True, FILL_TOT)
        put(R["re_b"], col, b["re"], True, NF_N, fill=fill)
        put(R["oeq"], col, b["eq"] - b["re"], True, NF_N, fill=fill)
        formula(R["teq"], col, f'={cl}{R["oeq"]}+{cl}{R["re_b"]}', NF_N, True, FILL_TOT)
        formula(R["tle"], col, f'={cl}{R["tliab"]}+{cl}{R["teq"]}', NF_N, True, FILL_TOT)

        formula(R["cf_ni"], col, f'={cl}{R["ni"]}', NF_N, fill=fill)
        formula(R["cf_da"], col, f'={cl}{R["da_ppe"]}+{cl}{R["da_ia"]}', NF_N, fill=fill)
        formula(R["cf_sbc"], col, f'={cl}{R["sbc_d"]}', NF_N, fill=fill)
        if pcl:
            formula(R["cf_ar"], col, f'=-({cl}{R["ar_b"]}-{pcl}{R["ar_b"]})', NF_N, fill=fill)
            formula(R["cf_inv"], col, f'=-({cl}{R["inv_b"]}-{pcl}{R["inv_b"]})', NF_N, fill=fill)
            formula(R["cf_oca"], col, f'=-({cl}{R["oca"]}-{pcl}{R["oca"]})', NF_N, fill=fill)
            formula(R["cf_ap"], col, f'={cl}{R["ap_b"]}-{pcl}{R["ap_b"]}', NF_N, fill=fill)
            formula(R["cf_ocl"], col, f'=({cl}{R["ocl"]}-{pcl}{R["ocl"]})+({cl}{R["oncl"]}-{pcl}{R["oncl"]})', NF_N, fill=fill)
            formula(R["cash_o"], col, f'={pcl}{R["cash"]}', NF_N, fill=fill)
            formula(R["cf_onca"], col, f'=-(({cl}{R["gw_b"]}+{cl}{R["ia_b"]}+{cl}{R["onca"]})-({pcl}{R["gw_b"]}+{pcl}{R["ia_b"]}+{pcl}{R["onca"]}))', NF_N, fill=fill)
            formula(R["cf_debt"], col, f'=({cl}{R["stdebt"]}-{pcl}{R["stdebt"]})+({cl}{R["ltdebt"]}-{pcl}{R["ltdebt"]})', NF_N, fill=fill)
            formula(R["cf_eq"], col, f'={cl}{R["oeq"]}-{pcl}{R["oeq"]}-{cl}{R["cf_sbc"]}', NF_N, fill=fill)
        else:
            for k, v in [("cf_ar", 0), ("cf_inv", 0), ("cf_oca", 0), ("cf_ap", 0), ("cf_ocl", 0),
                         ("cash_o", 68000), ("cf_onca", 0), ("cf_debt", 0), ("cf_eq", 0)]:
                put(R[k], col, v, True, NF_N, fill=fill)
        formula(R["cfo"], col, f'={cl}{R["cf_ni"]}+{cl}{R["cf_da"]}+{cl}{R["cf_sbc"]}+{cl}{R["cf_ar"]}+{cl}{R["cf_inv"]}+{cl}{R["cf_oca"]}+{cl}{R["cf_ap"]}+{cl}{R["cf_ocl"]}', NF_N, True, FILL_TOT)
        formula(R["cf_capex"], col, f'=-{cl}{R["capex_d"]}', NF_N, fill=fill)
        formula(R["cfi"], col, f'={cl}{R["cf_capex"]}+{cl}{R["cf_onca"]}', NF_N, True, FILL_TOT)
        formula(R["cff"], col, f'={cl}{R["cf_debt"]}+{cl}{R["cf_eq"]}', NF_N, True, FILL_TOT)
        formula(R["cf_fx"], col, f'={cl}{R["cash"]}-({cl}{R["cash_o"]}+{cl}{R["cfo"]}+{cl}{R["cfi"]}+{cl}{R["cff"]})', NF_N, fill=fill)
        formula(R["d_cash"], col, f'={cl}{R["cfo"]}+{cl}{R["cfi"]}+{cl}{R["cff"]}+{cl}{R["cf_fx"]}', NF_N, True, FILL_TOT)
        formula(R["cash_c"], col, f'={cl}{R["cash_o"]}+{cl}{R["d_cash"]}', NF_N, True, FILL_TOT)
    else:
        f = FCST[(y, q)]
        put(R["other_gm"], col, f["other_gm"], True, NF_P, fill=fill)
        put(R["sga_d"], col, f["sga"], True, NF_N, fill=fill)
        put(R["rd_d"], col, f["rd"], True, NF_N, fill=fill)
        put(R["rs_d"], col, f["rs"], True, NF_N, fill=fill)
        put(R["tax_r"], col, f["tax_rate"], True, NF_P, fill=fill)
        put(R["tax_min"], col, f["tax_min"], True, NF_N, fill=fill)
        put(R["yld"], col, f["yld"], True, NF_P, fill=fill)
        put(R["drate"], col, f["drate"], True, NF_P, fill=fill)
        put(R["oth_d"], col, f["oth"], True, NF_N, fill=fill)
        put(R["dso"], col, f["dso"], True, NF_1, fill=fill)
        put(R["dio"], col, f["dio"], True, NF_1, fill=fill)
        put(R["dpo"], col, f["dpo"], True, NF_1, fill=fill)
        put(R["capex_d"], col, f["capex"], True, NF_N, fill=fill)
        put(R["da_ppe"], col, f["da_ppe"], True, NF_N, fill=fill)
        put(R["da_ia"], col, f["da_ia"], True, NF_N, fill=fill)
        put(R["sbc_d"], col, f["sbc"], True, NF_N, fill=fill)
        put(R["sh_b"], col, f["sh_b"], True, NF_N, fill=fill)
        put(R["stdebt"], col, f["stdebt"], True, NF_N, fill=fill)
        put(R["ltdebt"], col, f["ltdebt"], True, NF_N, fill=fill)

        formula(R["rev"], col, f'={cl}{R["tot_drv"]}', NF_N, True, FILL_TOT)
        formula(R["cogs"], col, f'={cl}{R["dsp_cogs"]}+{cl}{R["tia_cogs"]}+{cl}{R["drv_cogs"]}+({cl}{R["oth_inf"]}+{cl}{R["bb"]}+{cl}{R["conn"]}+{cl}{R["ind"]})*(1-{cl}{R["other_gm"]})', NF_N, fill=fill)
        formula(R["gp"], col, f'={cl}{R["rev"]}-{cl}{R["cogs"]}', NF_N, True, FILL_TOT)
        formula(R["gpm"], col, f'=IF({cl}{R["rev"]}=0,0,{cl}{R["gp"]}/{cl}{R["rev"]})', NF_P, fill=fill)
        formula(R["gm_co"], col, f'={cl}{R["gpm"]}', NF_P, fill=fill)
        formula(R["sga"], col, f'={cl}{R["sga_d"]}', NF_N, fill=fill)
        formula(R["rd"], col, f'={cl}{R["rd_d"]}', NF_N, fill=fill)
        formula(R["rs"], col, f'={cl}{R["rs_d"]}', NF_N, fill=fill)
        formula(R["opex"], col, f'={cl}{R["rd"]}+{cl}{R["sga"]}+{cl}{R["rs"]}', NF_N, fill=fill)
        formula(R["oi"], col, f'={cl}{R["gp"]}-{cl}{R["opex"]}', NF_N, True, FILL_TOT)
        formula(R["om"], col, f'=IF({cl}{R["rev"]}=0,0,{cl}{R["oi"]}/{cl}{R["rev"]})', NF_P, fill=fill)
        formula(R["ii"], col, f'={pcl}{R["cash"]}*{cl}{R["yld"]}', NF_N, fill=fill)
        formula(R["ie"], col, f'=({pcl}{R["stdebt"]}+{pcl}{R["ltdebt"]})*{cl}{R["drate"]}', NF_N, fill=fill)
        formula(R["intn"], col, f'={cl}{R["ii"]}-{cl}{R["ie"]}', NF_N, fill=fill)
        formula(R["oth"], col, f'={cl}{R["oth_d"]}', NF_N, fill=fill)
        formula(R["ebt"], col, f'={cl}{R["oi"]}+{cl}{R["intn"]}+{cl}{R["oth"]}', NF_N, fill=fill)
        formula(R["tax"], col, f'=IF({cl}{R["ebt"]}>0,{cl}{R["ebt"]}*{cl}{R["tax_r"]},{cl}{R["tax_min"]})', NF_N, fill=fill)
        formula(R["ni"], col, f'={cl}{R["ebt"]}-{cl}{R["tax"]}', NF_N, True, FILL_TOT)
        put(R["shd"], col, f["sh_d"], True, NF_N, fill=fill)
        formula(R["eps"], col, f'=IF({cl}{R["shd"]}=0,0,{cl}{R["ni"]}/{cl}{R["shd"]})', NF_2, fill=fill)

        formula(R["ar_b"], col, f'={cl}{R["dso"]}/{days}*{cl}{R["rev"]}', NF_N, fill=fill)
        formula(R["inv_b"], col, f'={cl}{R["dio"]}/{days}*{cl}{R["cogs"]}', NF_N, fill=fill)
        formula(R["oca"], col, f'={pcl}{R["oca"]}', NF_N, fill=fill)
        formula(R["ppe_b"], col, f'={pcl}{R["ppe_b"]}+{cl}{R["capex_d"]}-{cl}{R["da_ppe"]}', NF_N, fill=fill)
        formula(R["gw_b"], col, f'={pcl}{R["gw_b"]}', NF_N, fill=fill)
        formula(R["ia_b"], col, f'=MAX(20000,{pcl}{R["ia_b"]}-{cl}{R["da_ia"]})', NF_N, fill=fill)
        formula(R["onca"], col, f'={pcl}{R["onca"]}', NF_N, fill=fill)
        formula(R["ap_b"], col, f'={cl}{R["dpo"]}/{days}*{cl}{R["cogs"]}', NF_N, fill=fill)
        formula(R["ocl"], col, f'={pcl}{R["ocl"]}', NF_N, fill=fill)
        formula(R["tcl"], col, f'={cl}{R["ap_b"]}+{cl}{R["stdebt"]}+{cl}{R["ocl"]}', NF_N, fill=fill)
        formula(R["oncl"], col, f'={pcl}{R["oncl"]}', NF_N, fill=fill)
        formula(R["tliab"], col, f'={cl}{R["tcl"]}+{cl}{R["ltdebt"]}+{cl}{R["oncl"]}', NF_N, True, FILL_TOT)
        formula(R["oeq"], col, f'={pcl}{R["oeq"]}+{cl}{R["sbc_d"]}', NF_N, fill=fill)
        formula(R["re_b"], col, f'={pcl}{R["re_b"]}+{cl}{R["ni"]}', NF_N, fill=fill)
        formula(R["teq"], col, f'={cl}{R["oeq"]}+{cl}{R["re_b"]}', NF_N, True, FILL_TOT)

        formula(R["cf_ni"], col, f'={cl}{R["ni"]}', NF_N, fill=fill)
        # Add back only PPE D&A plus actual IA rundown (floor may bind)
        formula(R["cf_da"], col, f'={cl}{R["da_ppe"]}+({pcl}{R["ia_b"]}-{cl}{R["ia_b"]})', NF_N, fill=fill)
        formula(R["cf_sbc"], col, f'={cl}{R["sbc_d"]}', NF_N, fill=fill)
        formula(R["cf_ar"], col, f'=-({cl}{R["ar_b"]}-{pcl}{R["ar_b"]})', NF_N, fill=fill)
        formula(R["cf_inv"], col, f'=-({cl}{R["inv_b"]}-{pcl}{R["inv_b"]})', NF_N, fill=fill)
        formula(R["cf_oca"], col, f'=-({cl}{R["oca"]}-{pcl}{R["oca"]})', NF_N, fill=fill)
        formula(R["cf_ap"], col, f'={cl}{R["ap_b"]}-{pcl}{R["ap_b"]}', NF_N, fill=fill)
        formula(R["cf_ocl"], col, f'=({cl}{R["ocl"]}-{pcl}{R["ocl"]})+({cl}{R["oncl"]}-{pcl}{R["oncl"]})', NF_N, fill=fill)
        formula(R["cfo"], col, f'={cl}{R["cf_ni"]}+{cl}{R["cf_da"]}+{cl}{R["cf_sbc"]}+{cl}{R["cf_ar"]}+{cl}{R["cf_inv"]}+{cl}{R["cf_oca"]}+{cl}{R["cf_ap"]}+{cl}{R["cf_ocl"]}', NF_N, True, FILL_TOT)
        formula(R["cf_capex"], col, f'=-{cl}{R["capex_d"]}', NF_N, fill=fill)
        # IA amortization is non-cash (in D&A); do not put ΔIA in CFI. GW/ONCA held constant.
        formula(R["cf_onca"], col, f'=-({cl}{R["onca"]}-{pcl}{R["onca"]})', NF_N, fill=fill)
        formula(R["cfi"], col, f'={cl}{R["cf_capex"]}+{cl}{R["cf_onca"]}', NF_N, True, FILL_TOT)
        formula(R["cf_debt"], col, f'=({cl}{R["stdebt"]}-{pcl}{R["stdebt"]})+({cl}{R["ltdebt"]}-{pcl}{R["ltdebt"]})', NF_N, fill=fill)
        put(R["cf_eq"], col, 0, True, NF_N, fill=fill)
        formula(R["cff"], col, f'={cl}{R["cf_debt"]}+{cl}{R["cf_eq"]}', NF_N, True, FILL_TOT)
        put(R["cf_fx"], col, f["fx"], True, NF_N, fill=fill)
        formula(R["d_cash"], col, f'={cl}{R["cfo"]}+{cl}{R["cfi"]}+{cl}{R["cff"]}+{cl}{R["cf_fx"]}', NF_N, True, FILL_TOT)
        formula(R["cash_o"], col, f'={pcl}{R["cash"]}', NF_N, fill=fill)
        formula(R["cash_c"], col, f'={cl}{R["cash_o"]}+{cl}{R["d_cash"]}', NF_N, True, FILL_TOT)
        formula(R["cash"], col, f'={cl}{R["cash_c"]}', NF_N, True, FILL_TOT)
        formula(R["tca"], col, f'={cl}{R["cash"]}+{cl}{R["ar_b"]}+{cl}{R["inv_b"]}+{cl}{R["oca"]}', NF_N, True, FILL_TOT)
        formula(R["tassets"], col, f'={cl}{R["tca"]}+{cl}{R["ppe_b"]}+{cl}{R["gw_b"]}+{cl}{R["ia_b"]}+{cl}{R["onca"]}', NF_N, True, FILL_TOT)
        formula(R["tle"], col, f'={cl}{R["tliab"]}+{cl}{R["teq"]}', NF_N, True, FILL_TOT)

    formula(R["ck_bs"], col, f'={cl}{R["tassets"]}-{cl}{R["tle"]}', NF_N, True, FILL_YELLOW)
    formula(R["ck_cash"], col, f'={cl}{R["cash"]}-{cl}{R["cash_c"]}', NF_N, True, FILL_YELLOW)
    if pcl:
        formula(R["ck_re"], col, f'={cl}{R["re_b"]}-({pcl}{R["re_b"]}+{cl}{R["ni"]})', NF_N, fill=FILL_YELLOW)
    else:
        put(R["ck_re"], col, 0, False, NF_N, fill=FILL_YELLOW)
    formula(R["ck_rev"], col, f'={cl}{R["rev"]}-{cl}{R["tot_drv"]}', NF_N, True, FILL_YELLOW)
    formula(R["ck_ni"], col, f'={cl}{R["ni"]}-({cl}{R["ebt"]}-{cl}{R["tax"]})', NF_N, fill=FILL_YELLOW)
    formula(R["ck_cfid"], col, f'={cl}{R["cfo"]}+{cl}{R["cfi"]}+{cl}{R["cff"]}+{cl}{R["cf_fx"]}-{cl}{R["d_cash"]}', NF_N, True, FILL_YELLOW)
    put(R["ck_annq"], col, 0, False, NF_N, fill=FILL_YELLOW)

red_fill = PatternFill("solid", fgColor="FFC7CE")
green_fill = PatternFill("solid", fgColor="C6EFCE")
for rkey in ["ck_bs", "ck_cash", "ck_rev", "ck_cfid"]:
    rng = f"{get_column_letter(ANN_COL0)}{R[rkey]}:{get_column_letter(NOTES_COL-1)}{R[rkey]}"
    ws.conditional_formatting.add(rng, CellIsRule(operator="notBetween", formula=["-2", "2"], fill=red_fill))
    ws.conditional_formatting.add(rng, CellIsRule(operator="between", formula=["-2", "2"], fill=green_fill))

ws.row_dimensions[1].height = 24
ws.sheet_view.zoomScale = 85
ws.oddFooter.left.text = "MaxLinear (MXL) | Hedge-fund style operating model | Not investment advice"
ws.oddFooter.right.text = "Blue = input / Black = formula  |  Confidential"

# Sanity: annual driver vs 10-K rev for history
for y in range(2016, 2026):
    seg = SEG_ANN[y]
    s = seg["bb"] + seg["conn"] + seg["inf"] + seg["ind"]
    if abs(s - ANN[y]["rev"]) > 2:
        raise SystemExit(f"Segment sum {y}: {s} vs {ANN[y]['rev']}")

out = "/workspace/MXL_Financial_Model.xlsx"
wb.save(out)
print("Saved", out)
print("Rows mapped:", len(R), "last data row", R["ck_annq"])
print("Annual cols", ANN_COL0, "to", ANN_COL0 + len(ANN_YEARS) - 1)
print("Quarter cols", Q_COL0, "to", Q_COL0 + len(QUARTERS) - 1, "notes", NOTES_COL)
print("Q3'26 driver rev check",
      FCST[(2026, 3)]["dsp_vol"] * FCST[(2026, 3)]["dsp_asp"]
      + FCST[(2026, 3)]["tia_vol"] * FCST[(2026, 3)]["tia_asp"]
      + FCST[(2026, 3)]["drv_vol"] * FCST[(2026, 3)]["drv_asp"]
      + FCST[(2026, 3)]["oth_inf"] + FCST[(2026, 3)]["bb"] + FCST[(2026, 3)]["conn"] + FCST[(2026, 3)]["ind"])
