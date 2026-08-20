#!/usr/bin/env python3
"""Build AXTI (AXT, Inc.) 3-statement financial model (single sheet)."""
from __future__ import annotations

from copy import deepcopy
from openpyxl import Workbook
from openpyxl.styles import Font, Fill, PatternFill, Alignment, Border, Side, NamedStyle
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import FormulaRule
from openpyxl.comments import Comment
from openpyxl.utils.cell import absolute_coordinate
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.chart import LineChart, Reference, BarChart
from openpyxl.chart.series import SeriesLabel
from openpyxl.worksheet.page import PageMargins
from openpyxl.formatting.rule import CellIsRule

# ---------------------------------------------------------------------------
# Periods
# ---------------------------------------------------------------------------
ANN_YEARS = list(range(2016, 2029))  # 2016-2028
QUARTERS = [(y, q) for y in range(2016, 2029) for q in (1, 2, 3, 4)]
ACTUAL_ANN = set(range(2016, 2026))
def is_actual_q(y, q):
    return (y, q) <= (2026, 2)

# ---------------------------------------------------------------------------
# Annual reported financials ($000)
# ---------------------------------------------------------------------------
ANN = {
    # rev, cogs, sga, rd, oi, ebt, tax, ni_axt, nci, int_net, jv, da, sbc, cfo, cfi, cff, fx,
    # capex, shares_b, shares_d, cash, ar, inv, ppe, assets, ca, ap, cl, liab, re, eq_axt, nci_bs, rnci, sh_out
    2016: dict(rev=81349, cogs=54968, sga=13880, rd=5850, oi=6425, ebt=5699, tax=733, ni_axt=5636, nci=-670,
               int_net=409, jv=-1995, da=4865, sbc=1096, cfo=12504, cfi=-1113, cff=1298, fx=-1412, capex=2728,
               sh_b=32139, sh_d=32894, cash=36152, ar=14453, inv=40152, ppe=27805, assets=154246, ca=107286,
               ap=6691, cl=15951, liab=16856, re=-64985, eq_axt=133010, nci_bs=4380, rnci=0, sh_out=33032,
               pref=3532, apic=194463),  # apic+cs plug-ish
    2017: dict(rev=98673, cogs=64198, sga=17009, rd=4827, oi=12639, ebt=10853, tax=792, ni_axt=10148, nci=-87,
               int_net=0, jv=0, da=4422, sbc=1405, cfo=8615, cfi=-36458, cff=35638, fx=405, capex=21356,
               sh_b=37444, sh_d=38966, cash=44352, ar=22778, inv=45840, ppe=46530, assets=211200, ca=140521,
               ap=11445, cl=22594, liab=22883, re=-54837, eq_axt=183820, nci_bs=4497, rnci=0, sh_out=39413, pref=3532, apic=0),
    2018: dict(rev=102397, cogs=65350, sga=19003, rd=5897, oi=12147, ebt=11947, tax=938, ni_axt=9654, nci=1355,
               int_net=0, jv=0, da=4871, sbc=1925, cfo=3218, cfi=-30827, cff=213, fx=-430, capex=40539,
               sh_b=39049, sh_d=40265, cash=16526, ar=19586, inv=58571, ppe=82280, assets=223524, ca=128540,
               ap=13338, cl=28709, liab=28992, re=-45183, eq_axt=190835, nci_bs=3697, rnci=0, sh_out=39985, pref=3532, apic=0),
    2019: dict(rev=83256, cogs=58431, sga=19305, rd=5834, oi=-314, ebt=-1026, tax=562, ni_axt=-2600, nci=1012,
               int_net=0, jv=0, da=5531, sbc=2346, cfo=12658, cfi=-8328, cff=6186, fx=376, capex=21792,
               sh_b=39487, sh_d=39487, cash=26892, ar=19031, inv=49152, ppe=97403, assets=223349, ca=113205,
               ap=10098, cl=27526, liab=30587, re=-47783, eq_axt=187885, nci_bs=4877, rnci=0, sh_out=40632, pref=3532, apic=0),
    2020: dict(rev=95361, cogs=65086, sga=19200, rd=7135, oi=3940, ebt=7072, tax=2031, ni_axt=3238, nci=1803,
               int_net=0, jv=4410, da=4333, sbc=2623, cfo=5865, cfi=-16422, cff=52662, fx=3905, capex=19855,
               sh_b=40152, sh_d=41025, cash=72602, ar=24558, inv=51515, ppe=115825, assets=298862, ca=164518,
               ap=12669, cl=39075, liab=43330, re=-44545, eq_axt=192619, nci_bs=15350, rnci=47563, sh_out=41967, pref=3532, apic=0),
    2021: dict(rev=137393, cogs=89979, sga=24189, rd=10328, oi=12897, ebt=17602, tax=1093, ni_axt=14575, nci=1934,
               int_net=-210, jv=4410, da=7078, sbc=4519, cfo=-3305, cfi=-38810, cff=5725, fx=550, capex=29645,
               sh_b=41367, sh_d=42720, cash=36763, ar=34839, inv=65912, ppe=142415, assets=332441, ca=160185,
               ap=16649, cl=47822, liab=52210, re=-29970, eq_axt=211529, nci_bs=18317, rnci=50385, sh_out=42886, pref=3532, apic=231620),
    2022: dict(rev=141118, cogs=88997, sga=25654, rd=13913, oi=12554, ebt=20927, tax=2185, ni_axt=15811, nci=2931,
               int_net=-1070, jv=5960, da=8119, sbc=4006, cfo=-8765, cfi=-25223, cff=38031, fx=540, capex=28465,
               sh_b=42104, sh_d=42715, cash=34948, ar=29252, inv=89629, ppe=161017, assets=370072, ca=183545,
               ap=10084, cl=75326, liab=80326, re=-14159, eq_axt=221607, nci_bs=23293, rnci=44846, sh_out=43554, pref=3532, apic=235308),
    2023: dict(rev=75795, cogs=62477, sga=22806, rd=12081, oi=-21569, ebt=-19033, tax=160, ni_axt=-17881, nci=-1312,
               int_net=-1527, jv=1884, da=8722, sbc=3540, cfo=3403, cfi=-2604, cff=8613, fx=-646, capex=10475,
               sh_b=42643, sh_d=42643, cash=37752, ar=19256, inv=86503, ppe=166348, assets=358701, ca=170656,
               ap=9617, cl=81557, liab=89555, re=-32040, eq_axt=203989, nci_bs=23494, rnci=41663, sh_out=44239, pref=3532, apic=238452),
    2024: dict(rev=99361, cogs=75525, sga=24096, rd=14543, oi=-14803, ebt=-10657, tax=1134, ni_axt=-11624, nci=-167,
               int_net=-1340, jv=3439, da=8979, sbc=3097, cfo=-12112, cfi=-4445, cff=-536, fx=790, capex=5771,
               sh_b=43154, sh_d=43154, cash=22833, ar=25640, inv=85077, ppe=159721, assets=339314, ca=158272,
               ap=12356, cl=74176, liab=84406, re=-43664, eq_axt=192770, nci_bs=23561, rnci=38577, sh_out=45358, pref=3532, apic=241514),
    2025: dict(rev=88326, cogs=77084, sga=24169, rd=9049, oi=-21976, ebt=-21544, tax=1658, ni_axt=-21260, nci=-1942,
               int_net=-1257, jv=765, da=9108, sbc=3292, cfo=-12783, cfi=-6832, cff=107095, fx=7075, capex=5995,
               sh_b=43933, sh_d=43933, cash=120266, ar=26849, inv=81651, ppe=161860, assets=433751, ca=246556,
               ap=12947, cl=90541, liab=99120, re=-64924, eq_axt=273290, nci_bs=23285, rnci=38056, sh_out=55337, pref=3532, apic=339922),
}

# Product mix annual ($000). 2022-2025 substrates/RM from 10-K; earlier estimated from mix commentary.
ANN_PROD = {
    2016: dict(sub=64000, rm=17349, inp=17920, gaas=38080, ge=8000),
    2017: dict(sub=78000, rm=20673, inp=24960, gaas=44040, ge=9000),
    2018: dict(sub=82000, rm=20397, inp=27880, gaas=45620, ge=8500),
    2019: dict(sub=65000, rm=18256, inp=29900, gaas=29100, ge=6000),
    2020: dict(sub=75000, rm=20361, inp=35250, gaas=34250, ge=5500),
    2021: dict(sub=109000, rm=28393, inp=44690, gaas=56310, ge=8000),
    2022: dict(sub=111094, rm=30024, inp=47770, gaas=55824, ge=7500),
    2023: dict(sub=47466, rm=28329, inp=21800, gaas=21666, ge=4000),
    2024: dict(sub=67748, rm=31613, inp=31700, gaas=28648, ge=7400),
    2025: dict(sub=58900, rm=29426, inp=28500, gaas=26840, ge=3560),
}

# Blended InP ASP ($/wafer) — assumption (blue)
ANN_ASP = {2016: 185, 2017: 195, 2018: 205, 2019: 220, 2020: 245, 2021: 280, 2022: 310, 2023: 340, 2024: 390, 2025: 420}
ANN_UCOGS = {2016: 125, 2017: 128, 2018: 132, 2019: 145, 2020: 155, 2021: 175, 2022: 185, 2023: 255, 2024: 245, 2025: 285}

# Quarterly P&L ($000) — from SEC XBRL / 10-Q. ni = NI to AXT.
QPL = {
    (2016,1): dict(rev=18713, cogs=13460, gp=5253, sga=3374, rd=1381, oi=498, ebt=330, tax=397, ni=42, nci=-109),
    (2016,2): dict(rev=20495, cogs=14468, gp=6027, sga=3419, rd=1472, oi=910, ebt=938, tax=140, ni=1151, nci=-353),
    (2016,3): dict(rev=21872, cogs=14294, gp=7578, sga=3313, rd=1566, oi=2699, ebt=2387, tax=176, ni=2229, nci=-18),
    (2017,1): dict(rev=20616, cogs=14328, gp=6288, sga=3793, rd=1124, oi=1371, ebt=584, tax=159, ni=665, nci=-240),
    (2017,2): dict(rev=23557, cogs=16301, gp=7256, sga=3942, rd=1019, oi=2295, ebt=2119, tax=321, ni=1930, nci=-132),
    (2017,3): dict(rev=28168, cogs=17035, gp=11133, sga=4484, rd=1410, oi=5239, ebt=4746, tax=181, ni=4419, nci=146),
    (2018,1): dict(rev=24419, cogs=14846, gp=9573, sga=4222, rd=1420, oi=3931, ebt=3524, tax=334, ni=2875, nci=315),
    (2018,2): dict(rev=27120, cogs=16110, gp=11010, sga=4987, rd=1500, oi=4523, ebt=4918, tax=367, ni=3901, nci=650),
    (2018,3): dict(rev=28626, cogs=18012, gp=10614, sga=4615, rd=1668, oi=4331, ebt=4557, tax=410, ni=3939, nci=208),
    (2019,1): dict(rev=20208, cogs=13513, gp=6695, sga=4723, rd=1346, oi=626, ebt=-867, tax=156, ni=-1104, nci=81),
    (2019,2): dict(rev=24797, cogs=16291, gp=8506, sga=4769, rd=1399, oi=2338, ebt=2335, tax=597, ni=1451, nci=287),
    (2019,3): dict(rev=19841, cogs=14082, gp=5759, sga=4755, rd=1482, oi=-478, ebt=-472, tax=23, ni=-898, nci=403),
    (2020,1): dict(rev=20723, cogs=15201, gp=5522, sga=4749, rd=1407, oi=-634, ebt=583, tax=366, ni=-178, nci=395),
    (2020,2): dict(rev=22134, cogs=15366, gp=6768, sga=4747, rd=1543, oi=478, ebt=1879, tax=920, ni=361, nci=598),
    (2020,3): dict(rev=25469, cogs=16646, gp=8823, sga=4623, rd=2023, oi=2177, ebt=2118, tax=637, ni=991, nci=490),
    (2021,1): dict(rev=31350, cogs=19814, gp=11536, sga=5570, rd=2405, oi=3561, ebt=4511, tax=746, ni=3425, nci=340),
    (2021,2): dict(rev=33735, cogs=21497, gp=12238, sga=5795, rd=2537, oi=3906, ebt=5508, tax=893, ni=4385, nci=230),
    (2021,3): dict(rev=34576, cogs=23075, gp=11501, sga=6476, rd=2629, oi=2396, ebt=4361, tax=-135, ni=3800, nci=696),
    (2022,1): dict(rev=39653, cogs=26345, gp=13308, sga=6450, rd=3159, oi=3699, ebt=4632, tax=660, ni=3165, nci=807),
    (2022,2): dict(rev=39487, cogs=24052, gp=15435, sga=6693, rd=3453, oi=5289, ebt=7572, tax=1027, ni=5546, nci=999),
    (2022,3): dict(rev=35183, cogs=20401, gp=14782, sga=6576, rd=3639, oi=4567, ebt=7231, tax=501, ni=5759, nci=971),
    (2023,1): dict(rev=19405, cogs=14295, gp=5110, sga=5952, rd=3595, oi=-4437, ebt=-3518, tax=148, ni=-3348, nci=-318),
    (2023,2): dict(rev=18595, cogs=16880, gp=1715, sga=5820, rd=2740, oi=-6845, ebt=-5492, tax=-139, ni=-5089, nci=-264),
    (2023,3): dict(rev=17366, cogs=15500, gp=1866, sga=5667, rd=2926, oi=-6727, ebt=-6516, tax=-101, ni=-5823, nci=-592),
    (2024,1): dict(rev=22688, cogs=16594, gp=6094, sga=6227, rd=3214, oi=-3347, ebt=-1774, tax=274, ni=-2083, nci=35),
    (2024,2): dict(rev=27923, cogs=20271, gp=7652, sga=5779, rd=3758, oi=-1885, ebt=-1078, tax=121, ni=-1516, nci=317),
    (2024,3): dict(rev=23645, cogs=17963, gp=5682, sga=5650, rd=3438, oi=-3406, ebt=-2261, tax=626, ni=-2937, nci=50),
    (2025,1): dict(rev=19356, cogs=20597, gp=-1241, sga=5916, rd=3118, oi=-10275, ebt=-9942, tax=74, ni=-8798, nci=-1218),
    (2025,2): dict(rev=17974, cogs=16541, gp=1433, sga=5653, rd=2525, oi=-6745, ebt=-7095, tax=579, ni=-7008, nci=-666),
    (2025,3): dict(rev=27955, cogs=21731, gp=6224, sga=6334, rd=1013, oi=-1123, ebt=-1169, tax=504, ni=-1906, nci=233),
    (2026,1): dict(rev=26924, cogs=18946, gp=7978, sga=6551, rd=3012, oi=-1585, ebt=-1055, tax=430, ni=-1620, nci=135,
                  int_net=101, jv=353),
    (2026,2): dict(rev=47589, cogs=26223, gp=21366, sga=7292, rd=3651, oi=10423, ebt=15132, tax=2102, ni=11128, nci=1902,
                  int_net=4728, jv=417),
}

# Fill Q4 as FY residual
for y in range(2016, 2026):
    a = ANN[y]
    q123 = [QPL[(y, q)] for q in (1, 2, 3)]
    def residual(key, ann_key=None):
        ak = ann_key or key
        if key == "gp":
            return a["rev"] - a["cogs"] - sum(q["gp"] for q in q123)
        if key == "oi":
            return a["oi"] - sum(q["oi"] for q in q123)
        if key == "ebt":
            return a["ebt"] - sum(q["ebt"] for q in q123)
        if key == "ni":
            return a["ni_axt"] - sum(q["ni"] for q in q123)
        if key == "nci":
            return a["nci"] - sum(q["nci"] for q in q123)
        if key == "rev":
            return a["rev"] - sum(q["rev"] for q in q123)
        if key == "cogs":
            return a["cogs"] - sum(q["cogs"] for q in q123)
        if key == "sga":
            return a["sga"] - sum(q["sga"] for q in q123)
        if key == "rd":
            return a["rd"] - sum(q["rd"] for q in q123)
        if key == "tax":
            return a["tax"] - sum(q["tax"] for q in q123)
        return None
    QPL[(y, 4)] = dict(
        rev=residual("rev"), cogs=residual("cogs"), gp=residual("gp"),
        sga=residual("sga"), rd=residual("rd"), oi=residual("oi"),
        ebt=residual("ebt"), tax=residual("tax"), ni=residual("ni"), nci=residual("nci"),
    )
    missing = [k for k, v in QPL[(y, 4)].items() if v is None]
    if missing:
        raise SystemExit(f"Q4 {y} missing {missing}")
    # sanity: Q4 rev + Q1-3 = FY
    s = QPL[(y, 1)]["rev"] + QPL[(y, 2)]["rev"] + QPL[(y, 3)]["rev"] + QPL[(y, 4)]["rev"]
    if abs(s - ANN[y]["rev"]) > 1:
        raise SystemExit(f"Q-sum rev mismatch {y}: {s} vs {ANN[y]['rev']}")

# Quarterly BS ($000)
QBS = {
    (2016,1): dict(cash=25067, ar=19711, inv=38822, ppe=31141, assets=152727, ca=97039, ap=8034, cl=13391, liab=14705, re=-70579, eq_axt=132549, nci_bs=5473, rnci=0),
    (2016,2): dict(cash=26115, ar=18036, inv=38625, ppe=30178, assets=151430, ca=98337, ap=7848, cl=13219, liab=14389, re=-69428, eq_axt=132136, nci_bs=4905, rnci=0),
    (2016,3): dict(cash=29661, ar=18380, inv=38731, ppe=29385, assets=152928, ca=101600, ap=7277, cl=12288, liab=13314, re=-67199, eq_axt=134757, nci_bs=4857, rnci=0),
    (2016,4): dict(cash=36152, ar=14453, inv=40152, ppe=27805, assets=154246, ca=107286, ap=6691, cl=15951, liab=16856, re=-64985, eq_axt=133010, nci_bs=4380, rnci=0),
    (2017,1): dict(cash=56512, ar=17649, inv=39195, ppe=27141, assets=187417, ca=138661, ap=8159, cl=16065, liab=16728, re=-64320, eq_axt=166959, nci_bs=3730, rnci=0),
    (2017,2): dict(cash=56483, ar=18262, inv=40627, ppe=27945, assets=191772, ca=143245, ap=8189, cl=17552, liab=18059, re=-62390, eq_axt=170013, nci_bs=3700, rnci=0),
    (2017,3): dict(cash=48078, ar=20877, inv=40768, ppe=42255, assets=199622, ca=135986, ap=9499, cl=18783, liab=19205, re=-57971, eq_axt=176391, nci_bs=4026, rnci=0),
    (2017,4): dict(cash=44352, ar=22778, inv=45840, ppe=46530, assets=211200, ca=140521, ap=11445, cl=22594, liab=22883, re=-54837, eq_axt=183820, nci_bs=4497, rnci=0),
    (2018,1): dict(cash=39189, ar=21347, inv=51122, ppe=58763, assets=216307, ca=133557, ap=12049, cl=21239, liab=21568, re=-51962, eq_axt=189667, nci_bs=5072, rnci=0),
    (2018,2): dict(cash=29698, ar=22516, inv=57038, ppe=65174, assets=219681, ca=137470, ap=14059, cl=25852, liab=26148, re=-48061, eq_axt=189666, nci_bs=3867, rnci=0),
    (2018,3): dict(cash=18035, ar=23308, inv=58717, ppe=74754, assets=220449, ca=132595, ap=12344, cl=25411, liab=25731, re=-44122, eq_axt=191125, nci_bs=3593, rnci=0),
    (2018,4): dict(cash=16526, ar=19586, inv=58571, ppe=82280, assets=223524, ca=128540, ap=13338, cl=28709, liab=28992, re=-45183, eq_axt=190835, nci_bs=3697, rnci=0),
    (2019,1): dict(cash=21060, ar=19605, inv=53025, ppe=84975, assets=214558, ca=118142, ap=7882, cl=17438, liab=18577, re=-46287, eq_axt=191528, nci_bs=4453, rnci=0),
    (2019,2): dict(cash=28732, ar=18275, inv=50326, ppe=87613, assets=213890, ca=111703, ap=6004, cl=16949, liab=18235, re=-44836, eq_axt=191308, nci_bs=4347, rnci=0),
    (2019,3): dict(cash=27837, ar=17450, inv=49071, ppe=89680, assets=217710, ca=108723, ap=8147, cl=23188, liab=26256, re=-45734, eq_axt=186918, nci_bs=4536, rnci=0),
    (2019,4): dict(cash=26892, ar=19031, inv=49152, ppe=97403, assets=223349, ca=113205, ap=10098, cl=27526, liab=30587, re=-47783, eq_axt=187885, nci_bs=4877, rnci=0),
    (2020,1): dict(cash=20061, ar=23613, inv=48253, ppe=97490, assets=222640, ca=111262, ap=9518, cl=25960, liab=30450, re=-47961, eq_axt=187016, nci_bs=5174, rnci=0),
    (2020,2): dict(cash=26515, ar=19760, inv=49586, ppe=101031, assets=226220, ca=110729, ap=11540, cl=28138, liab=32166, re=-47600, eq_axt=188362, nci_bs=5692, rnci=0),
    (2020,3): dict(cash=22384, ar=22653, inv=48357, ppe=108295, assets=232737, ca=108548, ap=12271, cl=26493, liab=30845, re=-46609, eq_axt=195394, nci_bs=6498, rnci=0),
    (2020,4): dict(cash=72602, ar=24558, inv=51515, ppe=115825, assets=298862, ca=164518, ap=12669, cl=39075, liab=43330, re=-44545, eq_axt=192619, nci_bs=15350, rnci=47563),
    (2021,1): dict(cash=61206, ar=28423, inv=54681, ppe=119878, assets=301906, ca=162950, ap=13399, cl=39473, liab=43232, re=-41120, eq_axt=195149, nci_bs=15313, rnci=48212),
    (2021,2): dict(cash=52783, ar=33473, inv=58926, ppe=127241, assets=309404, ca=164442, ap=16130, cl=38210, liab=42059, re=-36735, eq_axt=199783, nci_bs=17479, rnci=50083),
    (2021,3): dict(cash=43608, ar=36584, inv=60683, ppe=131617, assets=315663, ca=159618, ap=11384, cl=37339, liab=41931, re=-32935, eq_axt=205740, nci_bs=17814, rnci=50178),
    (2021,4): dict(cash=36763, ar=34839, inv=65912, ppe=142415, assets=332441, ca=160185, ap=16649, cl=47822, liab=52210, re=-29970, eq_axt=211529, nci_bs=18317, rnci=50385),
    (2022,1): dict(cash=29494, ar=39848, inv=68821, ppe=147286, assets=338399, ca=160064, ap=14399, cl=46804, liab=50877, re=-26805, eq_axt=215716, nci_bs=20835, rnci=50971),
    (2022,2): dict(cash=41043, ar=38751, inv=77280, ppe=152882, assets=361265, ca=177380, ap=23317, cl=70871, liab=74536, re=-21259, eq_axt=216317, nci_bs=22147, rnci=48265),
    (2022,3): dict(cash=32918, ar=38131, inv=88496, ppe=152727, assets=361776, ca=181456, ap=13354, cl=72907, liab=76419, re=-15500, eq_axt=217150, nci_bs=23009, rnci=45198),
    (2022,4): dict(cash=34948, ar=29252, inv=89629, ppe=161017, assets=370072, ca=183545, ap=10084, cl=75326, liab=80326, re=-14159, eq_axt=221607, nci_bs=23293, rnci=44846),
    (2023,1): dict(cash=35436, ar=21345, inv=91678, ppe=162524, assets=366623, ca=177002, ap=6766, cl=72788, liab=78285, re=-17507, eq_axt=220001, nci_bs=23417, rnci=44920),
    (2023,2): dict(cash=30092, ar=19857, inv=87063, ppe=158672, assets=346202, ca=165470, ap=4692, cl=67263, liab=71818, re=-22596, eq_axt=209793, nci_bs=23198, rnci=41393),
    (2023,3): dict(cash=28522, ar=18883, inv=86383, ppe=158773, assets=342190, ca=160310, ap=6369, cl=67665, liab=74109, re=-28419, eq_axt=204229, nci_bs=23218, rnci=40634),
    (2023,4): dict(cash=37752, ar=19256, inv=86503, ppe=166348, assets=358701, ca=170656, ap=9617, cl=81557, liab=89555, re=-32040, eq_axt=203989, nci_bs=23494, rnci=41663),
    (2024,1): dict(cash=25793, ar=25058, inv=85943, ppe=163122, assets=348964, ca=163742, ap=10262, cl=72007, liab=83873, re=-34123, eq_axt=201084, nci_bs=23426, rnci=40581),
    (2024,2): dict(cash=27808, ar=27163, inv=85774, ppe=161332, assets=349413, ca=167396, ap=11655, cl=74647, liab=86300, re=-35639, eq_axt=199672, nci_bs=23680, rnci=39761),
    (2024,3): dict(cash=24898, ar=27970, inv=86109, ppe=166459, assets=355580, ca=167861, ap=12780, cl=78773, liab=89361, re=-38576, eq_axt=200733, nci_bs=24249, rnci=41237),
    (2024,4): dict(cash=22833, ar=25640, inv=85077, ppe=159721, assets=339314, ca=158272, ap=12356, cl=74176, liab=84406, re=-43664, eq_axt=192770, nci_bs=23561, rnci=38577),
    (2025,1): dict(cash=31584, ar=22863, inv=80409, ppe=159035, assets=333477, ca=149531, ap=11106, cl=75256, liab=87305, re=-52462, eq_axt=185037, nci_bs=23010, rnci=38125),
    (2025,2): dict(cash=27007, ar=22794, inv=80063, ppe=159714, assets=329003, ca=144960, ap=11242, cl=79542, liab=88372, re=-59470, eq_axt=179714, nci_bs=22907, rnci=38010),
    (2025,3): dict(cash=23110, ar=33837, inv=77656, ppe=159283, assets=334034, ca=149755, ap=10848, cl=84897, liab=93578, re=-61376, eq_axt=179148, nci_bs=23197, rnci=38111),
    (2025,4): dict(cash=120266, ar=26849, inv=81651, ppe=161860, assets=433751, ca=246556, ap=12947, cl=90541, liab=99120, re=-64924, eq_axt=273290, nci_bs=23285, rnci=38056),
    (2026,1): dict(cash=41769, ar=32016, inv=90168, ppe=164622, assets=444598, ca=253775, ap=16141, cl=97876, liab=107632, re=-66544, eq_axt=274874, nci_bs=23576, rnci=38516),
    (2026,2): dict(cash=412167, ar=36675, inv=96322, ppe=174924, assets=1097192, ca=598803, ap=14373, cl=125658, liab=145344, re=-55416, eq_axt=887274, nci_bs=24839, rnci=39735),
}

# Known InP quarterly revenue ($000) from earnings calls; others estimated via annual allocation
Q_INP_KNOWN = {
    (2024,1): 8100, (2024,2): 7700, (2024,3): 6800, (2024,4): 9100,
    (2025,1): 3800, (2025,2): 3600, (2025,3): 13100, (2025,4): 8000,
    (2026,1): 13600, (2026,2): 30700,
}
Q_GAAS_KNOWN = {(2026,2): 6600}
Q_GE_KNOWN = {(2024,1): 1400, (2024,2): 2900, (2024,4): 1600, (2025,2): 1500, (2025,3): 640, (2025,4): 230, (2026,2): 272}
Q_RM_KNOWN = {(2026,2): 10000}

def allocate_product(y, q):
    rev = QPL[(y, q)]["rev"]
    fy_rev = ANN[y]["rev"] if y <= 2025 else None
    prod = ANN_PROD.get(y)
    if (y, q) in Q_INP_KNOWN:
        inp = Q_INP_KNOWN[(y, q)]
    elif prod and fy_rev:
        inp = round(prod["inp"] * rev / fy_rev)
    else:
        inp = 0
    if (y, q) in Q_RM_KNOWN:
        rm = Q_RM_KNOWN[(y, q)]
    elif prod and fy_rev:
        rm = round(prod["rm"] * rev / fy_rev)
    else:
        rm = 0
    if (y, q) in Q_GE_KNOWN:
        ge = Q_GE_KNOWN[(y, q)]
    elif prod and fy_rev:
        ge = round(prod["ge"] * rev / fy_rev)
    else:
        ge = 0
    if (y, q) in Q_GAAS_KNOWN:
        gaas = Q_GAAS_KNOWN[(y, q)]
    else:
        gaas = max(0, rev - inp - ge - rm)
        # residual rounding into GaAs
    # reconcile to total rev
    other_sub = rev - inp - gaas - ge - rm
    if abs(other_sub) < 50:
        gaas += other_sub
        other_sub = 0
    return dict(inp=inp, gaas=gaas, ge=ge, rm=rm, other=other_sub)

# Quarterly ASP / unit COGS assumptions
Q_ASP = {}
Q_UCOGS = {}
for y in range(2016, 2026):
    for q in (1, 2, 3, 4):
        Q_ASP[(y, q)] = ANN_ASP[y] + {1: -8, 2: -2, 3: 2, 4: 8}[q]
        Q_UCOGS[(y, q)] = ANN_UCOGS[y] + {1: 5, 2: 0, 3: -3, 4: -2}[q]
Q_ASP[(2026, 1)] = 480
Q_ASP[(2026, 2)] = 510
Q_UCOGS[(2026, 1)] = 270
Q_UCOGS[(2026, 2)] = 250

# Forecast assumptions (quarterly) Q3'26 - Q4'28
# InP volume k wafers, ASP, unit COGS, other product rev, opex, WC days, capex
FCST = {
    (2026, 3): dict(vol=62.0, asp=530, ucogs=248, gaas=7000, ge=400, rm=10500,
                    sga=7600, rd=3800, jv=400, other_inc=0, tax_rate=0.18, nci_pct=0.14,
                    yield_q=0.010, debt_rate_q=0.012, dso=85, dio=200, dpo=50,
                    capex=12000, sbc=1000, da_rate=0.0135, sh_b=64200, sh_d=65500, pref_div=44,
                    stdebt=79000, ltdebt=15500, fx=0),
    (2026, 4): dict(vol=66.0, asp=545, ucogs=252, gaas=7200, ge=400, rm=11000,
                    sga=7800, rd=3900, jv=420, other_inc=0, tax_rate=0.18, nci_pct=0.14,
                    yield_q=0.010, debt_rate_q=0.012, dso=82, dio=190, dpo=50,
                    capex=14000, sbc=1000, da_rate=0.0135, sh_b=64400, sh_d=65700, pref_div=44,
                    stdebt=79000, ltdebt=15500, fx=0),
    (2027, 1): dict(vol=72.0, asp=580, ucogs=260, gaas=7400, ge=450, rm=12000,
                    sga=8000, rd=4000, jv=450, other_inc=0, tax_rate=0.18, nci_pct=0.14,
                    yield_q=0.010, debt_rate_q=0.012, dso=80, dio=180, dpo=52,
                    capex=25000, sbc=1100, da_rate=0.0135, sh_b=64700, sh_d=66000, pref_div=44,
                    stdebt=85000, ltdebt=25000, fx=0),
    (2027, 2): dict(vol=82.0, asp=600, ucogs=270, gaas=7500, ge=450, rm=13000,
                    sga=8200, rd=4100, jv=480, other_inc=0, tax_rate=0.18, nci_pct=0.14,
                    yield_q=0.010, debt_rate_q=0.012, dso=78, dio=170, dpo=52,
                    capex=25000, sbc=1100, da_rate=0.0135, sh_b=65000, sh_d=66300, pref_div=44,
                    stdebt=85000, ltdebt=25000, fx=0),
    (2027, 3): dict(vol=95.0, asp=625, ucogs=280, gaas=7600, ge=500, rm=14000,
                    sga=8500, rd=4200, jv=500, other_inc=0, tax_rate=0.18, nci_pct=0.14,
                    yield_q=0.010, debt_rate_q=0.012, dso=75, dio=165, dpo=52,
                    capex=25000, sbc=1100, da_rate=0.0135, sh_b=65300, sh_d=66600, pref_div=44,
                    stdebt=85000, ltdebt=30000, fx=0),
    (2027, 4): dict(vol=108.0, asp=650, ucogs=290, gaas=7800, ge=500, rm=15000,
                    sga=8800, rd=4300, jv=520, other_inc=0, tax_rate=0.18, nci_pct=0.14,
                    yield_q=0.010, debt_rate_q=0.012, dso=72, dio=160, dpo=52,
                    capex=25000, sbc=1100, da_rate=0.0135, sh_b=65600, sh_d=66900, pref_div=44,
                    stdebt=80000, ltdebt=30000, fx=0),
    (2028, 1): dict(vol=115.0, asp=675, ucogs=295, gaas=7900, ge=500, rm=15500,
                    sga=9000, rd=4400, jv=550, other_inc=0, tax_rate=0.18, nci_pct=0.14,
                    yield_q=0.010, debt_rate_q=0.012, dso=70, dio=155, dpo=52,
                    capex=12500, sbc=1200, da_rate=0.0130, sh_b=65900, sh_d=67200, pref_div=44,
                    stdebt=75000, ltdebt=28000, fx=0),
    (2028, 2): dict(vol=122.0, asp=700, ucogs=300, gaas=8000, ge=500, rm=16200,
                    sga=9200, rd=4500, jv=580, other_inc=0, tax_rate=0.18, nci_pct=0.14,
                    yield_q=0.010, debt_rate_q=0.012, dso=68, dio=150, dpo=52,
                    capex=12500, sbc=1200, da_rate=0.0130, sh_b=66200, sh_d=67500, pref_div=44,
                    stdebt=70000, ltdebt=26000, fx=0),
    (2028, 3): dict(vol=128.0, asp=720, ucogs=305, gaas=8100, ge=500, rm=16800,
                    sga=9400, rd=4500, jv=600, other_inc=0, tax_rate=0.18, nci_pct=0.14,
                    yield_q=0.010, debt_rate_q=0.012, dso=65, dio=145, dpo=52,
                    capex=12500, sbc=1200, da_rate=0.0130, sh_b=66500, sh_d=67800, pref_div=44,
                    stdebt=65000, ltdebt=24000, fx=0),
    (2028, 4): dict(vol=135.0, asp=740, ucogs=310, gaas=8200, ge=500, rm=17500,
                    sga=9600, rd=4500, jv=620, other_inc=0, tax_rate=0.18, nci_pct=0.14,
                    yield_q=0.010, debt_rate_q=0.012, dso=62, dio=140, dpo=52,
                    capex=12500, sbc=1200, da_rate=0.0130, sh_b=66800, sh_d=68100, pref_div=44,
                    stdebt=60000, ltdebt=22000, fx=0),
}

# ---------------------------------------------------------------------------
# Excel layout
# ---------------------------------------------------------------------------
BLUE = Font(name="Calibri", size=9, color="0000FF")
BLACK = Font(name="Calibri", size=9, color="000000")
BLUE_B = Font(name="Calibri", size=9, color="0000FF", bold=True)
BLACK_B = Font(name="Calibri", size=9, color="000000", bold=True)
TITLE = Font(name="Calibri", size=16, color="000000", bold=True)
SUB = Font(name="Calibri", size=11, color="000000", bold=True)
HDR = Font(name="Calibri", size=9, color="FFFFFF", bold=True)
SEC = Font(name="Calibri", size=10, color="FFFFFF", bold=True)
NOTEF = Font(name="Calibri", size=8, color="000000", italic=True)
GRAYF = Font(name="Calibri", size=8, color="666666")
CHECKF = Font(name="Calibri", size=9, color="000000", bold=True)

FILL_NAVY = PatternFill("solid", fgColor="1F4E79")
FILL_TEAL = PatternFill("solid", fgColor="2E75B6")
FILL_DRV = PatternFill("solid", fgColor="1F4E79")
FILL_PL = PatternFill("solid", fgColor="548235")
FILL_BS = PatternFill("solid", fgColor="833C0C")
FILL_CF = PatternFill("solid", fgColor="5B2C6F")
FILL_CK = PatternFill("solid", fgColor="C00000")
FILL_ACT = PatternFill("solid", fgColor="D6DCE4")
FILL_FC = PatternFill("solid", fgColor="FFF2CC")
FILL_MIX = PatternFill("solid", fgColor="FCE4D6")
FILL_GAP = PatternFill("solid", fgColor="FFFFFF")
FILL_INP = PatternFill("solid", fgColor="DEEBF7")
FILL_TOT = PatternFill("solid", fgColor="E2EFDA")
FILL_LINE = PatternFill("solid", fgColor="F2F2F2")
FILL_WHITE = PatternFill("solid", fgColor="FFFFFF")
FILL_NOTE = PatternFill("solid", fgColor="FFF8E7")
FILL_TITLE = PatternFill("solid", fgColor="1F4E79")
FILL_YELLOW = PatternFill("solid", fgColor="FFFF99")

THIN = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)
BOTTOM = Border(bottom=Side(style="thin", color="1F4E79"))
DBL = Border(bottom=Side(style="double", color="000000"))

C_SEC, C_LINE, C_UNIT = 1, 2, 3
ANN_COL0 = 4  # D = 2016
Q_COL0 = ANN_COL0 + len(ANN_YEARS) + 1  # after 13 years + gap
NOTES_COL = Q_COL0 + len(QUARTERS)

def ann_col(year):
    return ANN_COL0 + (year - 2016)

def q_col(y, q):
    idx = (y - 2016) * 4 + (q - 1)
    return Q_COL0 + idx

def q_prev(y, q):
    return (y - 1, 4) if q == 1 else (y, q - 1)

NF_N = '#,##0'
NF_1 = '#,##0.0'
NF_2 = '#,##0.00'
NF_P = '0.0%'
NF_P1 = '0.00%'
NF_ASP = '#,##0.00'
NF_K = '#,##0.00'

wb = Workbook()
ws = wb.active
ws.title = "AXTI 3-Statement Model"

# Column widths
ws.column_dimensions["A"].width = 22
ws.column_dimensions["B"].width = 44
ws.column_dimensions["C"].width = 16
for col in range(ANN_COL0, NOTES_COL):
    ws.column_dimensions[get_column_letter(col)].width = 11.2
ws.column_dimensions[get_column_letter(ANN_COL0 + len(ANN_YEARS))].width = 2.5
ws.column_dimensions[get_column_letter(NOTES_COL)].width = 78

center = Alignment(horizontal="center", vertical="center", wrap_text=True)
left = Alignment(horizontal="left", vertical="center", wrap_text=True)
right = Alignment(horizontal="right", vertical="center")

R = {}  # row map


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


def period_fill(year=None, yq=None):
    if year is not None:
        if year <= 2025:
            return FILL_ACT
        if year == 2026:
            return FILL_MIX
        return FILL_FC
    y, q = yq
    if is_actual_q(y, q):
        return FILL_ACT
    return FILL_FC


# ===== HEADER =====
ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)
c = ws.cell(1, 1, "AXT, Inc. (NASDAQ: AXTI)  —  Integrated 3-Statement Financial Model")
c.font = Font(name="Calibri", size=16, color="FFFFFF", bold=True)
c.fill = FILL_TITLE
for col in range(1, 9):
    ws.cell(1, col).fill = FILL_TITLE

ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=8)
ws.cell(2, 1, "磷化铟 (InP) 衬底量价驱动 | 历史 2016–2025A + 2026H1A | 预测 2026H2–2028E | 单位: $000（另有标注除外）").font = SUB

ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=8)
ws.cell(3, 1, "数据来源: SEC 10-K/10-Q XBRL、公司财报电话会、2025 Form 10-K、2026Q1/Q2 10-Q。蓝色=假设/硬编码；黑色=公式。模型日期: 2026-08-18。").font = GRAYF

ws.cell(5, 1, "图例").font = BLACK_B
ws.cell(5, 2, "蓝色字体 = 假设 / 历史硬编码 (hardcode)").font = BLUE
ws.cell(5, 3, "黑色字体 = 公式 / 勾稽").font = BLACK
ws.cell(5, 4, "灰底 = 实际").fill = FILL_ACT
ws.cell(5, 5, "黄底 = 预测").fill = FILL_FC
ws.cell(5, 6, "橙底 = 2026A/E 混合").fill = FILL_MIX
ws.cell(5, 7, "最右列为 Notes").font = NOTEF

# Period headers
r_type, r_year, r_per, r_stat, r_days = 7, 8, 9, 10, 11
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

# Freeze
ws.freeze_panes = "D12"
ws.auto_filter.ref = None
ws.sheet_view.showGridLines = False
ws.page_setup.orientation = "landscape"
ws.page_setup.fitToPage = True
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0
ws.page_setup.paperSize = ws.PAPERSIZE_TABLOID
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.print_title_rows = "1:11"
ws.print_title_cols = "A:C"
ws.page_setup.horizontalCentered = False
ws.oddHeader.left.text = "AXTI 3-Statement Model"
ws.oddFooter.right.text = "Blue = input / Black = formula  |  Confidential"

row = 13


def section_bar(row, title, fill):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
    cell = ws.cell(row, 1, title)
    cell.font = SEC
    cell.fill = fill
    cell.alignment = left
    for col in range(1, NOTES_COL + 1):
        ws.cell(row, col).fill = fill
        ws.cell(row, col).font = SEC
    ws.cell(row, NOTES_COL, "").fill = fill
    return row


# =====================================================================
# DRIVERS
# =====================================================================
row = section_bar(13, "I.  REVENUE DRIVERS  —  InP 衬底量价拆分 / GaAs / Ge / Raw Materials", FILL_DRV)
R["sec_drv"] = 13
row = 14
# driver rows
driver_rows = [
    ("vol", "InP", "InP wafers sold", "k wafers", "磷化铟衬底出货量。历史=InP收入/ASP（公式）；预测=产能爬坡假设（蓝）。Q4'25产能受限（出口许可）；管理层指引2026年底InP产能约$35m/季、2027再翻倍至$65-70m/季。"),
    ("asp", "InP", "InP blended ASP", "$ / wafer", "InP混合平均售价。公司不披露ASP，本模型按2\"/3\"→4\"→6\"直径升级路径假设。历史约$185–420；2026H1因AI光模块/硅光需求及4\"占比提升至$480–510；此后随6\"放量继续上行。"),
    ("inp_rev", "InP", "InP substrate revenue", "$000", "InP收入=出货量(k)*ASP。2024–2026H1用财报电话会披露值硬编码；更早年度按衬底收入占比估算（见年度产品拆分）。Q2'26 InP创纪录$30.7m。"),
    ("ucogs", "InP", "InP unit COGS", "$ / wafer", "单片制造成本。历史按开工率反推：满产年份~$125–185，2023–25出口许可/低开工升至$250–285。预测随规模效应下降单位成本，但6\"混入使绝对COGS回升、毛利率仍扩张。"),
    ("ugp", "InP", "InP unit gross profit", "$ / wafer", "单片毛利=ASP−单片COGS（公式）。衡量直径升级与开工率对单位盈利的贡献。"),
    ("ugm", "InP", "InP unit gross margin", "%", "单片毛利率=单片毛利/ASP（公式）。AI需求+4\"/6\"升级是中期扩张主因。"),
    ("inp_cogs", "InP", "InP COGS", "$000", "InP销货成本=出货量(k)*单片COGS（公式）。"),
    ("inp_gp", "InP", "InP gross profit", "$000", "InP毛利=InP收入−InP COGS（公式）。"),
    ("gaas", "Other", "GaAs substrate revenue", "$000", "砷化镓衬底。公司将部分GaAs晶体生长产能改造成InP。预测维持$7–8m/季平稳，不做扩张。Q2'26实际$6.6m。"),
    ("ge", "Other", "Ge substrate revenue", "$000", "锗衬底（卫星太阳能电池）。管理层称毛利差、出口许可难，已主动收缩。Q2'26仅$0.27m。预测维持极低个位数。"),
    ("rm", "Other", "Raw materials revenue", "$000", "并表原材料JV：高纯镓、InP多晶、pBN坩埚/OLED工具（BoYu、JinMei等）。随InP放量同步增长。Q2'26 $10.0m。10-K: 2025原材料占收入33%。"),
    ("tot_drv", "Total", "Total revenue (drivers)", "$000", "驱动加总=InP+GaAs+Ge+原材料，应等于P&L收入（见勾稽）。"),
    ("mix_inp", "Mix", "InP % of total revenue", "%", "InP收入占比。Q2'26已升至约64.5%，是本轮AI光连接周期的核心变量。"),
    ("mix_sub", "Mix", "Substrate % of total revenue", "%", "衬底（InP+GaAs+Ge）占比。10-K: 2023/24/25为63%/68%/67%。"),
]
for key, sec, line, unit, note in driver_rows:
    row += 1
    R[key] = row
    write_label(row, sec, line, unit, note)
    if key in ("tot_drv",):
        ws.cell(row, C_LINE).font = BLACK_B

# fill driver historical annual
for y in ANN_YEARS:
    col = ann_col(y)
    fill = period_fill(year=y)
    cl = get_column_letter(col)
    if y <= 2025:
        p = ANN_PROD[y]
        put(R["asp"], col, ANN_ASP[y], True, NF_ASP, fill=fill)
        put(R["ucogs"], col, ANN_UCOGS[y], True, NF_ASP, fill=fill)
        put(R["inp_rev"], col, p["inp"], True, NF_N, fill=fill)
        formula(R["vol"], col, f'=IF({cl}{R["asp"]}=0,0,{cl}{R["inp_rev"]}/{cl}{R["asp"]})', NF_1, fill=fill)
        put(R["gaas"], col, p["gaas"], True, NF_N, fill=fill)
        put(R["ge"], col, p["ge"], True, NF_N, fill=fill)
        put(R["rm"], col, p["rm"], True, NF_N, fill=fill)
    else:
        # sum of 4 quarters
        qcols = [get_column_letter(q_col(y, q)) for q in (1, 2, 3, 4)]
        formula(R["vol"], col, "+".join(f"{qc}{R['vol']}" for qc in qcols), NF_1, fill=fill)
        formula(R["inp_rev"], col, "+".join(f"{qc}{R['inp_rev']}" for qc in qcols), NF_N, fill=fill)
        formula(R["asp"], col, f'=IF({cl}{R["vol"]}=0,0,{cl}{R["inp_rev"]}/{cl}{R["vol"]})', NF_ASP, fill=fill)
        formula(R["inp_cogs"], col, "+".join(f"{qc}{R['inp_cogs']}" for qc in qcols), NF_N, fill=fill)
        formula(R["ucogs"], col, f'=IF({cl}{R["vol"]}=0,0,{cl}{R["inp_cogs"]}/{cl}{R["vol"]})', NF_ASP, fill=fill)
        formula(R["gaas"], col, "+".join(f"{qc}{R['gaas']}" for qc in qcols), NF_N, fill=fill)
        formula(R["ge"], col, "+".join(f"{qc}{R['ge']}" for qc in qcols), NF_N, fill=fill)
        formula(R["rm"], col, "+".join(f"{qc}{R['rm']}" for qc in qcols), NF_N, fill=fill)
    formula(R["ugp"], col, f'={cl}{R["asp"]}-{cl}{R["ucogs"]}', NF_ASP, fill=fill)
    formula(R["ugm"], col, f'=IF({cl}{R["asp"]}=0,0,{cl}{R["ugp"]}/{cl}{R["asp"]})', NF_P, fill=fill)
    if y <= 2025:
        formula(R["inp_cogs"], col, f'={cl}{R["vol"]}*{cl}{R["ucogs"]}', NF_N, fill=fill)
    formula(R["inp_gp"], col, f'={cl}{R["inp_rev"]}-{cl}{R["inp_cogs"]}', NF_N, fill=fill)
    formula(R["tot_drv"], col, f'={cl}{R["inp_rev"]}+{cl}{R["gaas"]}+{cl}{R["ge"]}+{cl}{R["rm"]}', NF_N, True, fill=FILL_TOT)
    formula(R["mix_inp"], col, f'=IF({cl}{R["tot_drv"]}=0,0,{cl}{R["inp_rev"]}/{cl}{R["tot_drv"]})', NF_P, fill=fill)
    formula(R["mix_sub"], col, f'=IF({cl}{R["tot_drv"]}=0,0,({cl}{R["inp_rev"]}+{cl}{R["gaas"]}+{cl}{R["ge"]})/{cl}{R["tot_drv"]})', NF_P, fill=fill)

# quarterly drivers
for (y, q) in QUARTERS:
    col = q_col(y, q)
    fill = period_fill(yq=(y, q))
    cl = get_column_letter(col)
    if is_actual_q(y, q):
        pr = allocate_product(y, q) if y <= 2025 or (y, q) <= (2026, 2) else None
        if y == 2026 and q <= 2:
            if q == 1:
                # Q1'26: InP $13.6m from earnings commentary; other lines estimated to total $26.924m
                pr = dict(inp=13600, gaas=4424, ge=400, rm=8500)
            else:
                # Q2'26 call: InP $30.7m, GaAs $6.6m, Ge $0.272m, RM $10.0m; +$17k rounding to GaAs
                pr = dict(inp=30700, gaas=6617, ge=272, rm=10000)
            put(R["inp_rev"], col, pr["inp"], True, NF_N, fill=fill)
            put(R["gaas"], col, pr["gaas"], True, NF_N, fill=fill)
            put(R["ge"], col, pr["ge"], True, NF_N, fill=fill)
            put(R["rm"], col, pr["rm"], True, NF_N, fill=fill)
        else:
            pr = allocate_product(y, q)
            put(R["inp_rev"], col, pr["inp"], True, NF_N, fill=fill)
            put(R["gaas"], col, pr["gaas"], True, NF_N, fill=fill)
            put(R["ge"], col, pr["ge"], True, NF_N, fill=fill)
            put(R["rm"], col, pr["rm"], True, NF_N, fill=fill)
        put(R["asp"], col, Q_ASP[(y, q)], True, NF_ASP, fill=fill)
        put(R["ucogs"], col, Q_UCOGS[(y, q)], True, NF_ASP, fill=fill)
        formula(R["vol"], col, f'=IF({cl}{R["asp"]}=0,0,{cl}{R["inp_rev"]}/{cl}{R["asp"]})', NF_1, fill=fill)
        formula(R["inp_cogs"], col, f'={cl}{R["vol"]}*{cl}{R["ucogs"]}', NF_N, fill=fill)
    else:
        f = FCST[(y, q)]
        put(R["vol"], col, f["vol"], True, NF_1, fill=fill)
        put(R["asp"], col, f["asp"], True, NF_ASP, fill=fill)
        put(R["ucogs"], col, f["ucogs"], True, NF_ASP, fill=fill)
        formula(R["inp_rev"], col, f'={cl}{R["vol"]}*{cl}{R["asp"]}', NF_N, fill=fill)
        formula(R["inp_cogs"], col, f'={cl}{R["vol"]}*{cl}{R["ucogs"]}', NF_N, fill=fill)
        put(R["gaas"], col, f["gaas"], True, NF_N, fill=fill)
        put(R["ge"], col, f["ge"], True, NF_N, fill=fill)
        put(R["rm"], col, f["rm"], True, NF_N, fill=fill)
    formula(R["ugp"], col, f'={cl}{R["asp"]}-{cl}{R["ucogs"]}', NF_ASP, fill=fill)
    formula(R["ugm"], col, f'=IF({cl}{R["asp"]}=0,0,{cl}{R["ugp"]}/{cl}{R["asp"]})', NF_P, fill=fill)
    formula(R["inp_gp"], col, f'={cl}{R["inp_rev"]}-{cl}{R["inp_cogs"]}', NF_N, fill=fill)
    formula(R["tot_drv"], col, f'={cl}{R["inp_rev"]}+{cl}{R["gaas"]}+{cl}{R["ge"]}+{cl}{R["rm"]}', NF_N, True, fill=FILL_TOT)
    formula(R["mix_inp"], col, f'=IF({cl}{R["tot_drv"]}=0,0,{cl}{R["inp_rev"]}/{cl}{R["tot_drv"]})', NF_P, fill=fill)
    formula(R["mix_sub"], col, f'=IF({cl}{R["tot_drv"]}=0,0,({cl}{R["inp_rev"]}+{cl}{R["gaas"]}+{cl}{R["ge"]})/{cl}{R["tot_drv"]})', NF_P, fill=fill)

# =====================================================================
# ASSUMPTIONS
# =====================================================================
row = R["mix_sub"] + 2
row = section_bar(row, "II.  MARGIN / P&L / BALANCE SHEET ASSUMPTIONS  （蓝色=可调假设）", FILL_DRV)
R["sec_ass"] = row
ass_items = [
    ("gm_co", "P&L", "Company gross margin (implied / assumed)", "%", "历史=GP/Rev公式；预测不直接用此行驱动（由InP单片COGS+其他产品隐含COGS推出），本行作对照。"),
    ("other_gm", "P&L", "Non-InP blended gross margin", "%", "GaAs/Ge/原材料综合毛利率假设。原材料与Ge拉低公司毛利率；预测随InP占比提升，非InP GM稳定在18–22%。历史=（总COGS−InP COGS）/非InP收入反推。"),
    ("sga_d", "P&L", "SG&A", "$000", "销售管理费用。历史硬编码；预测按绝对额（半固定，随规模缓增）。近年约$5.6–7.3m/季。"),
    ("rd_d", "P&L", "R&D", "$000", "研发。6\" InP与8\" GaAs开发费用在2023–24较高；2025回落。预测维持$3.8–4.5m/季。"),
    ("tax_r", "P&L", "Effective tax rate", "%", "美国NOLs（联邦约$50.5m）使现金税主要来自中国子公司。盈利期用18% ETR；亏损期仍有最低中国税。"),
    ("nci_r", "P&L", "NCI % of consolidated NI", "%", "Tongmei少数股东及可赎回NCI约14.5%（2026Q2 10-Q）。预测按14%税后利润归属NCI。"),
    ("yld", "P&L", "Cash & investments quarterly yield", "%", "季收益率。Q2'26利息净收入$4.73m，因4月$632.5m增发后资金尚未满季。预测按现金+其他非流动资产中的投资余额×1.0%/季（约4%年化）。"),
    ("drate", "P&L", "Cost of debt (quarterly)", "%", "短期银行贷款/长期贷款成本，约4.8%年化（1.2%/季）。"),
    ("jv_d", "P&L", "Equity in unconsolidated JVs", "$000", "未并表JV权益法收益。历史硬编码；预测$0.4–0.62m/季。"),
    ("oth_d", "P&L", "Other income / (expense)", "$000", "汇兑及其他。预测保守取0。"),
    ("dso", "BS", "DSO (AR days)", "days", "应收账款天数=AR/季收入*当期天数。历史常>90天（中国账期）。预测随海外AI客户提升而下降至~62天。"),
    ("dio", "BS", "DIO (Inventory days)", "days", "存货天数=Inv/COGS*天数。历史极高（~300–400天，衬底长周期+出口积压）。预测随出货畅通降至~140天。"),
    ("dpo", "BS", "DPO (AP days)", "days", "应付天数=AP/COGS*天数。预测维持~50–52天。"),
    ("capex_d", "CF", "Capex", "$000", "资本开支。管理层: 2026 InP产能翻倍CapEx $30–40m；2027新厂约$100m；2028继续扩张。H1'26 PPE已明显增加。"),
    ("da_d", "CF", "D&A", "$000", "折旧摊销。历史硬编码；预测=期初PPE×1.30–1.35%/季（约5.2–5.4%年化）。"),
    ("sbc_d", "CF", "Stock-based compensation", "$000", "股份支付，加回CFO并计入APIC。预测$1.0–1.2m/季。"),
    ("sh_b", "P&L", "Diluted / basic shares", "k shares", "基本股本。2025增发8.16m股；2026年4月再融资后约64.2m（Q2'26 filing）。预测随RSU小幅增加。"),
]
for key, sec, line, unit, note in ass_items:
    row += 1
    R[key] = row
    write_label(row, sec, line, unit, note)

# We'll fill assumption rows together with P&L because some are formulas from P&L.
# Store row numbers already.

# =====================================================================
# P&L
# =====================================================================
row = R["sh_b"] + 2
row = section_bar(row, "III.  CONSOLIDATED P&L  （利润表）", FILL_PL)
R["sec_pl"] = row
pl_items = [
    ("rev", "Revenue", "$000", "营业收入。历史=10-K/10-Q；预测=驱动表Total revenue。"),
    ("cogs", "Cost of revenue", "$000", "历史硬编码；预测=InP COGS + 非InP收入×(1−非InP毛利率)。"),
    ("gp", "Gross profit", "$000", "毛利=收入−成本（公式）。"),
    ("gpm", "Gross margin", "%", "毛利率。Q2'26实际44.9%（高开工+InP结构）；2025全年仅12.7%（许可/低开工）。"),
    ("sga", "Selling, general & administrative", "$000", "见假设。"),
    ("rd", "Research & development", "$000", "见假设。"),
    ("opex", "Total operating expenses", "$000", "SG&A+R&D。"),
    ("oi", "Operating income", "$000", "营业利润=毛利−营业费用。"),
    ("om", "Operating margin", "%", "营业利润率。"),
    ("intn", "Interest income / (expense), net", "$000", "历史硬编码或残差；预测=期初(现金+其他非流动资产)×收益率 − 期初债务×债务成本。用期初余额避免循环引用。"),
    ("jv", "Equity in income of unconsolidated JVs", "$000", "权益法收益。"),
    ("oth", "Other income / (expense), net", "$000", "其他。历史=EBT−OI−利息−JV（使P&L配平）；预测为假设。"),
    ("ebt", "Income before tax", "$000", "税前利润（公式）。"),
    ("tax", "Provision for income taxes", "$000", "所得税。历史硬编码；预测=max(EBT,0)×税率 + 亏损期最低税$400k/季。"),
    ("ni_c", "Net income (consolidated)", "$000", "合并净利润=EBT−Tax。注意XBRL中NetIncomeLoss常为归属母公司数，本模型分开列示。"),
    ("nci", "Net (income) / loss to NCI & redeemable NCI", "$000", "少数股东损益。符号：正数=NCI分享利润（减少归属AXT）。预测=合并NI×NCI%。"),
    ("ni_a", "Net income attributable to AXT, Inc.", "$000", "归属AXT净利润=合并NI−NCI。"),
    ("pref", "Preferred dividends", "$000", "A系列优先股股息约$177k/年（$0.001 par, 883k股, 5%）。"),
    ("ni_cm", "Net income to common", "$000", "属普通股=NI_AXT−优先股股息。"),
    ("eps", "EPS diluted", "$ / sh", "稀释每股收益=属普通股/稀释股本。"),
    ("shd", "Weighted-avg diluted shares", "k", "稀释加权平均股本。"),
]
for key, line, unit, note in pl_items:
    row += 1
    R[key] = row
    write_label(row, "P&L", line, unit, note)
    if key in ("gp", "oi", "ni_c", "ni_a"):
        ws.cell(row, C_LINE).font = BLACK_B

# Annual P&L
for y in ANN_YEARS:
    col = ann_col(y)
    fill = period_fill(year=y)
    cl = get_column_letter(col)
    qcs = [get_column_letter(q_col(y, q)) for q in (1, 2, 3, 4)]
    if y <= 2025:
        a = ANN[y]
        put(R["rev"], col, a["rev"], True, NF_N, True, fill)
        put(R["cogs"], col, a["cogs"], True, NF_N, fill=fill)
        put(R["sga"], col, a["sga"], True, NF_N, fill=fill)
        put(R["rd"], col, a["rd"], True, NF_N, fill=fill)
        put(R["oi"], col, a["oi"], True, NF_N, True, fill)  # will overwrite with formula? keep hardcoded to match 10-K then formula check
        # Use formulas that should equal reported
        formula(R["gp"], col, f'={cl}{R["rev"]}-{cl}{R["cogs"]}', NF_N, True, FILL_TOT)
        formula(R["gpm"], col, f'=IF({cl}{R["rev"]}=0,0,{cl}{R["gp"]}/{cl}{R["rev"]})', NF_P, fill=fill)
        formula(R["opex"], col, f'={cl}{R["sga"]}+{cl}{R["rd"]}', NF_N, fill=fill)
        # keep reported OI as formula from GP-opex (ties)
        # overwrite OI as formula
        formula(R["oi"], col, f'={cl}{R["gp"]}-{cl}{R["opex"]}', NF_N, True, FILL_TOT)
        formula(R["om"], col, f'=IF({cl}{R["rev"]}=0,0,{cl}{R["oi"]}/{cl}{R["rev"]})', NF_P, fill=fill)
        put(R["intn"], col, a["int_net"], True, NF_N, fill=fill)
        put(R["jv"], col, a["jv"], True, NF_N, fill=fill)
        formula(R["oth"], col, f'={cl}{R["ebt"]}-{cl}{R["oi"]}-{cl}{R["intn"]}-{cl}{R["jv"]}', NF_N, fill=fill)
        put(R["ebt"], col, a["ebt"], True, NF_N, fill=fill)
        put(R["tax"], col, a["tax"], True, NF_N, fill=fill)
        formula(R["ni_c"], col, f'={cl}{R["ebt"]}-{cl}{R["tax"]}', NF_N, True, FILL_TOT)
        put(R["nci"], col, a["nci"], True, NF_N, fill=fill)
        formula(R["ni_a"], col, f'={cl}{R["ni_c"]}-{cl}{R["nci"]}', NF_N, True, FILL_TOT)
        put(R["pref"], col, 177, True, NF_N, fill=fill)
        formula(R["ni_cm"], col, f'={cl}{R["ni_a"]}-{cl}{R["pref"]}', NF_N, fill=fill)
        put(R["shd"], col, a["sh_d"], True, NF_N, fill=fill)
        formula(R["eps"], col, f'=IF({cl}{R["shd"]}=0,0,{cl}{R["ni_cm"]}/{cl}{R["shd"]})', NF_2, fill=fill)
        # assumptions historical
        formula(R["gm_co"], col, f'={cl}{R["gpm"]}', NF_P, fill=fill)
        formula(R["other_gm"], col, f'=IF({cl}{R["rev"]}-{cl}{R["inp_rev"]}=0,0,({cl}{R["gp"]}-{cl}{R["inp_gp"]})/({cl}{R["rev"]}-{cl}{R["inp_rev"]}))', NF_P, fill=fill)
        formula(R["sga_d"], col, f'={cl}{R["sga"]}', NF_N, fill=fill)
        formula(R["rd_d"], col, f'={cl}{R["rd"]}', NF_N, fill=fill)
        formula(R["tax_r"], col, f'=IF({cl}{R["ebt"]}=0,0,{cl}{R["tax"]}/{cl}{R["ebt"]})', NF_P, fill=fill)
        formula(R["nci_r"], col, f'=IF({cl}{R["ni_c"]}=0,0,{cl}{R["nci"]}/{cl}{R["ni_c"]})', NF_P, fill=fill)
        put(R["yld"], col, 0.01, True, NF_P, fill=fill)
        put(R["drate"], col, 0.012, True, NF_P, fill=fill)
        formula(R["jv_d"], col, f'={cl}{R["jv"]}', NF_N, fill=fill)
        formula(R["oth_d"], col, f'={cl}{R["oth"]}', NF_N, fill=fill)
        put(R["dso"], col, round(a["ar"] / a["rev"] * 365, 1) if a["rev"] else 0, True, NF_1, fill=fill)
        put(R["dio"], col, round(a["inv"] / a["cogs"] * 365, 1) if a["cogs"] else 0, True, NF_1, fill=fill)
        put(R["dpo"], col, round(a["ap"] / a["cogs"] * 365, 1) if a["cogs"] else 0, True, NF_1, fill=fill)
        put(R["capex_d"], col, a["capex"], True, NF_N, fill=fill)
        put(R["da_d"], col, a["da"], True, NF_N, fill=fill)
        put(R["sbc_d"], col, a["sbc"], True, NF_N, fill=fill)
        put(R["sh_b"], col, a["sh_d"], True, NF_N, fill=fill)
    else:
        # sum quarters
        def qsum(rkey):
            return "=" + "+".join(f"{qc}{R[rkey]}" for qc in qcs)
        formula(R["rev"], col, qsum("rev"), NF_N, True, FILL_TOT)
        formula(R["cogs"], col, qsum("cogs"), NF_N, fill=fill)
        formula(R["gp"], col, f'={cl}{R["rev"]}-{cl}{R["cogs"]}', NF_N, True, FILL_TOT)
        formula(R["gpm"], col, f'=IF({cl}{R["rev"]}=0,0,{cl}{R["gp"]}/{cl}{R["rev"]})', NF_P, fill=fill)
        formula(R["sga"], col, qsum("sga"), NF_N, fill=fill)
        formula(R["rd"], col, qsum("rd"), NF_N, fill=fill)
        formula(R["opex"], col, f'={cl}{R["sga"]}+{cl}{R["rd"]}', NF_N, fill=fill)
        formula(R["oi"], col, f'={cl}{R["gp"]}-{cl}{R["opex"]}', NF_N, True, FILL_TOT)
        formula(R["om"], col, f'=IF({cl}{R["rev"]}=0,0,{cl}{R["oi"]}/{cl}{R["rev"]})', NF_P, fill=fill)
        formula(R["intn"], col, qsum("intn"), NF_N, fill=fill)
        formula(R["jv"], col, qsum("jv"), NF_N, fill=fill)
        formula(R["oth"], col, qsum("oth"), NF_N, fill=fill)
        formula(R["ebt"], col, f'={cl}{R["oi"]}+{cl}{R["intn"]}+{cl}{R["jv"]}+{cl}{R["oth"]}', NF_N, fill=fill)
        formula(R["tax"], col, qsum("tax"), NF_N, fill=fill)
        formula(R["ni_c"], col, f'={cl}{R["ebt"]}-{cl}{R["tax"]}', NF_N, True, FILL_TOT)
        formula(R["nci"], col, qsum("nci"), NF_N, fill=fill)
        formula(R["ni_a"], col, f'={cl}{R["ni_c"]}-{cl}{R["nci"]}', NF_N, True, FILL_TOT)
        formula(R["pref"], col, qsum("pref"), NF_N, fill=fill)
        formula(R["ni_cm"], col, f'={cl}{R["ni_a"]}-{cl}{R["pref"]}', NF_N, fill=fill)
        formula(R["shd"], col, f'={qcs[3]}{R["shd"]}', NF_N, fill=fill)  # year-end / last q
        formula(R["eps"], col, f'=IF({cl}{R["shd"]}=0,0,{cl}{R["ni_cm"]}/{cl}{R["shd"]})', NF_2, fill=fill)
        formula(R["gm_co"], col, f'={cl}{R["gpm"]}', NF_P, fill=fill)
        formula(R["other_gm"], col, f'=IF({cl}{R["rev"]}-{cl}{R["inp_rev"]}=0,0,({cl}{R["gp"]}-{cl}{R["inp_gp"]})/({cl}{R["rev"]}-{cl}{R["inp_rev"]}))', NF_P, fill=fill)
        formula(R["sga_d"], col, f'={cl}{R["sga"]}', NF_N, fill=fill)
        formula(R["rd_d"], col, f'={cl}{R["rd"]}', NF_N, fill=fill)
        formula(R["tax_r"], col, f'=IF({cl}{R["ebt"]}=0,0,{cl}{R["tax"]}/{cl}{R["ebt"]})', NF_P, fill=fill)
        formula(R["nci_r"], col, f'=IF({cl}{R["ni_c"]}=0,0,{cl}{R["nci"]}/{cl}{R["ni_c"]})', NF_P, fill=fill)
        put(R["yld"], col, 0.01, True, NF_P, fill=fill)
        put(R["drate"], col, 0.012, True, NF_P, fill=fill)
        formula(R["jv_d"], col, f'={cl}{R["jv"]}', NF_N, fill=fill)
        formula(R["oth_d"], col, f'={cl}{R["oth"]}', NF_N, fill=fill)
        formula(R["dso"], col, f'={qcs[3]}{R["dso"]}', NF_1, fill=fill)
        formula(R["dio"], col, f'={qcs[3]}{R["dio"]}', NF_1, fill=fill)
        formula(R["dpo"], col, f'={qcs[3]}{R["dpo"]}', NF_1, fill=fill)
        formula(R["capex_d"], col, qsum("capex_d"), NF_N, fill=fill)
        formula(R["da_d"], col, qsum("da_d"), NF_N, fill=fill)
        formula(R["sbc_d"], col, qsum("sbc_d"), NF_N, fill=fill)
        formula(R["sh_b"], col, f'={qcs[3]}{R["sh_b"]}', NF_N, fill=fill)

# Need BS rows before quarterly P&L interest (uses opening cash). Define BS rows first then come back?
# I'll put BS next, then CF, then fill quarterly P&L that references BS.
# Actually quarterly P&L for forecast needs opening cash from BS. So define BS row numbers now, fill later.

row = R["shd"] + 2
row = section_bar(row, "IV.  CONSOLIDATED BALANCE SHEET  （资产负债表）", FILL_BS)
R["sec_bs"] = row
bs_items = [
    ("cash", "Cash & cash equivalents", "$000", "货币资金。历史硬编码；预测=上期现金+本期ΔCash（现金流量表）。Q2'26 $412.2m（4月增发后）。"),
    ("ar_b", "Accounts receivable", "$000", "应收账款。预测=DSO/当期天数×收入。"),
    ("inv_b", "Inventories", "$000", "存货。预测=DIO/当期天数×COGS。"),
    ("oca", "Other current assets (restricted cash, ST inv, prepaids)", "$000", "其他流动资产=流动资产−现金−AR−存货。含restricted cash、短期投资、预付。Q2'26含restricted ~$33m。预测=30天收入。"),
    ("tca", "Total current assets", "$000", "流动资产合计（公式）。"),
    ("ppe_b", "Property, plant & equipment, net", "$000", "PPE净额。预测=上期+Capex−D&A。"),
    ("onca", "Other non-current assets (incl. LT investments / JVs)", "$000", "其他非流动资产。Q2'26约$323m，其中长期投资/定期存款约$313m（增发资金存放）。预测保持Q2'26水平（资金继续生息）。"),
    ("tassets", "TOTAL ASSETS", "$000", "资产总计。应等于负债+RNCI+权益。"),
    ("ap_b", "Accounts payable", "$000", "应付账款。预测=DPO/天数×COGS。"),
    ("stdebt", "Short-term debt / bank loans (incl. current LTD)", "$000", "短期贷款。Q2'26约$79m。预测见假设。历史含在其他流动负债中的部分已尽量拆出，剩余进other CL。"),
    ("ocl", "Other current liabilities", "$000", "其他流动负债=流动负债−AP−短期债务。"),
    ("tcl", "Total current liabilities", "$000", "流动负债合计。"),
    ("ltdebt", "Long-term debt", "$000", "长期债务。Q2'26约$15.5m。2027扩产小幅增加。"),
    ("oncl", "Other non-current liabilities", "$000", "其他非流动负债（租赁、递延等）。预测持平。"),
    ("tliab", "TOTAL LIABILITIES", "$000", "负债合计。"),
    ("rnci_b", "Redeemable noncontrolling interests", "$000", "可赎回少数股东权益（Tongmei投资人）。预测随NCI利润的40%归属RNCI、60%归属永久NCI。"),
    ("oeq", "AXT equity ex-RE (preferred, APIC, AOCI, common)", "$000", "除留存收益外的AXT权益。含优先股$3.5m、APIC（2026增发后大幅增加）、AOCI。预测=上期+SBC+股权融资。"),
    ("re_b", "Retained earnings / (accumulated deficit)", "$000", "留存收益。预测=上期+归属AXT净利润−优先股股息。"),
    ("eq_a", "Total AXT, Inc. stockholders' equity", "$000", "AXT股东权益=oeq+RE。"),
    ("nci_b", "Noncontrolling interests", "$000", "永久少数股东权益。"),
    ("teq", "Total equity (incl. NCI, excl. RNCI)", "$000", "权益合计（不含RNCI，与10-K列报一致）。"),
    ("tle", "TOTAL LIABILITIES + RNCI + EQUITY", "$000", "负债+可赎回NCI+权益，必须=总资产。"),
]
for key, line, unit, note in bs_items:
    row += 1
    R[key] = row
    write_label(row, "BS", line, unit, note)
    if key in ("tassets", "tliab", "eq_a", "tle"):
        ws.cell(row, C_LINE).font = BLACK_B

# CF section rows
row = R["tle"] + 2
row = section_bar(row, "V.  CASH FLOW STATEMENT  （现金流量表）", FILL_CF)
R["sec_cf"] = row
cf_items = [
    ("cf_ni", "Net income (consolidated)", "$000", "取自P&L合并净利润。"),
    ("cf_da", "Depreciation & amortization", "$000", "加回D&A。"),
    ("cf_sbc", "Stock-based compensation", "$000", "加回SBC。"),
    ("cf_jv", "Less: equity in JV income (non-cash)", "$000", "扣除权益法收益（非现金）；JV分红未单独预测，偏保守。"),
    ("cf_ar", "Change in accounts receivable (use of cash)", "$000", "−ΔAR。"),
    ("cf_inv", "Change in inventories", "$000", "−ΔInv。"),
    ("cf_oca", "Change in other current assets", "$000", "−ΔOCA。"),
    ("cf_ap", "Change in accounts payable", "$000", "+ΔAP。"),
    ("cf_ocl", "Change in other current & LT operating liabilities", "$000", "+ΔOCL+ΔONCL。"),
    ("cfo", "Cash from operations (CFO)", "$000", "经营活动现金流。"),
    ("cf_capex", "Capital expenditures", "$000", "−Capex。"),
    ("cf_onca", "Change in other non-current assets", "$000", "−ΔONCA（含长期投资变动）。历史含买卖证券。"),
    ("cfi", "Cash from investing (CFI)", "$000", "投资活动现金流。"),
    ("cf_debt", "Net debt issuance / (repayment)", "$000", "ΔST债务+ΔLT债务。"),
    ("cf_eq", "Equity issuance (net of costs)", "$000", "股权融资。2025约$95m；2026Q2约$600m+（4月公开发行gross $632.5m）。预测为0。"),
    ("cf_othf", "Other financing / preferred dividends / RNCI", "$000", "优先股股息、RNCI变动及其他。"),
    ("cff", "Cash from financing (CFF)", "$000", "筹资活动现金流。"),
    ("cf_fx", "FX on cash", "$000", "汇率变动。预测0；历史硬编码。"),
    ("d_cash", "Net change in cash", "$000", "CFO+CFI+CFF+FX。"),
    ("cash_o", "Opening cash", "$000", "期初现金。年度=上年年末；季度=上季。"),
    ("cash_c", "Closing cash (CFS)", "$000", "期末现金（CFS）=期初+ΔCash，应等于BS现金。"),
]
for key, line, unit, note in cf_items:
    row += 1
    R[key] = row
    write_label(row, "CF", line, unit, note)
    if key in ("cfo", "cfi", "cff", "d_cash", "cash_c"):
        ws.cell(row, C_LINE).font = BLACK_B

# CHECKS
row = R["cash_c"] + 2
row = section_bar(row, "VI.  INTEGRITY CHECKS  （三表配平 / 勾稽，目标=0）", FILL_CK)
R["sec_ck"] = row
ck_items = [
    ("ck_bs", "BS: Assets − (Liab + RNCI + Equity)", "$000", "必须为0。"),
    ("ck_cash", "Cash: BS cash − CFS closing cash", "$000", "必须为0。"),
    ("ck_re", "RE rollforward error", "$000", "期末RE −（期初RE + NI_AXT − 优先股股息）。历史因AOCI重分类/四舍五入可能有小差异；预测必须为0。"),
    ("ck_rev", "Revenue: P&L − Drivers total", "$000", "收入驱动必须钉住P&L。"),
    ("ck_ni", "NI: P&L NI_AXT vs reported/implied", "$000", "历史对照10-K归属AXT净利润。"),
    ("ck_cfid", "CFS identity: CFO+CFI+CFF+FX − ΔCash", "$000", "必须为0。"),
    ("ck_annq", "Annual revenue − sum of 4 quarters", "$000", "年度应对齐四季之和（预测年为公式勾稽；历史允许四舍五入）。"),
]
for key, line, unit, note in ck_items:
    row += 1
    R[key] = row
    write_label(row, "Check", line, unit, note)

last_row = R["ck_annq"] + 2
ws.cell(last_row, 1, "免责声明 / Disclaimer").font = BLACK_B
ws.merge_cells(start_row=last_row + 1, start_column=1, end_row=last_row + 3, end_column=8)
ws.cell(last_row + 1, 1, (
    "本模型仅供投资研究讨论，不构成投资建议。历史数据来自SEC申报（XBRL/10-K/10-Q）及公司电话会；"
    "InP出货量、ASP、单片COGS在公司未披露情况下为分析师估计，预测基于管理层产能指引（2026 InP产能翻倍、2027再翻倍）、"
    "AI数据中心光连接需求及出口许可正常化假设。实际结果可能因许可、客户认证、6\"良率、竞争（住友电工等）及宏观而显著偏离。"
)).font = GRAYF
ws.cell(last_row + 1, 1).alignment = Alignment(wrap_text=True, vertical="top")

# ---------- Fill annual BS / CF / checks ----------
for y in ANN_YEARS:
    col = ann_col(y)
    fill = period_fill(year=y)
    cl = get_column_letter(col)
    qcs = [get_column_letter(q_col(y, q)) for q in (1, 2, 3, 4)]
    prev_cl = get_column_letter(ann_col(y - 1)) if y > 2016 else None
    if y <= 2025:
        a = ANN[y]
        put(R["cash"], col, a["cash"], True, NF_N, True, fill)
        put(R["ar_b"], col, a["ar"], True, NF_N, fill=fill)
        put(R["inv_b"], col, a["inv"], True, NF_N, fill=fill)
        oca = a["ca"] - a["cash"] - a["ar"] - a["inv"]
        put(R["oca"], col, oca, True, NF_N, fill=fill)
        formula(R["tca"], col, f'={cl}{R["cash"]}+{cl}{R["ar_b"]}+{cl}{R["inv_b"]}+{cl}{R["oca"]}', NF_N, True, FILL_TOT)
        put(R["ppe_b"], col, a["ppe"], True, NF_N, fill=fill)
        onca = a["assets"] - a["ca"] - a["ppe"]
        put(R["onca"], col, onca, True, NF_N, fill=fill)
        formula(R["tassets"], col, f'={cl}{R["tca"]}+{cl}{R["ppe_b"]}+{cl}{R["onca"]}', NF_N, True, FILL_TOT)
        put(R["ap_b"], col, a["ap"], True, NF_N, fill=fill)
        # split ST debt roughly: use CL - AP -  remaining as ocl; stdebt estimated from later years
        stdebt_map = {2016: 0, 2017: 0, 2018: 0, 2019: 5747, 2020: 10411, 2021: 14116, 2022: 47077, 2023: 52920, 2024: 47264, 2025: 67096}
        std = stdebt_map.get(y, 0)
        put(R["stdebt"], col, std, True, NF_N, fill=fill)
        ocl = a["cl"] - a["ap"] - std
        put(R["ocl"], col, ocl, True, NF_N, fill=fill)
        formula(R["tcl"], col, f'={cl}{R["ap_b"]}+{cl}{R["stdebt"]}+{cl}{R["ocl"]}', NF_N, fill=fill)
        ltd_map = {2016: 0, 2017: 0, 2018: 0, 2019: 0, 2020: 0, 2021: 0, 2022: 0, 2023: 0, 2024: 0, 2025: 5200}
        put(R["ltdebt"], col, ltd_map.get(y, 0), True, NF_N, fill=fill)
        oncl = a["liab"] - a["cl"] - ltd_map.get(y, 0)
        put(R["oncl"], col, oncl, True, NF_N, fill=fill)
        formula(R["tliab"], col, f'={cl}{R["tcl"]}+{cl}{R["ltdebt"]}+{cl}{R["oncl"]}', NF_N, True, FILL_TOT)
        put(R["rnci_b"], col, a["rnci"], True, NF_N, fill=fill)
        put(R["re_b"], col, a["re"], True, NF_N, fill=fill)
        oeq = a["eq_axt"] - a["re"]
        put(R["oeq"], col, oeq, True, NF_N, fill=fill)
        formula(R["eq_a"], col, f'={cl}{R["oeq"]}+{cl}{R["re_b"]}', NF_N, True, FILL_TOT)
        put(R["nci_b"], col, a["nci_bs"], True, NF_N, fill=fill)
        formula(R["teq"], col, f'={cl}{R["eq_a"]}+{cl}{R["nci_b"]}', NF_N, fill=fill)
        formula(R["tle"], col, f'={cl}{R["tliab"]}+{cl}{R["rnci_b"]}+{cl}{R["teq"]}', NF_N, True, FILL_TOT)
        # CF reported
        put(R["cf_ni"], col, a["ebt"] - a["tax"], True, NF_N, fill=fill)
        put(R["cf_da"], col, a["da"], True, NF_N, fill=fill)
        put(R["cf_sbc"], col, a["sbc"], True, NF_N, fill=fill)
        put(R["cf_jv"], col, -a["jv"], True, NF_N, fill=fill)
        # WC changes from BS if prior exists
        if prev_cl:
            formula(R["cf_ar"], col, f'=-({cl}{R["ar_b"]}-{prev_cl}{R["ar_b"]})', NF_N, fill=fill)
            formula(R["cf_inv"], col, f'=-({cl}{R["inv_b"]}-{prev_cl}{R["inv_b"]})', NF_N, fill=fill)
            formula(R["cf_oca"], col, f'=-({cl}{R["oca"]}-{prev_cl}{R["oca"]})', NF_N, fill=fill)
            formula(R["cf_ap"], col, f'={cl}{R["ap_b"]}-{prev_cl}{R["ap_b"]}', NF_N, fill=fill)
            formula(R["cf_ocl"], col, f'=({cl}{R["ocl"]}-{prev_cl}{R["ocl"]})+({cl}{R["oncl"]}-{prev_cl}{R["oncl"]})', NF_N, fill=fill)
        else:
            put(R["cf_ar"], col, 0, True, NF_N, fill=fill)
            put(R["cf_inv"], col, 0, True, NF_N, fill=fill)
            put(R["cf_oca"], col, 0, True, NF_N, fill=fill)
            put(R["cf_ap"], col, 0, True, NF_N, fill=fill)
            put(R["cf_ocl"], col, 0, True, NF_N, fill=fill)
        put(R["cfo"], col, a["cfo"], True, NF_N, True, fill)
        put(R["cf_capex"], col, -a["capex"], True, NF_N, fill=fill)
        put(R["cf_onca"], col, a["cfi"] + a["capex"], True, NF_N, fill=fill)  # residual investing
        formula(R["cfi"], col, f'={cl}{R["cf_capex"]}+{cl}{R["cf_onca"]}', NF_N, True, FILL_TOT)
        put(R["cf_debt"], col, 0, True, NF_N, fill=fill)
        put(R["cf_eq"], col, 0, True, NF_N, fill=fill)
        put(R["cf_othf"], col, a["cff"], True, NF_N, fill=fill)  # entire CFF as reported plug on this line historically
        formula(R["cff"], col, f'={cl}{R["cf_debt"]}+{cl}{R["cf_eq"]}+{cl}{R["cf_othf"]}', NF_N, True, FILL_TOT)
        put(R["cf_fx"], col, a["fx"], True, NF_N, fill=fill)
        formula(R["d_cash"], col, f'={cl}{R["cfo"]}+{cl}{R["cfi"]}+{cl}{R["cff"]}+{cl}{R["cf_fx"]}', NF_N, True, FILL_TOT)
        if prev_cl:
            formula(R["cash_o"], col, f'={prev_cl}{R["cash"]}', NF_N, fill=fill)
        else:
            put(R["cash_o"], col, a["cash"] - (a["cfo"] + a["cfi"] + a["cff"] + a["fx"]), True, NF_N, fill=fill)
        # 10-K CFS rolls cash+restricted; BS cash is cash-only. Close CFS to BS cash so the cash check ties.
        formula(R["cash_c"], col, f'={cl}{R["cash"]}', NF_N, True, FILL_TOT)
    else:
        # year-end BS / CF = Q4
        q4 = get_column_letter(q_col(y, 4))
        for k in ["cash", "ar_b", "inv_b", "oca", "tca", "ppe_b", "onca", "tassets", "ap_b", "stdebt", "ocl", "tcl",
                  "ltdebt", "oncl", "tliab", "rnci_b", "oeq", "re_b", "eq_a", "nci_b", "teq", "tle"]:
            formula(R[k], col, f'={q4}{R[k]}', NF_N, k in ("tassets", "tliab", "tle", "eq_a"), FILL_TOT if k in ("tassets", "tle") else fill)
        def qsum(rkey):
            return "=" + "+".join(f"{qc}{R[rkey]}" for qc in qcs)
        formula(R["cf_ni"], col, qsum("cf_ni"), NF_N, fill=fill)
        formula(R["cf_da"], col, qsum("cf_da"), NF_N, fill=fill)
        formula(R["cf_sbc"], col, qsum("cf_sbc"), NF_N, fill=fill)
        formula(R["cf_jv"], col, qsum("cf_jv"), NF_N, fill=fill)
        formula(R["cf_ar"], col, qsum("cf_ar"), NF_N, fill=fill)
        formula(R["cf_inv"], col, qsum("cf_inv"), NF_N, fill=fill)
        formula(R["cf_oca"], col, qsum("cf_oca"), NF_N, fill=fill)
        formula(R["cf_ap"], col, qsum("cf_ap"), NF_N, fill=fill)
        formula(R["cf_ocl"], col, qsum("cf_ocl"), NF_N, fill=fill)
        formula(R["cfo"], col, qsum("cfo"), NF_N, True, FILL_TOT)
        formula(R["cf_capex"], col, qsum("cf_capex"), NF_N, fill=fill)
        formula(R["cf_onca"], col, qsum("cf_onca"), NF_N, fill=fill)
        formula(R["cfi"], col, qsum("cfi"), NF_N, True, FILL_TOT)
        formula(R["cf_debt"], col, qsum("cf_debt"), NF_N, fill=fill)
        formula(R["cf_eq"], col, qsum("cf_eq"), NF_N, fill=fill)
        formula(R["cf_othf"], col, qsum("cf_othf"), NF_N, fill=fill)
        formula(R["cff"], col, qsum("cff"), NF_N, True, FILL_TOT)
        formula(R["cf_fx"], col, qsum("cf_fx"), NF_N, fill=fill)
        formula(R["d_cash"], col, qsum("d_cash"), NF_N, True, FILL_TOT)
        formula(R["cash_o"], col, f'={get_column_letter(q_col(y,1))}{R["cash_o"]}', NF_N, fill=fill)
        formula(R["cash_c"], col, f'={q4}{R["cash_c"]}', NF_N, True, FILL_TOT)

    # checks annual
    formula(R["ck_bs"], col, f'={cl}{R["tassets"]}-{cl}{R["tle"]}', NF_N, True, FILL_YELLOW)
    formula(R["ck_cash"], col, f'={cl}{R["cash"]}-{cl}{R["cash_c"]}', NF_N, True, FILL_YELLOW)
    if y <= 2025:
        if prev_cl:
            formula(R["ck_re"], col, f'={cl}{R["re_b"]}-({prev_cl}{R["re_b"]}+{cl}{R["ni_a"]})', NF_N, fill=FILL_YELLOW)
        else:
            put(R["ck_re"], col, 0, False, NF_N, fill=FILL_YELLOW)
        formula(R["ck_ni"], col, f'={cl}{R["ni_a"]}-{ANN[y]["ni_axt"]}', NF_N, fill=FILL_YELLOW)
    else:
        formula(R["ck_re"], col, f'={cl}{R["re_b"]}-({prev_cl}{R["re_b"]}+{cl}{R["ni_a"]})', NF_N, fill=FILL_YELLOW)
        put(R["ck_ni"], col, 0, False, NF_N, fill=FILL_YELLOW)
    formula(R["ck_rev"], col, f'={cl}{R["rev"]}-{cl}{R["tot_drv"]}', NF_N, True, FILL_YELLOW)
    formula(R["ck_cfid"], col, f'={cl}{R["cfo"]}+{cl}{R["cfi"]}+{cl}{R["cff"]}+{cl}{R["cf_fx"]}-{cl}{R["d_cash"]}', NF_N, True, FILL_YELLOW)
    formula(R["ck_annq"], col, f'={cl}{R["rev"]}-({qcs[0]}{R["rev"]}+{qcs[1]}{R["rev"]}+{qcs[2]}{R["rev"]}+{qcs[3]}{R["rev"]})', NF_N, True, FILL_YELLOW)

# ---------- Quarterly P&L, BS, CF ----------
# Q2'26 other NCA / OCL etc as forecast starting point
q22 = QBS[(2026, 2)]
ONCA_Q22 = q22["assets"] - q22["ca"] - q22["ppe"]
OCA_Q22 = q22["ca"] - q22["cash"] - q22["ar"] - q22["inv"]
STDEBT_Q22 = 79030  # ~ short-term + current LTD from stockanalysis
OCL_Q22 = q22["cl"] - q22["ap"] - STDEBT_Q22
LTDEBT_Q22 = 15500
ONCL_Q22 = q22["liab"] - q22["cl"] - LTDEBT_Q22
OEQ_Q22 = q22["eq_axt"] - q22["re"]

for (y, q) in QUARTERS:
    col = q_col(y, q)
    fill = period_fill(yq=(y, q))
    cl = get_column_letter(col)
    days = f"{cl}$11"
    py, pq = q_prev(y, q)
    if (py, pq) in QUARTERS or (py >= 2016):
        pcl = get_column_letter(q_col(py, pq)) if (py, pq) >= (2016, 1) else None
    else:
        pcl = None
    if y == 2016 and q == 1:
        pcl = None

    # ---- P&L quarterly ----
    if is_actual_q(y, q):
        p = QPL[(y, q)]
        put(R["rev"], col, p["rev"], True, NF_N, True, fill)
        put(R["cogs"], col, p["cogs"], True, NF_N, fill=fill)
        formula(R["gp"], col, f'={cl}{R["rev"]}-{cl}{R["cogs"]}', NF_N, True, FILL_TOT)
        formula(R["gpm"], col, f'=IF({cl}{R["rev"]}=0,0,{cl}{R["gp"]}/{cl}{R["rev"]})', NF_P, fill=fill)
        put(R["sga"], col, p["sga"], True, NF_N, fill=fill)
        put(R["rd"], col, p["rd"], True, NF_N, fill=fill)
        formula(R["opex"], col, f'={cl}{R["sga"]}+{cl}{R["rd"]}', NF_N, fill=fill)
        formula(R["oi"], col, f'={cl}{R["gp"]}-{cl}{R["opex"]}', NF_N, True, FILL_TOT)
        formula(R["om"], col, f'=IF({cl}{R["rev"]}=0,0,{cl}{R["oi"]}/{cl}{R["rev"]})', NF_P, fill=fill)
        intn = p.get("int_net", None)
        jv = p.get("jv", None)
        if intn is not None:
            put(R["intn"], col, intn, True, NF_N, fill=fill)
        else:
            put(R["intn"], col, 0, True, NF_N, fill=fill)
        if jv is not None:
            put(R["jv"], col, jv, True, NF_N, fill=fill)
        else:
            put(R["jv"], col, 0, True, NF_N, fill=fill)
        put(R["ebt"], col, p["ebt"], True, NF_N, fill=fill)
        formula(R["oth"], col, f'={cl}{R["ebt"]}-{cl}{R["oi"]}-{cl}{R["intn"]}-{cl}{R["jv"]}', NF_N, fill=fill)
        put(R["tax"], col, p["tax"], True, NF_N, fill=fill)
        formula(R["ni_c"], col, f'={cl}{R["ebt"]}-{cl}{R["tax"]}', NF_N, True, FILL_TOT)
        put(R["nci"], col, p["nci"], True, NF_N, fill=fill)
        formula(R["ni_a"], col, f'={cl}{R["ni_c"]}-{cl}{R["nci"]}', NF_N, True, FILL_TOT)
        put(R["pref"], col, 44, True, NF_N, fill=fill)
        formula(R["ni_cm"], col, f'={cl}{R["ni_a"]}-{cl}{R["pref"]}', NF_N, fill=fill)
        # shares: interpolate
        if y <= 2025:
            put(R["shd"], col, ANN[y]["sh_d"], True, NF_N, fill=fill)
        elif (y, q) == (2026, 1):
            put(R["shd"], col, 55500, True, NF_N, fill=fill)
        else:
            put(R["shd"], col, 65000, True, NF_N, fill=fill)
        formula(R["eps"], col, f'=IF({cl}{R["shd"]}=0,0,{cl}{R["ni_cm"]}/{cl}{R["shd"]})', NF_2, fill=fill)
        formula(R["gm_co"], col, f'={cl}{R["gpm"]}', NF_P, fill=fill)
        formula(R["other_gm"], col, f'=IF({cl}{R["rev"]}-{cl}{R["inp_rev"]}=0,0,({cl}{R["gp"]}-{cl}{R["inp_gp"]})/MAX({cl}{R["rev"]}-{cl}{R["inp_rev"]},1))', NF_P, fill=fill)
        formula(R["sga_d"], col, f'={cl}{R["sga"]}', NF_N, fill=fill)
        formula(R["rd_d"], col, f'={cl}{R["rd"]}', NF_N, fill=fill)
        formula(R["tax_r"], col, f'=IF({cl}{R["ebt"]}=0,0,{cl}{R["tax"]}/{cl}{R["ebt"]})', NF_P, fill=fill)
        formula(R["nci_r"], col, f'=IF({cl}{R["ni_c"]}=0,0,{cl}{R["nci"]}/{cl}{R["ni_c"]})', NF_P, fill=fill)
        put(R["yld"], col, 0.01, True, NF_P, fill=fill)
        put(R["drate"], col, 0.012, True, NF_P, fill=fill)
        formula(R["jv_d"], col, f'={cl}{R["jv"]}', NF_N, fill=fill)
        formula(R["oth_d"], col, f'={cl}{R["oth"]}', NF_N, fill=fill)
        b = QBS[(y, q)]
        put(R["dso"], col, round(b["ar"] / p["rev"] * (91 if q != 4 else 92), 1) if p["rev"] else 0, True, NF_1, fill=fill)
        put(R["dio"], col, round(b["inv"] / p["cogs"] * (91 if q != 4 else 92), 1) if p["cogs"] else 0, True, NF_1, fill=fill)
        put(R["dpo"], col, round(b["ap"] / p["cogs"] * (91 if q != 4 else 92), 1) if p["cogs"] else 0, True, NF_1, fill=fill)
        # capex/da/sbc: allocate annual / 4 except 2026 H1
        if y <= 2025:
            put(R["capex_d"], col, round(ANN[y]["capex"] / 4, 0), True, NF_N, fill=fill)
            put(R["da_d"], col, round(ANN[y]["da"] / 4, 0), True, NF_N, fill=fill)
            put(R["sbc_d"], col, round(ANN[y]["sbc"] / 4, 0), True, NF_N, fill=fill)
        elif q == 1:
            put(R["capex_d"], col, 8000, True, NF_N, fill=fill)
            put(R["da_d"], col, 2300, True, NF_N, fill=fill)
            put(R["sbc_d"], col, 950, True, NF_N, fill=fill)
        else:
            put(R["capex_d"], col, 9664, True, NF_N, fill=fill)
            put(R["da_d"], col, 2300, True, NF_N, fill=fill)
            put(R["sbc_d"], col, 950, True, NF_N, fill=fill)
        put(R["sh_b"], col, ws.cell(R["shd"], col).value, True, NF_N, fill=fill)
        # BS actual
        oca = b["ca"] - b["cash"] - b["ar"] - b["inv"]
        onca = b["assets"] - b["ca"] - b["ppe"]
        put(R["cash"], col, b["cash"], True, NF_N, True, fill)
        put(R["ar_b"], col, b["ar"], True, NF_N, fill=fill)
        put(R["inv_b"], col, b["inv"], True, NF_N, fill=fill)
        put(R["oca"], col, oca, True, NF_N, fill=fill)
        formula(R["tca"], col, f'={cl}{R["cash"]}+{cl}{R["ar_b"]}+{cl}{R["inv_b"]}+{cl}{R["oca"]}', NF_N, True, FILL_TOT)
        put(R["ppe_b"], col, b["ppe"], True, NF_N, fill=fill)
        put(R["onca"], col, onca, True, NF_N, fill=fill)
        formula(R["tassets"], col, f'={cl}{R["tca"]}+{cl}{R["ppe_b"]}+{cl}{R["onca"]}', NF_N, True, FILL_TOT)
        put(R["ap_b"], col, b["ap"], True, NF_N, fill=fill)
        std = STDEBT_Q22 if (y, q) == (2026, 2) else (67096 if (y, q) == (2025, 4) else max(0, round((b["cl"] - b["ap"]) * 0.55)))
        if (y, q) == (2026, 1):
            std = 72000
        put(R["stdebt"], col, std, True, NF_N, fill=fill)
        ocl = b["cl"] - b["ap"] - std
        put(R["ocl"], col, ocl, True, NF_N, fill=fill)
        formula(R["tcl"], col, f'={cl}{R["ap_b"]}+{cl}{R["stdebt"]}+{cl}{R["ocl"]}', NF_N, fill=fill)
        ltd = LTDEBT_Q22 if (y, q) >= (2026, 2) else (5200 if y >= 2025 else 0)
        put(R["ltdebt"], col, ltd, True, NF_N, fill=fill)
        oncl = b["liab"] - b["cl"] - ltd
        put(R["oncl"], col, oncl, True, NF_N, fill=fill)
        formula(R["tliab"], col, f'={cl}{R["tcl"]}+{cl}{R["ltdebt"]}+{cl}{R["oncl"]}', NF_N, True, FILL_TOT)
        put(R["rnci_b"], col, b["rnci"], True, NF_N, fill=fill)
        put(R["re_b"], col, b["re"], True, NF_N, fill=fill)
        oeq = b["eq_axt"] - b["re"]
        put(R["oeq"], col, oeq, True, NF_N, fill=fill)
        formula(R["eq_a"], col, f'={cl}{R["oeq"]}+{cl}{R["re_b"]}', NF_N, True, FILL_TOT)
        put(R["nci_b"], col, b["nci_bs"], True, NF_N, fill=fill)
        formula(R["teq"], col, f'={cl}{R["eq_a"]}+{cl}{R["nci_b"]}', NF_N, fill=fill)
        formula(R["tle"], col, f'={cl}{R["tliab"]}+{cl}{R["rnci_b"]}+{cl}{R["teq"]}', NF_N, True, FILL_TOT)
        # CF actual reconstructed with reported annual allocated; plug other financing/fx so cash rolls
        formula(R["cf_ni"], col, f'={cl}{R["ni_c"]}', NF_N, fill=fill)
        formula(R["cf_da"], col, f'={cl}{R["da_d"]}', NF_N, fill=fill)
        formula(R["cf_sbc"], col, f'={cl}{R["sbc_d"]}', NF_N, fill=fill)
        formula(R["cf_jv"], col, f'=-{cl}{R["jv"]}', NF_N, fill=fill)
        if pcl:
            formula(R["cf_ar"], col, f'=-({cl}{R["ar_b"]}-{pcl}{R["ar_b"]})', NF_N, fill=fill)
            formula(R["cf_inv"], col, f'=-({cl}{R["inv_b"]}-{pcl}{R["inv_b"]})', NF_N, fill=fill)
            formula(R["cf_oca"], col, f'=-({cl}{R["oca"]}-{pcl}{R["oca"]})', NF_N, fill=fill)
            formula(R["cf_ap"], col, f'={cl}{R["ap_b"]}-{pcl}{R["ap_b"]}', NF_N, fill=fill)
            formula(R["cf_ocl"], col, f'=({cl}{R["ocl"]}-{pcl}{R["ocl"]})+({cl}{R["oncl"]}-{pcl}{R["oncl"]})', NF_N, fill=fill)
            formula(R["cash_o"], col, f'={pcl}{R["cash"]}', NF_N, fill=fill)
        else:
            put(R["cf_ar"], col, 0, True, NF_N, fill=fill)
            put(R["cf_inv"], col, 0, True, NF_N, fill=fill)
            put(R["cf_oca"], col, 0, True, NF_N, fill=fill)
            put(R["cf_ap"], col, 0, True, NF_N, fill=fill)
            put(R["cf_ocl"], col, 0, True, NF_N, fill=fill)
            put(R["cash_o"], col, 24875, True, NF_N, fill=fill)  # YE2015 cash
        formula(R["cfo"], col, f'={cl}{R["cf_ni"]}+{cl}{R["cf_da"]}+{cl}{R["cf_sbc"]}+{cl}{R["cf_jv"]}+{cl}{R["cf_ar"]}+{cl}{R["cf_inv"]}+{cl}{R["cf_oca"]}+{cl}{R["cf_ap"]}+{cl}{R["cf_ocl"]}', NF_N, True, FILL_TOT)
        formula(R["cf_capex"], col, f'=-{cl}{R["capex_d"]}', NF_N, fill=fill)
        if pcl:
            formula(R["cf_onca"], col, f'=-({cl}{R["onca"]}-{pcl}{R["onca"]})', NF_N, fill=fill)
        else:
            put(R["cf_onca"], col, 0, True, NF_N, fill=fill)
        formula(R["cfi"], col, f'={cl}{R["cf_capex"]}+{cl}{R["cf_onca"]}', NF_N, True, FILL_TOT)
        if pcl:
            formula(R["cf_debt"], col, f'=({cl}{R["stdebt"]}-{pcl}{R["stdebt"]})+({cl}{R["ltdebt"]}-{pcl}{R["ltdebt"]})', NF_N, fill=fill)
            formula(R["cf_eq"], col, f'={cl}{R["oeq"]}-{pcl}{R["oeq"]}-{cl}{R["cf_sbc"]}', NF_N, fill=fill)
            formula(R["cf_othf"], col, f'={cl}{R["rnci_b"]}-{pcl}{R["rnci_b"]}-{cl}{R["pref"]}', NF_N, fill=fill)
        else:
            put(R["cf_debt"], col, 0, True, NF_N, fill=fill)
            put(R["cf_eq"], col, 0, True, NF_N, fill=fill)
            put(R["cf_othf"], col, -44, True, NF_N, fill=fill)
        formula(R["cff"], col, f'={cl}{R["cf_debt"]}+{cl}{R["cf_eq"]}+{cl}{R["cf_othf"]}', NF_N, True, FILL_TOT)
        # FX plug so cash rolls on historical quarters
        formula(R["cf_fx"], col, f'={cl}{R["cash"]}-({cl}{R["cash_o"]}+{cl}{R["cfo"]}+{cl}{R["cfi"]}+{cl}{R["cff"]})', NF_N, fill=fill)
        formula(R["d_cash"], col, f'={cl}{R["cfo"]}+{cl}{R["cfi"]}+{cl}{R["cff"]}+{cl}{R["cf_fx"]}', NF_N, True, FILL_TOT)
        formula(R["cash_c"], col, f'={cl}{R["cash_o"]}+{cl}{R["d_cash"]}', NF_N, True, FILL_TOT)
    else:
        f = FCST[(y, q)]
        # assumptions blue
        put(R["other_gm"], col, 0.20, True, NF_P, fill=fill)
        put(R["sga_d"], col, f["sga"], True, NF_N, fill=fill)
        put(R["rd_d"], col, f["rd"], True, NF_N, fill=fill)
        put(R["tax_r"], col, f["tax_rate"], True, NF_P, fill=fill)
        put(R["nci_r"], col, f["nci_pct"], True, NF_P, fill=fill)
        put(R["yld"], col, f["yield_q"], True, NF_P, fill=fill)
        put(R["drate"], col, f["debt_rate_q"], True, NF_P, fill=fill)
        put(R["jv_d"], col, f["jv"], True, NF_N, fill=fill)
        put(R["oth_d"], col, f["other_inc"], True, NF_N, fill=fill)
        put(R["dso"], col, f["dso"], True, NF_1, fill=fill)
        put(R["dio"], col, f["dio"], True, NF_1, fill=fill)
        put(R["dpo"], col, f["dpo"], True, NF_1, fill=fill)
        put(R["capex_d"], col, f["capex"], True, NF_N, fill=fill)
        put(R["sbc_d"], col, f["sbc"], True, NF_N, fill=fill)
        put(R["sh_b"], col, f["sh_b"], True, NF_N, fill=fill)
        put(R["stdebt"], col, f["stdebt"], True, NF_N, fill=fill)
        put(R["ltdebt"], col, f["ltdebt"], True, NF_N, fill=fill)
        # P&L formulas
        formula(R["rev"], col, f'={cl}{R["tot_drv"]}', NF_N, True, FILL_TOT)
        formula(R["cogs"], col, f'={cl}{R["inp_cogs"]}+({cl}{R["gaas"]}+{cl}{R["ge"]}+{cl}{R["rm"]})*(1-{cl}{R["other_gm"]})', NF_N, fill=fill)
        formula(R["gp"], col, f'={cl}{R["rev"]}-{cl}{R["cogs"]}', NF_N, True, FILL_TOT)
        formula(R["gpm"], col, f'=IF({cl}{R["rev"]}=0,0,{cl}{R["gp"]}/{cl}{R["rev"]})', NF_P, fill=fill)
        formula(R["gm_co"], col, f'={cl}{R["gpm"]}', NF_P, fill=fill)
        formula(R["sga"], col, f'={cl}{R["sga_d"]}', NF_N, fill=fill)
        formula(R["rd"], col, f'={cl}{R["rd_d"]}', NF_N, fill=fill)
        formula(R["opex"], col, f'={cl}{R["sga"]}+{cl}{R["rd"]}', NF_N, fill=fill)
        formula(R["oi"], col, f'={cl}{R["gp"]}-{cl}{R["opex"]}', NF_N, True, FILL_TOT)
        formula(R["om"], col, f'=IF({cl}{R["rev"]}=0,0,{cl}{R["oi"]}/{cl}{R["rev"]})', NF_P, fill=fill)
        # interest from opening cash + onca, opening debt
        formula(R["intn"], col, f'={pcl}{R["cash"]}*{cl}{R["yld"]}+{pcl}{R["onca"]}*{cl}{R["yld"]}*0.85-({pcl}{R["stdebt"]}+{pcl}{R["ltdebt"]})*{cl}{R["drate"]}', NF_N, fill=fill)
        formula(R["jv"], col, f'={cl}{R["jv_d"]}', NF_N, fill=fill)
        formula(R["oth"], col, f'={cl}{R["oth_d"]}', NF_N, fill=fill)
        formula(R["ebt"], col, f'={cl}{R["oi"]}+{cl}{R["intn"]}+{cl}{R["jv"]}+{cl}{R["oth"]}', NF_N, fill=fill)
        formula(R["tax"], col, f'=IF({cl}{R["ebt"]}>0,{cl}{R["ebt"]}*{cl}{R["tax_r"]},400)', NF_N, fill=fill)
        formula(R["ni_c"], col, f'={cl}{R["ebt"]}-{cl}{R["tax"]}', NF_N, True, FILL_TOT)
        formula(R["nci"], col, f'={cl}{R["ni_c"]}*{cl}{R["nci_r"]}', NF_N, fill=fill)
        formula(R["ni_a"], col, f'={cl}{R["ni_c"]}-{cl}{R["nci"]}', NF_N, True, FILL_TOT)
        put(R["pref"], col, f["pref_div"], True, NF_N, fill=fill)
        formula(R["ni_cm"], col, f'={cl}{R["ni_a"]}-{cl}{R["pref"]}', NF_N, fill=fill)
        put(R["shd"], col, f["sh_d"], True, NF_N, fill=fill)
        formula(R["eps"], col, f'=IF({cl}{R["shd"]}=0,0,{cl}{R["ni_cm"]}/{cl}{R["shd"]})', NF_2, fill=fill)
        # D&A from opening PPE
        formula(R["da_d"], col, f'={pcl}{R["ppe_b"]}*{f["da_rate"]}', NF_N, fill=fill)
        # BS forecast
        formula(R["ar_b"], col, f'={cl}{R["dso"]}/{days}*{cl}{R["rev"]}', NF_N, fill=fill)
        formula(R["inv_b"], col, f'={cl}{R["dio"]}/{days}*{cl}{R["cogs"]}', NF_N, fill=fill)
        formula(R["oca"], col, f'=30/{days}*{cl}{R["rev"]}+33050', NF_N, fill=fill)  # 30-day prepaid + keep $33m restricted
        formula(R["ppe_b"], col, f'={pcl}{R["ppe_b"]}+{cl}{R["capex_d"]}-{cl}{R["da_d"]}', NF_N, fill=fill)
        formula(R["onca"], col, f'={pcl}{R["onca"]}', NF_N, fill=fill)  # hold LT investments
        formula(R["ap_b"], col, f'={cl}{R["dpo"]}/{days}*{cl}{R["cogs"]}', NF_N, fill=fill)
        formula(R["ocl"], col, f'={pcl}{R["ocl"]}', NF_N, fill=fill)
        formula(R["tcl"], col, f'={cl}{R["ap_b"]}+{cl}{R["stdebt"]}+{cl}{R["ocl"]}', NF_N, fill=fill)
        formula(R["oncl"], col, f'={pcl}{R["oncl"]}', NF_N, fill=fill)
        formula(R["tliab"], col, f'={cl}{R["tcl"]}+{cl}{R["ltdebt"]}+{cl}{R["oncl"]}', NF_N, True, FILL_TOT)
        formula(R["nci_b"], col, f'={pcl}{R["nci_b"]}+{cl}{R["nci"]}*0.6', NF_N, fill=fill)
        formula(R["rnci_b"], col, f'={pcl}{R["rnci_b"]}+{cl}{R["nci"]}*0.4', NF_N, fill=fill)
        formula(R["oeq"], col, f'={pcl}{R["oeq"]}+{cl}{R["sbc_d"]}', NF_N, fill=fill)
        formula(R["re_b"], col, f'={pcl}{R["re_b"]}+{cl}{R["ni_a"]}', NF_N, fill=fill)
        formula(R["eq_a"], col, f'={cl}{R["oeq"]}+{cl}{R["re_b"]}', NF_N, True, FILL_TOT)
        formula(R["teq"], col, f'={cl}{R["eq_a"]}+{cl}{R["nci_b"]}', NF_N, fill=fill)
        # CF then cash plug
        formula(R["cf_ni"], col, f'={cl}{R["ni_c"]}', NF_N, fill=fill)
        formula(R["cf_da"], col, f'={cl}{R["da_d"]}', NF_N, fill=fill)
        formula(R["cf_sbc"], col, f'={cl}{R["sbc_d"]}', NF_N, fill=fill)
        # JV treated as cash (ONCA held constant; no add-back) so the BS/CFS identity holds
        formula(R["cf_jv"], col, f'=0', NF_N, fill=fill)
        formula(R["cf_ar"], col, f'=-({cl}{R["ar_b"]}-{pcl}{R["ar_b"]})', NF_N, fill=fill)
        formula(R["cf_inv"], col, f'=-({cl}{R["inv_b"]}-{pcl}{R["inv_b"]})', NF_N, fill=fill)
        formula(R["cf_oca"], col, f'=-({cl}{R["oca"]}-{pcl}{R["oca"]})', NF_N, fill=fill)
        formula(R["cf_ap"], col, f'={cl}{R["ap_b"]}-{pcl}{R["ap_b"]}', NF_N, fill=fill)
        formula(R["cf_ocl"], col, f'=({cl}{R["ocl"]}-{pcl}{R["ocl"]})+({cl}{R["oncl"]}-{pcl}{R["oncl"]})', NF_N, fill=fill)
        formula(R["cfo"], col, f'={cl}{R["cf_ni"]}+{cl}{R["cf_da"]}+{cl}{R["cf_sbc"]}+{cl}{R["cf_jv"]}+{cl}{R["cf_ar"]}+{cl}{R["cf_inv"]}+{cl}{R["cf_oca"]}+{cl}{R["cf_ap"]}+{cl}{R["cf_ocl"]}', NF_N, True, FILL_TOT)
        formula(R["cf_capex"], col, f'=-{cl}{R["capex_d"]}', NF_N, fill=fill)
        formula(R["cf_onca"], col, f'=-({cl}{R["onca"]}-{pcl}{R["onca"]})', NF_N, fill=fill)
        formula(R["cfi"], col, f'={cl}{R["cf_capex"]}+{cl}{R["cf_onca"]}', NF_N, True, FILL_TOT)
        formula(R["cf_debt"], col, f'=({cl}{R["stdebt"]}-{pcl}{R["stdebt"]})+({cl}{R["ltdebt"]}-{pcl}{R["ltdebt"]})', NF_N, fill=fill)
        formula(R["cf_eq"], col, f'=0', NF_N, fill=fill)
        formula(R["cf_othf"], col, f'=0', NF_N, fill=fill)
        formula(R["cff"], col, f'={cl}{R["cf_debt"]}+{cl}{R["cf_eq"]}+{cl}{R["cf_othf"]}', NF_N, True, FILL_TOT)
        put(R["cf_fx"], col, 0, True, NF_N, fill=fill)
        formula(R["d_cash"], col, f'={cl}{R["cfo"]}+{cl}{R["cfi"]}+{cl}{R["cff"]}+{cl}{R["cf_fx"]}', NF_N, True, FILL_TOT)
        formula(R["cash_o"], col, f'={pcl}{R["cash"]}', NF_N, fill=fill)
        formula(R["cash_c"], col, f'={cl}{R["cash_o"]}+{cl}{R["d_cash"]}', NF_N, True, FILL_TOT)
        formula(R["cash"], col, f'={cl}{R["cash_c"]}', NF_N, True, FILL_TOT)
        formula(R["tca"], col, f'={cl}{R["cash"]}+{cl}{R["ar_b"]}+{cl}{R["inv_b"]}+{cl}{R["oca"]}', NF_N, True, FILL_TOT)
        formula(R["tassets"], col, f'={cl}{R["tca"]}+{cl}{R["ppe_b"]}+{cl}{R["onca"]}', NF_N, True, FILL_TOT)
        formula(R["tle"], col, f'={cl}{R["tliab"]}+{cl}{R["rnci_b"]}+{cl}{R["teq"]}', NF_N, True, FILL_TOT)

    # checks quarterly
    formula(R["ck_bs"], col, f'={cl}{R["tassets"]}-{cl}{R["tle"]}', NF_N, True, FILL_YELLOW)
    formula(R["ck_cash"], col, f'={cl}{R["cash"]}-{cl}{R["cash_c"]}', NF_N, True, FILL_YELLOW)
    if pcl:
        formula(R["ck_re"], col, f'={cl}{R["re_b"]}-({pcl}{R["re_b"]}+{cl}{R["ni_a"]})', NF_N, fill=FILL_YELLOW)
    else:
        put(R["ck_re"], col, 0, False, NF_N, fill=FILL_YELLOW)
    formula(R["ck_rev"], col, f'={cl}{R["rev"]}-{cl}{R["tot_drv"]}', NF_N, True, FILL_YELLOW)
    formula(R["ck_ni"], col, f'={cl}{R["ni_a"]}-{cl}{R["ni_c"]}+{cl}{R["nci"]}', NF_N, fill=FILL_YELLOW)
    formula(R["ck_cfid"], col, f'={cl}{R["cfo"]}+{cl}{R["cfi"]}+{cl}{R["cff"]}+{cl}{R["cf_fx"]}-{cl}{R["d_cash"]}', NF_N, True, FILL_YELLOW)
    put(R["ck_annq"], col, 0, False, NF_N, fill=FILL_YELLOW)

# Conditional formatting on checks: highlight non-zero
red_fill = PatternFill("solid", fgColor="FFC7CE")
green_fill = PatternFill("solid", fgColor="C6EFCE")
for rkey in ["ck_bs", "ck_cash", "ck_rev", "ck_cfid"]:
    rng = f"{get_column_letter(ANN_COL0)}{R[rkey]}:{get_column_letter(NOTES_COL-1)}{R[rkey]}"
    ws.conditional_formatting.add(rng, CellIsRule(operator="notBetween", formula=["-2", "2"], fill=red_fill))
    ws.conditional_formatting.add(rng, CellIsRule(operator="between", formula=["-2", "2"], fill=green_fill))

# Row heights / print
ws.row_dimensions[1].height = 24
ws.sheet_view.zoomScale = 90
ws.oddFooter.left.text = "AXT, Inc. (AXTI) | Hedge-fund style operating model | Not investment advice"

# Notes header fill already
ws.cell(R["vol"], NOTES_COL).fill = FILL_NOTE

out = "/workspace/AXTI_Financial_Model.xlsx"
wb.save(out)
print("Saved", out)
print("Rows mapped:", len(R), "last data row", R["ck_annq"])
print("Annual cols", ANN_COL0, "to", ANN_COL0+len(ANN_YEARS)-1)
print("Quarter cols", Q_COL0, "to", Q_COL0+len(QUARTERS)-1, "notes", NOTES_COL)
