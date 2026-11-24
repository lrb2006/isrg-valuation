"""Build the Intuitive Surgical (ISRG) valuation workbook with live Excel formulas.

Sources: SEC XBRL company facts (financial statements), Intuitive's earnings releases
(procedures, installed base, placements, revenue mix), and a market price quote.
Run: python scripts/build_model.py  ->  ISRG_Valuation_Model.xlsx
"""
from pathlib import Path
import json

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

OUT = Path(__file__).resolve().parents[1] / "ISRG_Valuation_Model.xlsx"
YEARS = list(range(2026, 2036))           # forecast years, columns C..L
BLUE = Font(color="0000CC")               # inputs (banking convention)
BOLD = Font(bold=True)
HEAD = Font(bold=True, color="FFFFFF")
HEAD_FILL = PatternFill("solid", fgColor="2F4F4F")
SECTION = Font(bold=True, size=12)
THIN = Border(top=Side(style="thin"))
M = '#,##0;(#,##0)'
M1 = '#,##0.0;(#,##0.0)'
PCT = '0.0%'
USD = '$#,##0'

# ---- facts --------------------------------------------------------------------
HIST_YEARS = [2020, 2021, 2022, 2023, 2024, 2025]
HIST = {  # $ millions, from SEC XBRL company facts (10-K)
    "Revenue": [4358.4, 5710.1, 6222.2, 7124.1, 8352.1, 10064.7],
    "Cost of revenue": [1497.2, 1751.6, 2026.2, 2394.6, 2717.9, 3422.4],
    "Gross profit": [2861.2, 3958.5, 4196.0, 4729.5, 5634.2, 6642.3],
    "R&D": [595.1, 671.0, 879.0, 998.8, 1145.3, 1311.8],
    "SG&A": [1216.3, 1466.5, 1739.9, 1963.9, 2140.0, 2385.0],
    "Operating income": [1049.8, 1821.0, 1577.1, 1766.8, 2348.9, 2945.5],
    "Pre-tax income": [1207.0, 1890.3, 1606.8, 1958.9, 2673.8, 3311.4],
    "Income tax": [140.2, 162.2, 262.4, 141.6, 336.3, 434.8],
    "Net income attributable to Intuitive": [1060.6, 1704.6, 1322.3, 1798.0, 2322.6, 2856.0],
    "Depreciation": [221.0, 280.0, 326.0, 382.0, 439.0, 600.0],
    "Amortization of intangibles": [49.8, 27.4, 27.8, 20.2, 16.7, 13.2],
    "Stock-based compensation": [395.4, 449.2, 513.2, 592.8, 676.8, 788.2],
    "Operating cash flow": [1484.8, 2089.4, 1490.8, 1813.8, 2415.0, 3030.5],
    "Capital expenditures": [341.5, 339.5, 532.4, 1064.2, 1111.2, 539.8],
    "Share repurchases": [134.3, 0.0, 2607.4, 416.3, 0.0, 2295.3],
    "Diluted shares (M)": [361.0, 365.8, 362.0, 357.4, 362.0, 362.7],
}
NINE_MONTHS = {  # 9 months to Sep 30 ($M), SEC XBRL quarterly frames
    "Revenue": (7198.5, 8711.8),
    "Cost of revenue": (2460.5, 2836.6),
    "R&D": (958.9, 1140.8),
    "SG&A": (1697.9, 1864.6),
    "Operating income": (2081.2, 2419.8),
}
KPI = {  # Intuitive earnings releases / 10-K
    "da Vinci procedures": (2_683_000, 3_153_000),
    "da Vinci installed base (year end)": (9_902, 11_106),
    "da Vinci systems placed": (1_526, 1_721),
}
REV_MIX_2025 = {"Systems": 2473.7, "Services": 1572.1}   # $M; I&A = total - systems - services
BALANCE_SEP26 = {"Cash and equivalents": 2978.1, "Marketable securities, current": 1064.4,
                 "Marketable securities, non-current": 2757.9, "Noncontrolling interest": 131.6,
                 "Shares outstanding (M, Oct 15 2026)": 346.643718}
SETTLEMENT = 450.0    # proposed antitrust settlement disclosed with Q3 2026 results ($M)
PRICE, PRICE_DATE = 407.24, "Nov 19, 2026 benchmark from supplied project; quote not independently verified"
VALUATION_INPUTS = dict(procedures=KPI['da Vinci procedures'][1],
    installed_prior=KPI['da Vinci installed base (year end)'][0],
    installed=KPI['da Vinci installed base (year end)'][1],
    placements=KPI['da Vinci systems placed'][1], revenue=HIST['Revenue'][-1],
    instruments=HIST['Revenue'][-1]-sum(REV_MIX_2025.values()),
    systems=REV_MIX_2025['Systems'], services=REV_MIX_2025['Services'],
    shares=BALANCE_SEP26['Shares outstanding (M, Oct 15 2026)'],
    cash=sum(BALANCE_SEP26[k] for k in ['Cash and equivalents',
        'Marketable securities, current','Marketable securities, non-current']),
    nci=BALANCE_SEP26['Noncontrolling interest'], settlement=SETTLEMENT)

# ---- scenarios ------------------------------------------------------------------
def sga_path(start, step, floor):
    return [max(start - step * i, floor) for i in range(len(YEARS))]

SCENARIOS = {
    "Base": dict(
        proc_g=[.145, .13, .12, .11, .10, .09, .08, .07, .06, .05],
        iapp_g=[.025] + [.01] * 9,
        place_g=[.18, .08, .06, .05, .05, .04, .04, .03, .03, .03],
        svc_g=[.07, .03, .03, .03, .03, .02, .02, .02, .02, .02],
        net_add=.70, gm=.675, rd=.13, sga=sga_path(.215, .003, .19),
        tax=.17, da=.06, capex=.065, nwc=.30, wacc=.09, tg=.035, sbc=.078),
    "Bull": dict(
        proc_g=[.155, .15, .14, .13, .12, .11, .10, .09, .08, .07],
        iapp_g=[.03] + [.015] * 9,
        place_g=[.22, .10, .08, .07, .06, .05, .05, .04, .04, .04],
        svc_g=[.08, .04, .04, .03, .03, .03, .03, .02, .02, .02],
        net_add=.70, gm=.69, rd=.125, sga=sga_path(.21, .004, .18),
        tax=.16, da=.06, capex=.065, nwc=.25, wacc=.085, tg=.04, sbc=.078),
    "Bear": dict(
        proc_g=[.13, .11, .10, .09, .08, .07, .06, .05, .045, .04],
        iapp_g=[.015] + [0.0] * 9,
        place_g=[.12, .04, .03, .03, .03, .02, .02, .02, .02, .02],
        svc_g=[.05, .02, .02, .02, .02, .02, .02, .02, .02, .02],
        net_add=.70, gm=.66, rd=.135, sga=sga_path(.22, .002, .20),
        tax=.18, da=.06, capex=.065, nwc=.40, wacc=.095, tg=.03, sbc=.078),
}


from valuation import value, solve_growth_uplift
REVERSE_UPLIFT = solve_growth_uplift(SCENARIOS['Base'], VALUATION_INPUTS, PRICE)
_reverse_rows = value(SCENARIOS['Base'], VALUATION_INPUTS, REVERSE_UPLIFT)[1]
REVERSE_CAGR = (_reverse_rows[-1]['revenue']/VALUATION_INPUTS['revenue'])**.1-1
KEY_DEBATE = [("Base case", value(SCENARIOS['Base'], VALUATION_INPUTS)[0]),
    ("Base, stock-based compensation added back", value(SCENARIOS['Base'], VALUATION_INPUTS,sbc_addback=.078)[0]),
    ("Base, WACC 8.0% and terminal growth 4.0%", value(SCENARIOS['Base'], VALUATION_INPUTS,wacc=.08,tg=.04)[0]),
    ("Both changes", value(SCENARIOS['Base'], VALUATION_INPUTS,sbc_addback=.078,wacc=.08,tg=.04)[0])]


def header(ws, row, first="", years=True):
    ws.cell(row=row, column=1, value=first)
    if years:
        ws.cell(row=row, column=2, value="2025A")
        for j, y in enumerate(YEARS):
            ws.cell(row=row, column=3 + j, value=f"{y}E")
    for c in range(1, 13):
        ws.cell(row=row, column=c).font = HEAD
        ws.cell(row=row, column=c).fill = HEAD_FILL
        ws.cell(row=row, column=c).alignment = Alignment(horizontal="center" if c > 1 else "left")


def build_historicals(wb):
    ws = wb.create_sheet("Historicals")
    ws["A1"] = "Intuitive Surgical - historical financials ($M unless noted)"; ws["A1"].font = SECTION
    ws.cell(row=3, column=1, value="Fiscal year")
    for j, y in enumerate(HIST_YEARS):
        ws.cell(row=3, column=2 + j, value=y)
    for c in range(1, 8):
        ws.cell(row=3, column=c).font = HEAD; ws.cell(row=3, column=c).fill = HEAD_FILL
    refs, r = {}, 4
    for name, vals in HIST.items():
        ws.cell(row=r, column=1, value=name)
        for j, v in enumerate(vals):
            ws.cell(row=r, column=2 + j, value=v).number_format = M1
        refs[name] = r; r += 1
    for label, num, den in [("Gross margin", "Gross profit", "Revenue"), ("R&D % revenue", "R&D", "Revenue"),
                            ("SG&A % revenue", "SG&A", "Revenue"), ("Operating margin", "Operating income", "Revenue"),
                            ("Effective tax rate", "Income tax", "Pre-tax income"),
                            ("Capex % revenue", "Capital expenditures", "Revenue"),
                            ("SBC % revenue", "Stock-based compensation", "Revenue")]:
        ws.cell(row=r, column=1, value=label).font = Font(italic=True)
        for j in range(len(HIST_YEARS)):
            c = L(2 + j)
            ws.cell(row=r, column=2 + j, value=f"={c}{refs[num]}/{c}{refs[den]}").number_format = PCT
        r += 1
    ws.cell(row=r, column=1, value="Revenue growth").font = Font(italic=True)
    for j in range(1, len(HIST_YEARS)):
        c, p = L(2 + j), L(1 + j)
        ws.cell(row=r, column=2 + j, value=f"={c}{refs['Revenue']}/{p}{refs['Revenue']}-1").number_format = PCT
    r += 2

    ws.cell(row=r, column=1, value="Nine months to Sep 30 ($M)").font = SECTION; r += 1
    for c, v in enumerate(["", "9M 2025", "9M 2026", "Growth / margin 2026"], start=1):
        ws.cell(row=r, column=c, value=v).font = BOLD
    r += 1
    nm = {}
    for name, (a, b) in NINE_MONTHS.items():
        ws.cell(row=r, column=1, value=name)
        ws.cell(row=r, column=2, value=a).number_format = M1
        ws.cell(row=r, column=3, value=b).number_format = M1
        nm[name] = r; r += 1
    ws.cell(row=nm["Revenue"], column=4, value=f"=C{nm['Revenue']}/B{nm['Revenue']}-1").number_format = PCT
    ws.cell(row=nm["Operating income"], column=4,
            value=f"=C{nm['Operating income']}/C{nm['Revenue']}").number_format = PCT
    ws.cell(row=r, column=1, value="Operating income excl. $450M settlement (9M 2026)")
    ws.cell(row=r, column=3, value=f"=C{nm['Operating income']}+{SETTLEMENT}").number_format = M1
    ws.cell(row=r, column=4, value=f"=C{r}/C{nm['Revenue']}").number_format = PCT
    r += 2

    ws.cell(row=r, column=1, value="Operating data").font = SECTION; r += 1
    for c, v in enumerate(["", 2024, 2025, "Growth"], start=1):
        ws.cell(row=r, column=c, value=v).font = BOLD
    r += 1
    for name, (a, b) in KPI.items():
        ws.cell(row=r, column=1, value=name)
        ws.cell(row=r, column=2, value=a).number_format = '#,##0'
        ws.cell(row=r, column=3, value=b).number_format = '#,##0'
        ws.cell(row=r, column=4, value=f"=C{r}/B{r}-1").number_format = PCT
        refs[name] = r; r += 1
    r += 1
    ws.cell(row=r, column=1, value="2025 revenue mix ($M)").font = SECTION; r += 1
    for name, v in REV_MIX_2025.items():
        ws.cell(row=r, column=1, value=f"{name} revenue")
        ws.cell(row=r, column=3, value=v).number_format = M1
        refs[name] = r; r += 1
    ws.cell(row=r, column=1, value="Instruments & accessories revenue (= total - systems - services)")
    ws.cell(row=r, column=3, value=f"=G{refs['Revenue']}-C{refs['Systems']}-C{refs['Services']}").number_format = M1
    refs["I&A"] = r; r += 2

    ws.cell(row=r, column=1, value="Balance sheet items, Sep 30 2026 ($M)").font = SECTION; r += 1
    for name, v in BALANCE_SEP26.items():
        ws.cell(row=r, column=1, value=name)
        ws.cell(row=r, column=3, value=v).number_format = M1
        refs[name] = r; r += 1
    ws.cell(row=r, column=1, value="Total cash and investments (no debt)")
    ws.cell(row=r, column=3, value=f"=SUM(C{refs['Cash and equivalents']}:C{refs['Marketable securities, non-current']})").number_format = M1
    refs["Cash total"] = r
    ws.column_dimensions["A"].width = 58
    for c in "BCDEFG":
        ws.column_dimensions[c].width = 12
    return refs


def build_model(wb, name, a, h):
    ws = wb.create_sheet(f"Model_{name}")
    H = "Historicals!"
    ws["A1"] = f"ISRG - {name} case: driver-based forecast and DCF ($M unless noted)"; ws["A1"].font = SECTION
    ws["A2"] = "Blue = input. Valuation date: Dec 31, 2026; FCF for 2027-2035 discounted, terminal value after 2035."

    header(ws, 4, "Assumptions")
    rows = [("Procedure growth", a["proc_g"]), ("I&A revenue per procedure growth", a["iapp_g"]),
            ("System placement growth", a["place_g"]), ("Service revenue per installed system growth", a["svc_g"]),
            ("Net additions to installed base (% of placements)", [a["net_add"]] * 10),
            ("Gross margin", [a["gm"]] * 10), ("R&D % revenue", [a["rd"]] * 10), ("SG&A % revenue", a["sga"]),
            ("Tax rate", [a["tax"]] * 10), ("D&A % revenue", [a["da"]] * 10), ("Capex % revenue", [a["capex"]] * 10),
            ("Working capital & leased systems, % of revenue increase", [a["nwc"]] * 10),
            ("Stock-based compensation % revenue (used only if B21 = 1)", [a["sbc"]] * 10)]
    for i, (label, vals) in enumerate(rows):
        r = 5 + i
        ws.cell(row=r, column=1, value=label)
        for j, v in enumerate(vals):
            c = ws.cell(row=r, column=3 + j, value=v); c.number_format = PCT; c.font = BLUE
    ws["A18"] = "WACC"; ws["B18"] = a["wacc"]
    ws["A19"] = "Terminal growth"; ws["B19"] = a["tg"]
    ws["A20"] = "Reverse DCF: extra procedure growth per year (use Goal Seek)"; ws["B20"] = 0.0
    ws["A21"] = "Add back stock-based compensation to FCF? (1 = yes, 0 = no; base = 0)"; ws["B21"] = 0
    ws["B21"].font = BLUE
    for c in ("B18", "B19", "B20"):
        ws[c].number_format = PCT; ws[c].font = BLUE

    header(ws, 22, "Operating drivers")
    ws["A23"] = "da Vinci procedures"; ws["B23"] = f"={H}C{h['da Vinci procedures']}"
    ws["A24"] = "I&A revenue per procedure ($)"; ws["B24"] = f"={H}C{h['I&A']}*1e6/B23"
    ws["A25"] = "da Vinci system placements"; ws["B25"] = f"={H}C{h['da Vinci systems placed']}"
    ws["A26"] = "Systems revenue per placement ($)"; ws["B26"] = f"={H}C{h['Systems']}*1e6/B25"
    ws["A27"] = "Installed base, year end"; ws["B27"] = f"={H}C{h['da Vinci installed base (year end)']}"
    ws["A28"] = "Service revenue per average installed system ($)"
    ws["B28"] = (f"={H}C{h['Services']}*1e6/AVERAGE({H}B{h['da Vinci installed base (year end)']},"
                 f"{H}C{h['da Vinci installed base (year end)']})")
    for j in range(10):
        c, p = L(3 + j), L(2 + j)
        ws[f"{c}23"] = f"={p}23*(1+{c}5+$B$20)"
        ws[f"{c}24"] = f"={p}24*(1+{c}6)"
        ws[f"{c}25"] = f"={p}25*(1+{c}7)"
        ws[f"{c}26"] = f"={p}26"
        ws[f"{c}27"] = f"={p}27+{c}25*{c}9"
        ws[f"{c}28"] = f"={p}28*(1+{c}8)"
    for col in range(2, 13):
        for r, fmt in [(23, '#,##0'), (24, USD), (25, '#,##0'), (26, USD), (27, '#,##0'), (28, USD)]:
            ws.cell(row=r, column=col).number_format = fmt

    header(ws, 30, "Income statement")
    ws["A31"] = "Instruments & accessories revenue"; ws["B31"] = f"={H}C{h['I&A']}"
    ws["A32"] = "Systems revenue"; ws["B32"] = f"={H}C{h['Systems']}"
    ws["A33"] = "Services revenue"; ws["B33"] = f"={H}C{h['Services']}"
    ws["A34"] = "Total revenue"; ws["B34"] = "=SUM(B31:B33)"
    ws["A35"] = "Revenue growth"
    ws["A36"] = "Gross profit"; ws["B36"] = f"={H}G{h['Gross profit']}"
    ws["A37"] = "R&D"; ws["B37"] = f"={H}G{h['R&D']}"
    ws["A38"] = "SG&A"; ws["B38"] = f"={H}G{h['SG&A']}"
    ws["A39"] = "Operating income (EBIT), excl. one-time settlement"; ws["B39"] = "=B36-B37-B38"
    ws["A40"] = "EBIT margin"; ws["B40"] = "=B39/B34"
    for j in range(10):
        c, p = L(3 + j), L(2 + j)
        ws[f"{c}31"] = f"={c}23*{c}24/1e6"
        ws[f"{c}32"] = f"={c}25*{c}26/1e6"
        ws[f"{c}33"] = f"={c}28*AVERAGE({p}27,{c}27)/1e6"
        ws[f"{c}34"] = f"=SUM({c}31:{c}33)"
        ws[f"{c}35"] = f"={c}34/{p}34-1"
        ws[f"{c}36"] = f"={c}34*{c}10"
        ws[f"{c}37"] = f"={c}34*{c}11"
        ws[f"{c}38"] = f"={c}34*{c}12"
        ws[f"{c}39"] = f"={c}36-{c}37-{c}38"
        ws[f"{c}40"] = f"={c}39/{c}34"

    header(ws, 42, "Unlevered free cash flow")
    ws["A43"] = "NOPAT = EBIT x (1 - tax)"
    ws["A44"] = "+ Depreciation & amortization"
    ws["A45"] = "- Capital expenditures"
    ws["A46"] = "- Investment in working capital & leased systems"
    ws["A47"] = "Unlevered free cash flow"
    ws["A48"] = "Note: by default stock-based compensation is a real cost (B21 = 0); set B21 = 1 to add it back."
    for j in range(10):
        c, p = L(3 + j), L(2 + j)
        ws[f"{c}43"] = f"={c}39*(1-{c}13)"
        ws[f"{c}44"] = f"={c}34*{c}14"
        ws[f"{c}45"] = f"={c}34*{c}15"
        ws[f"{c}46"] = f"=({c}34-{p}34)*{c}16"
        ws[f"{c}47"] = f"={c}43+{c}44-{c}45-{c}46+$B$21*{c}34*{c}17"
    for r in list(range(31, 40)) + list(range(43, 48)):
        for col in range(2, 13):
            ws.cell(row=r, column=col).number_format = M
    for col in range(2, 13):
        ws.cell(row=35, column=col).number_format = PCT
        ws.cell(row=40, column=col).number_format = PCT
    for r in (34, 39, 47):
        for col in range(1, 13):
            ws.cell(row=r, column=col).font = BOLD
            ws.cell(row=r, column=col).border = THIN

    header(ws, 50, "DCF (value at Dec 31, 2026)")
    ws["A51"] = "Discount period (years)"
    ws["A52"] = "Discount factor"
    ws["A53"] = "PV of free cash flow"
    for j in range(1, 10):
        c = L(3 + j)
        ws[f"{c}51"] = j
        ws[f"{c}52"] = f"=1/(1+$B$18)^{c}51"; ws[f"{c}52"].number_format = '0.000'
        ws[f"{c}53"] = f"={c}47*{c}52"; ws[f"{c}53"].number_format = M
    ws["A54"] = "Stock compensation cash-flow addback sensitivity"
    for j in range(10):
        c = L(3+j)
        ws[f"{c}54"] = f"={c}34*{c}17"
        ws[f"{c}54"].number_format = M
    lines = [
        (55, "Sum of PV of FCF, 2027-2035", "=SUM(D53:L53)"),
        (56, "Terminal value at end-2035 (Gordon growth)", "=L47*(1+B19)/(B18-B19)"),
        (57, "PV of terminal value", "=B56*L52"),
        (58, "Enterprise value", "=B55+B57"),
        (59, "+ Q4 2026E free cash flow (25% of 2026E)", "=0.25*C47"),
        (60, "+ Cash and investments, Sep 30 2026", f"={H}C{h['Cash total']}"),
        (61, "- Proposed antitrust settlement", f"=-{SETTLEMENT}"),
        (62, "- Noncontrolling interest", f"=-{H}C{h['Noncontrolling interest']}"),
        (63, "Equity value", "=SUM(B58:B62)"),
        (64, "Shares outstanding (M)", f"={H}C{h['Shares outstanding (M, Oct 15 2026)']}"),
        (65, "Value per share ($)", "=B63/B64"),
        (66, "Terminal value as % of enterprise value", "=B57/B58"),
    ]
    for r, label, f in lines:
        ws[f"A{r}"] = label; ws[f"B{r}"] = f
        ws[f"B{r}"].number_format = M if r not in (64, 65, 66) else ('#,##0.0' if r == 64 else (USD if r == 65 else PCT))
    for r in (58, 63, 65):
        ws[f"A{r}"].font = BOLD; ws[f"B{r}"].font = BOLD
    ws.column_dimensions["A"].width = 58
    for col in range(2, 13):
        ws.column_dimensions[L(col)].width = 11
    ws.freeze_panes = "B5"
    return ws.title


def build_summary(wb, sheets, h):
    ws = wb["Summary"]
    ws["A1"] = "Intuitive Surgical (NASDAQ: ISRG) - valuation summary"; ws["A1"].font = Font(bold=True, size=14)
    ws["A3"] = "Share price ($)"; ws["B3"] = PRICE; ws["B3"].font = BLUE; ws["B3"].number_format = '$#,##0.00'
    ws["C3"] = PRICE_DATE
    ws["A4"] = "Shares outstanding (M)"; ws["B4"] = f"=Historicals!C{h['Shares outstanding (M, Oct 15 2026)']}"
    ws["B4"].number_format = '#,##0.0'
    ws["A5"] = "Market capitalization ($M)"; ws["B5"] = "=B3*B4"; ws["B5"].number_format = M
    ws["A6"] = "Net cash ($M)"; ws["B6"] = f"=Historicals!C{h['Cash total']}"; ws["B6"].number_format = M
    ws["A7"] = "Market enterprise value ($M)"; ws["B7"] = "=B5-B6"; ws["B7"].number_format = M

    header(ws, 9, "Scenario", years=False)
    for c, v in enumerate(["Scenario", "Value / share", "vs. price", "WACC", "Terminal g",
                           "2030E revenue ($M)", "2030E EBIT margin", "Revenue CAGR 2025-35"], start=1):
        ws.cell(row=9, column=c, value=v)
    for i, nm in enumerate(["Bear", "Base", "Bull"]):
        r, s = 10 + i, f"Model_{nm}!"
        ws[f"A{r}"] = nm
        ws[f"B{r}"] = f"={s}B65"; ws[f"B{r}"].number_format = USD
        ws[f"C{r}"] = f"=B{r}/$B$3-1"; ws[f"C{r}"].number_format = PCT
        ws[f"D{r}"] = f"={s}B18"; ws[f"D{r}"].number_format = PCT
        ws[f"E{r}"] = f"={s}B19"; ws[f"E{r}"].number_format = PCT
        ws[f"F{r}"] = f"={s}G34"; ws[f"F{r}"].number_format = M
        ws[f"G{r}"] = f"={s}G40"; ws[f"G{r}"].number_format = PCT
        ws[f"H{r}"] = f"=({s}L34/{s}B34)^(1/10)-1"; ws[f"H{r}"].number_format = PCT
    ws["A14"] = "Probability-weighted value (25% bear / 50% base / 25% bull)"
    ws["B14"] = "=0.25*B10+0.5*B11+0.25*B12"; ws["B14"].number_format = USD
    ws["C14"] = "=B14/$B$3-1"; ws["C14"].number_format = PCT
    ws["A14"].font = BOLD; ws["B14"].font = BOLD

    ws["A16"] = "What the market price implies (base-case FCF)"; ws["A16"].font = SECTION
    ws["A17"] = "Market EV / 2027E revenue"; ws["B17"] = "=B7/Model_Base!D34"; ws["B17"].number_format = '0.0"x"'
    ws["A18"] = "Market EV / 2027E NOPAT"; ws["B18"] = "=B7/Model_Base!D43"; ws["B18"].number_format = '0.0"x"'
    ws["A19"] = "Reverse DCF: extra procedure growth per year needed to justify the price"
    ws["B19"] = REVERSE_UPLIFT; ws["B19"].number_format = '+0.0%'
    ws["C19"] = "Build-time solver result; rerun after input changes or use Goal Seek"
    ws["A20"] = "  -> implied 2025-35 revenue CAGR vs. 10.2% in the base case"
    ws["B20"] = REVERSE_CAGR; ws["B20"].number_format = PCT

    ws["A21"] = "Sensitivity: base-case value per share by WACC (rows) and terminal growth (columns)"
    ws["A21"].font = SECTION
    ws["A30"] = "Key debate: two judgement calls drive most of the gap to the market price"; ws["A30"].font = SECTION
    for i, (label, v) in enumerate(KEY_DEBATE):
        ws.cell(row=31 + i, column=1, value=label)
        ws.cell(row=31 + i, column=2, value=v).number_format = USD
    ws.cell(row=31 + len(KEY_DEBATE), column=1,
            value="(values computed by re-running the base model with each change; reproduce with B21 and B18/B19)")
    ws["B31"] = "=Model_Base!B65"
    ws["A36"] = "Q4 cash roll-forward assumes 25% of annual modeled FCF; reverse outputs require rerun."
    ws["B32"] = "=B31+(1-Model_Base!B21)*(NPV(Model_Base!B18,Model_Base!D54:L54)+Model_Base!L54*(1+Model_Base!B19)/(Model_Base!B18-Model_Base!B19)/(1+Model_Base!B18)^9+0.25*Model_Base!C54)/Model_Base!B64"
    ws["B33"] = "=D23"
    ws["B34"] = "=B33+(1-Model_Base!B21)*(NPV($A$23,Model_Base!D54:L54)+Model_Base!L54*(1+$D$22)/($A$23-$D$22)/(1+$A$23)^9+0.25*Model_Base!C54)/Model_Base!B64"
    ws["A35"] = "Forward value at Dec 31, 2026; benchmark is from a different date."
    gs, ws_ = [0.03, 0.035, 0.04], [0.08, 0.085, 0.09, 0.095, 0.10]
    ws["A22"] = "WACC \\ terminal g"
    for j, g in enumerate(gs):
        ws.cell(row=22, column=2 + j, value=g).number_format = PCT
    adj = "(Model_Base!B59+Model_Base!B60+Model_Base!B61+Model_Base!B62)"
    for i, w in enumerate(ws_):
        r = 23 + i
        ws.cell(row=r, column=1, value=w).number_format = PCT
        for j in range(len(gs)):
            g = f"{L(2 + j)}$22"
            f = (f"=(NPV($A{r},Model_Base!$D$47:$L$47)+Model_Base!$L$47*(1+{g})/($A{r}-{g})/(1+$A{r})^9"
                 f"+{adj})/Model_Base!$B$64")
            ws.cell(row=r, column=2 + j, value=f).number_format = USD
    for c in range(1, 5):
        ws.cell(row=22, column=c).font = BOLD
    ws.column_dimensions["A"].width = 62
    for c in "BCDEFGH":
        ws.column_dimensions[c].width = 16


def build_comps(wb, h):
    ws = wb.create_sheet("Comps")
    ws["A1"] = "Trading comparables - large-cap medical technology"; ws["A1"].font = SECTION
    ws["A2"] = ("Blue cells are inputs. Peer data are unverified supplied assumptions; "
                "fill the others from a live market data source before use.")
    cols = ["Company", "Ticker", "Price ($)", "Price date", "Shares (M)", "Market cap ($M)", "Net debt ($M)",
            "EV ($M)", "LTM revenue ($M)", "LTM EBIT ($M)", "LTM net income ($M)", "EV / revenue",
            "EV / EBIT", "P / E", "Status"]
    for c, v in enumerate(cols, start=1):
        cell = ws.cell(row=4, column=c, value=v); cell.font = HEAD; cell.fill = HEAD_FILL
    rows = [
        ["Intuitive Surgical", "ISRG", "=Summary!B3", "Summary", "=Summary!B4", None,
         f"=-Historicals!C{h['Cash total']}", None, 11578.0, 3734.2, 3330.5,
         None, None, None, "LTM to Sep 2026 (SEC); EBIT and net income exclude the $450M settlement (tax-effected at 17%)"],
        ["Stryker", "SYK", 318.15, "Nov 18, 2026", 380.691, None, 11909.0, None, 26300.0, 5433.0, 3623.0,
         None, None, None, "Supplied inputs; not independently verified in this review"],
        ["Boston Scientific", "BSX", None, "", None, None, None, None, None, None, None, None, None, None, "fill in"],
        ["Medtronic", "MDT", None, "", None, None, None, None, None, None, None, None, None, None, "fill in"],
        ["Edwards Lifesciences", "EW", None, "", None, None, None, None, None, None, None, None, None, None, "fill in"],
    ]
    for i, row in enumerate(rows):
        r = 5 + i
        for c, v in enumerate(row, start=1):
            cell = ws.cell(row=r, column=c, value=v)
            if c in (3, 5, 7, 9, 10, 11) and i > 0:
                cell.font = BLUE
        ws[f"F{r}"] = f'=IF(AND(ISNUMBER(C{r}),ISNUMBER(E{r})),C{r}*E{r},"")'
        ws[f"H{r}"] = f'=IF(ISNUMBER(F{r}),F{r}+G{r},"")'
        ws[f"L{r}"] = f'=IF(AND(ISNUMBER(H{r}),ISNUMBER(I{r})),H{r}/I{r},"")'
        ws[f"M{r}"] = f'=IF(AND(ISNUMBER(H{r}),ISNUMBER(J{r})),H{r}/J{r},"")'
        ws[f"N{r}"] = f'=IF(AND(ISNUMBER(F{r}),ISNUMBER(K{r})),F{r}/K{r},"")'
        for c, fmt in [(3, '$#,##0.00'), (5, '#,##0.0'), (6, M), (7, M), (8, M), (9, M), (10, M), (11, M),
                       (12, '0.0"x"'), (13, '0.0"x"'), (14, '0.0"x"')]:
            ws.cell(row=r, column=c).number_format = fmt
    ws["A11"] = "Median of peers with data"
    for col in "LMN":
        ws[f"{col}11"] = 'Unverified'
        ws[f"{col}11"].number_format = '0.0"x"'
    ws["A12"] = "ISRG premium to peer median"
    for col in "LMN":
        ws[f"{col}12"] = 'Unverified'
        ws[f"{col}12"].number_format = PCT
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["O"].width = 60
    for c in range(2, 15):
        ws.column_dimensions[L(c)].width = 13


def build_notes(wb):
    ws = wb.create_sheet("Notes")
    notes = [
        "METHOD",
        "Revenue is built from three drivers: instruments & accessories = procedures x revenue per procedure;",
        "systems = placements x revenue per placement; services = average installed base x service revenue per system.",
        "Free cash flow = EBIT x (1 - tax) + D&A - capex - investment in working capital and leased systems.",
        "Stock-based compensation is treated as a real cost: it stays in operating expenses and is not added back.",
        "DCF: 10-year forecast (2026-2035), value at Dec 31 2026, Gordon-growth terminal value after 2035.",
        "",
        "KEY JUDGEMENTS (change these first)",
        "1. Procedure growth path. 2026 base = model assumption near Q2 guidance midpoint (14.5% da Vinci). It then fades by ~1pt a year.",
        "2. Working capital & leased systems. Many systems are placed on usage-based leases, so cash is tied up",
        "   in equipment that is not in capex. History is noisy (about 55% of the revenue increase in 2025, ~17%",
        "   in 9M 2026); base case uses 30%.",
        "3. WACC 9.0% base = risk-free 4.25% + beta 1.05 x equity risk premium 4.5% (no debt). Check rates.",
        "4. The $450M antitrust settlement is treated as one-time: excluded from EBIT, deducted from equity value.",
        "",
        "REVERSE DCF",
        "In Model_Base, Data > What-If Analysis > Goal Seek: set B65 to the share price by changing B20.",
        "B20 adds the same extra percentage points to procedure growth in every forecast year.",
        "",
        "SOURCES",
        "Financial statements: SEC XBRL company facts (data.sec.gov), 10-K and 10-Q filings, CIK 0001035267.",
        "Procedures, installed base, placements, revenue mix: Intuitive earnings releases (Q4 2025, Q3 2026).",
        "Share price: Nov 19 2026 close as reported by GuruFocus. Peer price: Stryker Nov 18 2026 (Macrotrends).",
        "",
        "LIMITATIONS",
        "Procedures are da Vinci only, while I&A revenue also includes Ion; revenue per procedure is approximate.",
        "Interest income on the cash pile and future buybacks are ignored. No debt.",
        "Peer inputs, including Stryker, remain unverified supplied assumptions; do not rely on the median.",
    ]
    reviewed = Path(__file__).resolve().parents[1] / 'review_notes.json'
    if reviewed.exists():
        notes = json.loads(reviewed.read_text())
    for i, t in enumerate(notes, start=1):
        ws.cell(row=i, column=1, value=t)
        if t.isupper() and t:
            ws.cell(row=i, column=1).font = BOLD
    ws.column_dimensions["A"].width = 120


def main():
    wb = Workbook()
    wb.active.title = "Summary"
    h = build_historicals(wb)
    sheets = [build_model(wb, nm, SCENARIOS[nm], h) for nm in ["Base", "Bull", "Bear"]]
    build_summary(wb, sheets, h)
    build_comps(wb, h)
    build_notes(wb)
    wb.save(OUT)
    print("saved", OUT)


if __name__ == "__main__":
    main()
