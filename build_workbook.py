"""
Build the daily gold reconciliation workbook.

All data in this script is synthetic. Rates are illustrative and are not
quotes from any exchange. No real business, karigar, or customer appears
anywhere in it.

Run:  python build_workbook.py
Out:  gold_reconciliation_system.xlsx
"""
import random
from datetime import date, timedelta

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

random.seed(20261010)

OUT = "gold_reconciliation_system.xlsx"

# ---------------------------------------------------------------- styling
ARIAL = "Arial"
H1 = Font(name=ARIAL, size=14, bold=True)
H2 = Font(name=ARIAL, size=11, bold=True, color="FFFFFF")
LBL = Font(name=ARIAL, size=10, bold=True)
BODY = Font(name=ARIAL, size=10)
INPUT = Font(name=ARIAL, size=10, color="0000FF")          # hardcoded input
LINK = Font(name=ARIAL, size=10, color="008000")           # cross-sheet link
NOTE = Font(name=ARIAL, size=9, italic=True, color="595959")

HDR_FILL = PatternFill("solid", fgColor="1F3864")
ASSUM_FILL = PatternFill("solid", fgColor="FFFF00")        # fill-me-in cells
BAND_FILL = PatternFill("solid", fgColor="F2F2F2")

THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

INR = '₹#,##0;(₹#,##0);-'
INR2 = '₹#,##0.00;(₹#,##0.00);-'
GRAM = '#,##0.000;(#,##0.000);-'
PCT = '0.0%'
DATE_F = "dd-mmm-yy"


def header_row(ws, row, labels, widths=None):
    for i, text in enumerate(labels, start=1):
        c = ws.cell(row=row, column=i, value=text)
        c.font = H2
        c.fill = HDR_FILL
        c.border = BOX
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    if widths:
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[row].height = 30
    ws.freeze_panes = ws.cell(row=row + 1, column=1)


def title(ws, text, sub=None):
    ws["A1"] = text
    ws["A1"].font = H1
    if sub:
        ws["A2"] = sub
        ws["A2"].font = NOTE


# ---------------------------------------------------------------- constants
PURITY = {"24K": 0.995, "22K": 0.916, "18K": 0.750, "14K": 0.585}
KARIGARS = ["K-01", "K-02", "K-03", "K-04", "K-05"]
KARIGAR_TRADE = {
    "K-01": "Chain making",
    "K-02": "Chain making",
    "K-03": "Stone setting",
    "K-04": "Casting",
    "K-05": "Hand finishing",
}
# wastage norm as a fraction of fine gold issued, by trade
NORM = {"Chain making": 0.018, "Stone setting": 0.035, "Casting": 0.025,
        "Hand finishing": 0.028}

CATEGORIES = ["Chain", "Bangle", "Ring", "Necklace", "Earring"]

START = date(2026, 10, 1)
N_DAYS = 22

business_days = []
d = START
while len(business_days) < N_DAYS:
    if d.weekday() != 6:                      # shop shut Sundays
        business_days.append(d)
    d += timedelta(days=1)

wb = Workbook()

# =======================================================================
# 1. README
# =======================================================================
ws = wb.active
ws.title = "README"
ws.column_dimensions["A"].width = 112
title(ws, "Daily Gold Reconciliation System",
      "All data in this workbook is synthetic. Rates are illustrative.")

readme_lines = [
    ("", ""),
    ("WHAT THIS SOLVES", "h"),
    ("A retail jeweller holds the same metal in four places at once: on the shelf, with "
     "karigars (goldsmiths) being worked into finished pieces, in scrap awaiting melt, and "
     "in old gold taken in part-exchange from customers.", "b"),
    ("Each sits at a different purity. 10 grams of 22K is not 10 grams of 24K. Adding gross "
     "weights together produces a number that means nothing, which is why shops that count "
     "in gross weight cannot tell a genuine loss from a purity mix shift.", "b"),
    ("Every quantity here is converted to FINE GOLD before anything is added or compared. "
     "Fine gold = gross weight x purity. On that basis every figure in the workbook is "
     "directly comparable.", "b"),
    ("", ""),
    ("THE THREE QUESTIONS IT ANSWERS EVERY DAY", "h"),
    ("1. Does counted physical stock match the book? Variance is shown in grams fine and in "
     "rupees, against a tolerance set in Setup.", "b"),
    ("2. Is the shop's board rate still tracking the market? A board rate left stale while "
     "spot moves means metal is sold below replacement cost on every ticket.", "b"),
    ("3. How much making charge is being given away? List making charge versus realised, by "
     "day and by month. This is the most commonly unmeasured leak in the P&L.", "b"),
    ("", ""),
    ("SHEETS", "h"),
    ("Setup              Purities, wastage norms by trade, variance and rate tolerances. "
     "Yellow cells are the ones to edit.", "b"),
    ("Rate_Sheet         Daily spot, the shop's board rate, and the drift between them.", "b"),
    ("Metal_Ledger       Opening, purchases, issues, returns, sales, old gold, closing book "
     "vs counted physical. All in grams fine.", "b"),
    ("Karigar_Ledger     Metal out with each karigar, what came back, wastage against the "
     "norm for that trade.", "b"),
    ("SKU_Register       Product master. Gross weight, stone weight, net gold, fine gold, "
     "and list making charge per piece.", "b"),
    ("Sales_Register     Line-level sales. Metal value, making charge list vs realised, old "
     "gold credit, net invoice.", "b"),
    ("Daily_Recon        The daily one-screen output. Pick a date in the yellow cell.", "b"),
    ("Monthly_Analysis   Monthly rollup of volume, leakage, realisation and variance.", "b"),
    ("", ""),
    ("CONVENTIONS", "h"),
    ("Blue text = hardcoded input.  Black = formula.  Green = link to another sheet.  "
     "Yellow fill = edit this.", "b"),
    ("Weights in grams to three decimals. Fine gold reference purity is set in Setup.", "b"),
]

r = 4
for text, kind in readme_lines:
    c = ws.cell(row=r, column=1, value=text)
    if kind == "h":
        c.font = LBL
    elif kind == "b":
        c.font = BODY
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 28 if len(text) > 100 else 14
    r += 1

# =======================================================================
# 2. Setup
# =======================================================================
ws = wb.create_sheet("Setup")
ws.column_dimensions["A"].width = 34
ws.column_dimensions["B"].width = 16
ws.column_dimensions["C"].width = 14
ws.column_dimensions["D"].width = 52
title(ws, "Setup", "Yellow cells are inputs. Everything else reads from here.")

ws["A4"] = "Shop reference"
ws["A4"].font = LBL
rows = [
    ("Business name", "Sample Jeweller (illustrative)", "text"),
    ("Reporting currency", "INR", "text"),
    ("Fine gold reference purity", 0.995, "num3"),
]
r = 5
for label, val, _ in rows:
    ws.cell(row=r, column=1, value=label).font = BODY
    c = ws.cell(row=r, column=2, value=val)
    c.font = INPUT
    c.fill = ASSUM_FILL
    c.border = BOX
    r += 1
ws["B7"].number_format = "0.000"

ws["A9"] = "Purity table"
ws["A9"].font = LBL
header_row(ws, 10, ["Karat", "Purity", "", ""])
ws.freeze_panes = None
r = 11
for k, p in PURITY.items():
    ws.cell(row=r, column=1, value=k).font = BODY
    c = ws.cell(row=r, column=2, value=p)
    c.font = INPUT
    c.fill = ASSUM_FILL
    c.number_format = "0.000"
    c.border = BOX
    r += 1
PURITY_RANGE = f"$A$11:$B${r-1}"

ws["A17"] = "Wastage norm by trade"
ws["A17"].font = LBL
ws["D17"] = "Fraction of fine gold issued that may be lost in working."
ws["D17"].font = NOTE
header_row(ws, 18, ["Trade", "Norm %", "", ""])
ws.freeze_panes = None
r = 19
for trade, n in NORM.items():
    ws.cell(row=r, column=1, value=trade).font = BODY
    c = ws.cell(row=r, column=2, value=n)
    c.font = INPUT
    c.fill = ASSUM_FILL
    c.number_format = PCT
    c.border = BOX
    r += 1
NORM_RANGE = f"$A$19:$B${r-1}"

ws["A24"] = "Tolerances"
ws["A24"].font = LBL
tol = [
    ("Stock variance tolerance (g fine)", 2.0, GRAM,
     "Above this, the day is flagged for investigation."),
    ("Board rate drift tolerance (₹/g fine)", 150.0, INR2,
     "Board rate allowed to sit this far from spot before repricing."),
    ("Making charge realisation floor", 0.85, PCT,
     "Below this, discounting is flagged."),
]
r = 25
for label, val, fmt, note in tol:
    ws.cell(row=r, column=1, value=label).font = BODY
    c = ws.cell(row=r, column=2, value=val)
    c.font = INPUT
    c.fill = ASSUM_FILL
    c.number_format = fmt
    c.border = BOX
    ws.cell(row=r, column=4, value=note).font = NOTE
    r += 1

ws["A30"] = ("Source: all values above are illustrative defaults chosen for this sample. "
             "Replace with the shop's own norms before use.")
ws["A30"].font = NOTE

# =======================================================================
# 3. Rate_Sheet
# =======================================================================
ws = wb.create_sheet("Rate_Sheet")
title(ws, "Rate Sheet", "Spot vs the shop's own board rate. Drift is the signal to reprice.")
header_row(ws, 4, [
    "Date", "Month", "Spot 24K (₹/10g)", "Spot fine (₹/g)",
    "Board rate 22K (₹/g)", "Board implied fine (₹/g)",
    "Drift (₹/g fine)", "Status",
], [12, 10, 16, 14, 16, 16, 15, 14])

# a spot path that trends up with noise, and a board rate that lags it
spot = 92400.0
board_22k = None
rate_rows = {}
r = 5
for i, dt in enumerate(business_days):
    spot *= (1 + random.gauss(0.0022, 0.0078))
    # the shop resets its board rate only every fourth day. That lag is the point:
    # in a rising market a stale board rate sells metal below replacement cost.
    if board_22k is None or i % 4 == 0:
        board_22k = (spot / 10) * PURITY["22K"] * (1 + random.gauss(0.0012, 0.0010))

    ws.cell(row=r, column=1, value=dt).number_format = DATE_F
    ws.cell(row=r, column=1).font = BODY
    ws.cell(row=r, column=2, value=f'=TEXT(A{r},"YYYY-MM")').font = BODY
    c = ws.cell(row=r, column=3, value=round(spot, 0))
    c.font = INPUT
    c.number_format = INR
    ws.cell(row=r, column=4, value=f"=C{r}/10").number_format = INR2
    c = ws.cell(row=r, column=5, value=round(board_22k, 2))
    c.font = INPUT
    c.number_format = INR2
    ws.cell(row=r, column=6,
            value=f"=E{r}/INDEX(Setup!$B$11:$B$14,MATCH(\"22K\",Setup!$A$11:$A$14,0))"
            ).number_format = INR2
    ws.cell(row=r, column=7, value=f"=F{r}-D{r}").number_format = INR2
    ws.cell(row=r, column=8,
            value=f'=IF(ABS(G{r})>Setup!$B$26,IF(G{r}<0,"REPRICE UP","REPRICE DOWN"),"OK")'
            ).font = BODY
    for col in range(1, 9):
        ws.cell(row=r, column=col).border = BOX
    rate_rows[dt] = r
    r += 1
RATE_LAST = r - 1

ws.cell(row=RATE_LAST + 2, column=1,
        value="Spot and board rates are synthetic. Board rate is reset every third day on "
              "purpose, to create the drift the recon is built to catch."
        ).font = NOTE

# =======================================================================
# 4. SKU_Register
# =======================================================================
ws = wb.create_sheet("SKU_Register")
title(ws, "SKU Register", "Product master. Fine gold per piece is what the ledger consumes.")
header_row(ws, 4, [
    "SKU", "Category", "Karat", "Purity", "Gross wt (g)", "Stone wt (g)",
    "Net gold wt (g)", "Fine gold (g)", "MC basis", "MC rate", "MC list (₹)",
], [12, 12, 8, 9, 13, 12, 13, 12, 13, 11, 14])

skus = []
r = 5
for i in range(18):
    cat = CATEGORIES[i % len(CATEGORIES)]
    karat = "22K" if cat != "Ring" else random.choice(["22K", "18K"])
    gross = round(random.uniform(4.0, 46.0), 3)
    stone = round(gross * random.uniform(0.0, 0.14), 3) if cat in ("Ring", "Necklace", "Earring") else 0.0
    basis = "Per gram" if cat in ("Chain", "Bangle") else "% of metal"
    mc_rate = round(random.uniform(340, 620), 0) if basis == "Per gram" else round(random.uniform(0.09, 0.17), 4)
    sku = f"{cat[:2].upper()}-{1000 + i}"

    ws.cell(row=r, column=1, value=sku).font = BODY
    ws.cell(row=r, column=2, value=cat).font = BODY
    ws.cell(row=r, column=3, value=karat).font = INPUT
    ws.cell(row=r, column=4,
            value=f"=INDEX(Setup!$B$11:$B$14,MATCH(C{r},Setup!$A$11:$A$14,0))"
            ).number_format = "0.000"
    ws.cell(row=r, column=5, value=gross).font = INPUT
    ws.cell(row=r, column=5).number_format = GRAM
    ws.cell(row=r, column=6, value=stone).font = INPUT
    ws.cell(row=r, column=6).number_format = GRAM
    ws.cell(row=r, column=7, value=f"=E{r}-F{r}").number_format = GRAM
    ws.cell(row=r, column=8, value=f"=G{r}*D{r}").number_format = GRAM
    ws.cell(row=r, column=9, value=basis).font = BODY
    c = ws.cell(row=r, column=10, value=mc_rate)
    c.font = INPUT
    c.number_format = INR if basis == "Per gram" else PCT
    # list MC: per-gram charges on net gold weight, percentage charges on metal value at spot
    ws.cell(row=r, column=11,
            value=f'=IF(I{r}="Per gram",G{r}*J{r},'
                  f"G{r}*D{r}*INDEX(Rate_Sheet!$D$5:$D${RATE_LAST},1)*J{r})"
            ).number_format = INR
    for col in range(1, 12):
        ws.cell(row=r, column=col).border = BOX
    skus.append({"sku": sku, "row": r, "cat": cat, "karat": karat,
                 "net": round(gross - stone, 3), "purity": PURITY[karat]})
    r += 1
SKU_LAST = r - 1

# =======================================================================
# 5. Karigar_Ledger
# =======================================================================
ws = wb.create_sheet("Karigar_Ledger")
title(ws, "Karigar Ledger",
      "Metal issued to goldsmiths, what returned, and wastage against the norm for the trade.")
header_row(ws, 4, [
    "Date", "Month", "Karigar", "Trade", "Karat", "Purity",
    "Issued gross (g)", "Issued fine (g)", "Returned finished gross (g)",
    "Returned fine (g)", "Scrap returned fine (g)", "Wastage fine (g)",
    "Wastage %", "Norm %", "vs Norm", "Status",
], [12, 10, 10, 15, 8, 9, 14, 13, 16, 13, 15, 14, 11, 10, 11, 14])

r = 5
karigar_rows = []
for dt in business_days:
    for kg in random.sample(KARIGARS, k=random.choice([2, 3])):
        trade = KARIGAR_TRADE[kg]
        karat = "22K"
        issued_gross = round(random.uniform(55, 240), 3)
        norm = NORM[trade]
        # actual wastage sits near the norm, with K-03 running persistently hot.
        # One karigar quietly over norm for months is the classic slow leak.
        drift = 0.019 if kg == "K-03" else 0.0
        actual_w = max(0.0, random.gauss(norm + drift, norm * 0.22))
        issued_fine = issued_gross * PURITY[karat]
        returned_fine = issued_fine * (1 - actual_w) * random.uniform(0.86, 0.94)
        scrap_fine = issued_fine * (1 - actual_w) - returned_fine

        ws.cell(row=r, column=1, value=dt).number_format = DATE_F
        ws.cell(row=r, column=1).font = BODY
        ws.cell(row=r, column=2, value=f'=TEXT(A{r},"YYYY-MM")').font = BODY
        ws.cell(row=r, column=3, value=kg).font = BODY
        ws.cell(row=r, column=4, value=trade).font = BODY
        ws.cell(row=r, column=5, value=karat).font = INPUT
        ws.cell(row=r, column=6,
                value=f"=INDEX(Setup!$B$11:$B$14,MATCH(E{r},Setup!$A$11:$A$14,0))"
                ).number_format = "0.000"
        ws.cell(row=r, column=7, value=round(issued_gross, 3)).font = INPUT
        ws.cell(row=r, column=7).number_format = GRAM
        ws.cell(row=r, column=8, value=f"=G{r}*F{r}").number_format = GRAM
        ws.cell(row=r, column=9, value=round(returned_fine / PURITY[karat], 3)).font = INPUT
        ws.cell(row=r, column=9).number_format = GRAM
        ws.cell(row=r, column=10, value=f"=I{r}*F{r}").number_format = GRAM
        ws.cell(row=r, column=11, value=round(scrap_fine, 3)).font = INPUT
        ws.cell(row=r, column=11).number_format = GRAM
        ws.cell(row=r, column=12, value=f"=H{r}-J{r}-K{r}").number_format = GRAM
        ws.cell(row=r, column=13, value=f"=IFERROR(L{r}/H{r},0)").number_format = PCT
        ws.cell(row=r, column=14,
                value=f"=INDEX(Setup!$B$19:$B$22,MATCH(D{r},Setup!$A$19:$A$22,0))"
                ).number_format = PCT
        ws.cell(row=r, column=15, value=f"=M{r}-N{r}").number_format = PCT
        ws.cell(row=r, column=16,
                value=f'=IF(O{r}>0.01,"OVER NORM",IF(O{r}<-0.005,"BELOW NORM","OK"))'
                ).font = BODY
        for col in range(1, 17):
            ws.cell(row=r, column=col).border = BOX
        karigar_rows.append((dt, r))
        r += 1
KG_LAST = r - 1

# =======================================================================
# 6. Sales_Register
# =======================================================================
ws = wb.create_sheet("Sales_Register")
title(ws, "Sales Register",
      "Line-level sales. Making charge list vs realised is the leakage measure.")
header_row(ws, 4, [
    "Date", "Month", "Invoice", "SKU", "Qty", "Net gold wt (g)", "Fine gold (g)",
    "Rate applied (₹/g fine)", "Metal value (₹)", "MC list (₹)", "Discount %",
    "MC realised (₹)", "MC given away (₹)", "Old gold in fine (g)",
    "Old gold credit (₹)", "Net invoice (₹)",
], [12, 10, 12, 12, 7, 14, 13, 17, 15, 14, 11, 15, 16, 15, 16, 15])

r = 5
inv = 4100
for dt in business_days:
    rate_row = rate_rows[dt]
    for _ in range(random.choice([2, 3, 3, 4])):
        s = random.choice(skus)
        qty = random.choice([1, 1, 1, 2])
        disc = random.choice([0.0, 0.05, 0.10, 0.10, 0.15, 0.20, 0.25, 0.30])
        old_gold = round(random.uniform(0, 18), 3) if random.random() < 0.30 else 0.0
        inv += 1

        ws.cell(row=r, column=1, value=dt).number_format = DATE_F
        ws.cell(row=r, column=1).font = BODY
        ws.cell(row=r, column=2, value=f'=TEXT(A{r},"YYYY-MM")').font = BODY
        ws.cell(row=r, column=3, value=f"INV-{inv}").font = BODY
        ws.cell(row=r, column=4, value=s["sku"]).font = INPUT
        ws.cell(row=r, column=5, value=qty).font = INPUT
        ws.cell(row=r, column=6,
                value=f"=E{r}*INDEX(SKU_Register!$G$5:$G${SKU_LAST},"
                      f"MATCH(D{r},SKU_Register!$A$5:$A${SKU_LAST},0))"
                ).number_format = GRAM
        ws.cell(row=r, column=7,
                value=f"=E{r}*INDEX(SKU_Register!$H$5:$H${SKU_LAST},"
                      f"MATCH(D{r},SKU_Register!$A$5:$A${SKU_LAST},0))"
                ).number_format = GRAM
        # metal is billed at the shop's board rate, not at spot: that is the leak
        ws.cell(row=r, column=8, value=f"=Rate_Sheet!$F${rate_row}").font = LINK
        ws.cell(row=r, column=8).number_format = INR2
        ws.cell(row=r, column=9, value=f"=G{r}*H{r}").number_format = INR
        ws.cell(row=r, column=10,
                value=f"=E{r}*INDEX(SKU_Register!$K$5:$K${SKU_LAST},"
                      f"MATCH(D{r},SKU_Register!$A$5:$A${SKU_LAST},0))"
                ).number_format = INR
        c = ws.cell(row=r, column=11, value=disc)
        c.font = INPUT
        c.number_format = PCT
        ws.cell(row=r, column=12, value=f"=J{r}*(1-K{r})").number_format = INR
        ws.cell(row=r, column=13, value=f"=J{r}-L{r}").number_format = INR
        c = ws.cell(row=r, column=14, value=old_gold)
        c.font = INPUT
        c.number_format = GRAM
        ws.cell(row=r, column=15, value=f"=N{r}*H{r}").number_format = INR
        ws.cell(row=r, column=16, value=f"=I{r}+L{r}-O{r}").number_format = INR
        for col in range(1, 17):
            ws.cell(row=r, column=col).border = BOX
        r += 1
SALES_LAST = r - 1

tr = SALES_LAST + 1
ws.cell(row=tr, column=1, value="TOTAL").font = LBL
for col, fmt in [(7, GRAM), (9, INR), (10, INR), (12, INR), (13, INR), (15, INR), (16, INR)]:
    L = get_column_letter(col)
    c = ws.cell(row=tr, column=col, value=f"=SUM({L}5:{L}{SALES_LAST})")
    c.font = LBL
    c.number_format = fmt
    c.fill = BAND_FILL

# =======================================================================
# 7. Metal_Ledger
# =======================================================================
ws = wb.create_sheet("Metal_Ledger")
title(ws, "Metal Ledger",
      "Everything on a fine gold basis, so shelf, karigar and scrap metal are comparable.")
header_row(ws, 4, [
    "Date", "Month", "Opening book (g fine)", "Purchases (g fine)",
    "Issued to karigar (g fine)", "Received from karigar (g fine)",
    "Sold (g fine)", "Old gold taken in (g fine)", "Closing book (g fine)",
    "Physical counted (g fine)", "Variance (g fine)", "Variance (₹)",
    "Tolerance (g)", "Status",
], [12, 10, 16, 15, 17, 18, 13, 16, 16, 16, 14, 14, 11, 16])

OPENING = 1820.0
r = 5
metal_rows = {}
for i, dt in enumerate(business_days):
    rate_row = rate_rows[dt]
    ws.cell(row=r, column=1, value=dt).number_format = DATE_F
    ws.cell(row=r, column=1).font = BODY
    ws.cell(row=r, column=2, value=f'=TEXT(A{r},"YYYY-MM")').font = BODY

    if i == 0:
        c = ws.cell(row=r, column=3, value=OPENING)
        c.font = INPUT
        c.fill = ASSUM_FILL
    else:
        ws.cell(row=r, column=3, value=f"=I{r-1}")
    ws.cell(row=r, column=3).number_format = GRAM

    purchase = round(random.uniform(0, 260), 3) if random.random() < 0.35 else 0.0
    c = ws.cell(row=r, column=4, value=purchase)
    c.font = INPUT
    c.number_format = GRAM

    ws.cell(row=r, column=5,
            value=f"=SUMIFS(Karigar_Ledger!$H$5:$H${KG_LAST},"
                  f"Karigar_Ledger!$A$5:$A${KG_LAST},A{r})"
            ).number_format = GRAM
    ws.cell(row=r, column=6,
            value=f"=SUMIFS(Karigar_Ledger!$J$5:$J${KG_LAST},"
                  f"Karigar_Ledger!$A$5:$A${KG_LAST},A{r})"
                  f"+SUMIFS(Karigar_Ledger!$K$5:$K${KG_LAST},"
                  f"Karigar_Ledger!$A$5:$A${KG_LAST},A{r})"
            ).number_format = GRAM
    ws.cell(row=r, column=7,
            value=f"=SUMIFS(Sales_Register!$G$5:$G${SALES_LAST},"
                  f"Sales_Register!$A$5:$A${SALES_LAST},A{r})"
            ).number_format = GRAM
    ws.cell(row=r, column=8,
            value=f"=SUMIFS(Sales_Register!$N$5:$N${SALES_LAST},"
                  f"Sales_Register!$A$5:$A${SALES_LAST},A{r})"
            ).number_format = GRAM
    ws.cell(row=r, column=9,
            value=f"=C{r}+D{r}-E{r}+F{r}-G{r}+H{r}").number_format = GRAM

    # Counted physical tracks the book with small genuine discrepancies, plus three
    # deliberate breach days so the exception logic demonstrates itself:
    #   day 4  a real short that never comes back
    #   day 11 a miscount, reversed the next day
    #   day 18 metal returned by a karigar but not booked until the following day
    seeded = {3: -3.42, 10: -4.18, 11: +4.05, 17: -2.86}
    discrepancy = seeded.get(i, round(random.gauss(0, 0.62), 3))
    ws.cell(row=r, column=10,
            value=f"=I{r}+{discrepancy}").number_format = GRAM
    ws.cell(row=r, column=11, value=f"=J{r}-I{r}").number_format = GRAM
    ws.cell(row=r, column=12,
            value=f"=K{r}*Rate_Sheet!$D${rate_row}").number_format = INR
    ws.cell(row=r, column=13, value="=Setup!$B$25").number_format = GRAM
    ws.cell(row=r, column=13).font = LINK
    ws.cell(row=r, column=14,
            value=f'=IF(ABS(K{r})>M{r},"INVESTIGATE","OK")').font = BODY
    for col in range(1, 15):
        ws.cell(row=r, column=col).border = BOX
    metal_rows[dt] = r
    r += 1
METAL_LAST = r - 1

ws.cell(row=METAL_LAST + 2, column=1,
        value="Issues, returns, sales and old gold all pull from their source sheets by "
              "SUMIFS on date, so the ledger cannot drift from the registers."
        ).font = NOTE

# =======================================================================
# 8. Daily_Recon
# =======================================================================
ws = wb.create_sheet("Daily_Recon")
ws.column_dimensions["A"].width = 42
ws.column_dimensions["B"].width = 18
ws.column_dimensions["C"].width = 18
ws.column_dimensions["D"].width = 46
title(ws, "Daily Reconciliation", "Change the date in the yellow cell. Everything recalculates.")

ws["A4"] = "Reporting date"
ws["A4"].font = LBL
c = ws["B4"]
c.value = business_days[10]      # opens on a day that breaches variance tolerance
c.font = INPUT
c.fill = ASSUM_FILL
c.number_format = DATE_F
c.border = BOX

MR = f"Metal_Ledger!$A$5:$A${METAL_LAST}"
RR = f"Rate_Sheet!$A$5:$A${RATE_LAST}"
SR = f"Sales_Register!$A$5:$A${SALES_LAST}"
KR = f"Karigar_Ledger!$A$5:$A${KG_LAST}"


def block(ws, row, heading):
    c = ws.cell(row=row, column=1, value=heading)
    c.font = H2
    c.fill = HDR_FILL
    for col in range(1, 5):
        ws.cell(row=row, column=col).fill = HDR_FILL
        ws.cell(row=row, column=col).border = BOX
    return row + 1


def line(ws, row, label, formula, fmt, note=""):
    ws.cell(row=row, column=1, value=label).font = BODY
    c = ws.cell(row=row, column=2, value=formula)
    c.number_format = fmt
    c.border = BOX
    if note:
        ws.cell(row=row, column=4, value=note).font = NOTE
    return row + 1


r = block(ws, 6, "1.  STOCK POSITION")
r = line(ws, r, "Opening book (g fine)",
         f"=SUMIFS(Metal_Ledger!$C$5:$C${METAL_LAST},{MR},$B$4)", GRAM)
r = line(ws, r, "Closing book (g fine)",
         f"=SUMIFS(Metal_Ledger!$I$5:$I${METAL_LAST},{MR},$B$4)", GRAM)
r = line(ws, r, "Physical counted (g fine)",
         f"=SUMIFS(Metal_Ledger!$J$5:$J${METAL_LAST},{MR},$B$4)", GRAM)
r = line(ws, r, "Variance (g fine)",
         f"=SUMIFS(Metal_Ledger!$K$5:$K${METAL_LAST},{MR},$B$4)", GRAM,
         "Physical minus book. Negative means metal is missing.")
r = line(ws, r, "Variance at today's spot (₹)",
         f"=SUMIFS(Metal_Ledger!$L$5:$L${METAL_LAST},{MR},$B$4)", INR)
r = line(ws, r, "Tolerance (g fine)", "=Setup!$B$25", GRAM)
r = line(ws, r, "Verdict",
         f'=IF(ABS(B10)>B12,"INVESTIGATE","WITHIN TOLERANCE")', "General")
r += 1

r = block(ws, r, "2.  METAL OUT WITH KARIGARS")
r = line(ws, r, "Issued today (g fine)",
         f"=SUMIFS(Karigar_Ledger!$H$5:$H${KG_LAST},{KR},$B$4)", GRAM)
r = line(ws, r, "Returned today, finished + scrap (g fine)",
         f"=SUMIFS(Karigar_Ledger!$J$5:$J${KG_LAST},{KR},$B$4)"
         f"+SUMIFS(Karigar_Ledger!$K$5:$K${KG_LAST},{KR},$B$4)", GRAM)
r = line(ws, r, "Wastage today (g fine)",
         f"=SUMIFS(Karigar_Ledger!$L$5:$L${KG_LAST},{KR},$B$4)", GRAM)
r = line(ws, r, "Cumulative out with karigars to date (g fine)",
         f"=SUMIFS(Karigar_Ledger!$H$5:$H${KG_LAST},{KR},\"<=\"&$B$4)"
         f"-SUMIFS(Karigar_Ledger!$J$5:$J${KG_LAST},{KR},\"<=\"&$B$4)"
         f"-SUMIFS(Karigar_Ledger!$K$5:$K${KG_LAST},{KR},\"<=\"&$B$4)", GRAM,
         "Unreturned metal. This is the shop's real exposure to its karigars.")
r = line(ws, r, "Lines over wastage norm today",
         f'=COUNTIFS({KR},$B$4,Karigar_Ledger!$P$5:$P${KG_LAST},"OVER NORM")',
         "#,##0")
r += 1

r = block(ws, r, "3.  RATE POSITION")
r = line(ws, r, "Spot 24K (₹/10g)",
         f"=SUMIFS(Rate_Sheet!$C$5:$C${RATE_LAST},{RR},$B$4)", INR)
r = line(ws, r, "Spot fine (₹/g)",
         f"=SUMIFS(Rate_Sheet!$D$5:$D${RATE_LAST},{RR},$B$4)", INR2)
r = line(ws, r, "Board rate implied fine (₹/g)",
         f"=SUMIFS(Rate_Sheet!$F$5:$F${RATE_LAST},{RR},$B$4)", INR2)
r = line(ws, r, "Drift (₹/g fine)",
         f"=SUMIFS(Rate_Sheet!$G$5:$G${RATE_LAST},{RR},$B$4)", INR2,
         "Board minus spot. Negative means selling below replacement cost.")
r = line(ws, r, "Fine gold sold today (g)",
         f"=SUMIFS(Sales_Register!$G$5:$G${SALES_LAST},{SR},$B$4)", GRAM)
rate_cost_row = r
r = line(ws, r, "Cost of the drift today (₹)",
         f"=B{r-1}*B{r-2}", INR,
         "Grams sold multiplied by the per-gram gap. Lost on today's tickets.")
r = line(ws, r, "Verdict",
         f'=IF(ABS(B{rate_cost_row-2})>Setup!$B$26,"REPRICE BOARD NOW","BOARD RATE OK")',
         "General")
r += 1

r = block(ws, r, "4.  MAKING CHARGE LEAKAGE")
r = line(ws, r, "MC list today (₹)",
         f"=SUMIFS(Sales_Register!$J$5:$J${SALES_LAST},{SR},$B$4)", INR)
mc_list_row = r - 1
r = line(ws, r, "MC realised today (₹)",
         f"=SUMIFS(Sales_Register!$L$5:$L${SALES_LAST},{SR},$B$4)", INR)
r = line(ws, r, "MC given away today (₹)",
         f"=SUMIFS(Sales_Register!$M$5:$M${SALES_LAST},{SR},$B$4)", INR,
         "Pure margin handed over at the counter.")
r = line(ws, r, "Realisation %",
         f"=IFERROR(B{mc_list_row+1}/B{mc_list_row},0)", PCT)
r = line(ws, r, "Floor", "=Setup!$B$27", PCT)
r = line(ws, r, "Verdict",
         f'=IF(B{r-2}<B{r-1},"DISCOUNTING TOO DEEP","OK")', "General")
r += 1

r = block(ws, r, "5.  DAY SUMMARY")
r = line(ws, r, "Invoices today",
         f"=COUNTIFS({SR},$B$4)", "#,##0")
r = line(ws, r, "Net invoiced value (₹)",
         f"=SUMIFS(Sales_Register!$P$5:$P${SALES_LAST},{SR},$B$4)", INR)
r = line(ws, r, "Metal value billed (₹)",
         f"=SUMIFS(Sales_Register!$I$5:$I${SALES_LAST},{SR},$B$4)", INR)
r = line(ws, r, "Old gold credited (₹)",
         f"=SUMIFS(Sales_Register!$O$5:$O${SALES_LAST},{SR},$B$4)", INR)

# =======================================================================
# 9. Monthly_Analysis
# =======================================================================
ws = wb.create_sheet("Monthly_Analysis")
title(ws, "Monthly Analysis", "Rollup by month. Leakage and variance are the two to watch.")
header_row(ws, 4, [
    "Month", "Invoices", "Fine gold sold (g)", "Net invoiced (₹)",
    "MC list (₹)", "MC realised (₹)", "MC given away (₹)", "Realisation %",
    "Gross variance (g fine)", "Days over tolerance", "Karigar wastage (g fine)",
    "Lines over norm",
], [11, 11, 16, 15, 14, 15, 15, 13, 17, 16, 17, 13])

months = sorted({dt.strftime("%Y-%m") for dt in business_days})
r = 5
for m in months:
    ws.cell(row=r, column=1, value=m).font = INPUT
    ws.cell(row=r, column=2,
            value=f"=COUNTIFS(Sales_Register!$B$5:$B${SALES_LAST},$A{r})"
            ).number_format = "#,##0"
    ws.cell(row=r, column=3,
            value=f"=SUMIFS(Sales_Register!$G$5:$G${SALES_LAST},"
                  f"Sales_Register!$B$5:$B${SALES_LAST},$A{r})").number_format = GRAM
    ws.cell(row=r, column=4,
            value=f"=SUMIFS(Sales_Register!$P$5:$P${SALES_LAST},"
                  f"Sales_Register!$B$5:$B${SALES_LAST},$A{r})").number_format = INR
    ws.cell(row=r, column=5,
            value=f"=SUMIFS(Sales_Register!$J$5:$J${SALES_LAST},"
                  f"Sales_Register!$B$5:$B${SALES_LAST},$A{r})").number_format = INR
    ws.cell(row=r, column=6,
            value=f"=SUMIFS(Sales_Register!$L$5:$L${SALES_LAST},"
                  f"Sales_Register!$B$5:$B${SALES_LAST},$A{r})").number_format = INR
    ws.cell(row=r, column=7,
            value=f"=SUMIFS(Sales_Register!$M$5:$M${SALES_LAST},"
                  f"Sales_Register!$B$5:$B${SALES_LAST},$A{r})").number_format = INR
    ws.cell(row=r, column=8, value=f"=IFERROR(F{r}/E{r},0)").number_format = PCT
    ws.cell(row=r, column=9,
            value=f"=SUMIFS(Metal_Ledger!$K$5:$K${METAL_LAST},"
                  f"Metal_Ledger!$B$5:$B${METAL_LAST},$A{r})").number_format = GRAM
    ws.cell(row=r, column=10,
            value=f'=COUNTIFS(Metal_Ledger!$B$5:$B${METAL_LAST},$A{r},'
                  f'Metal_Ledger!$N$5:$N${METAL_LAST},"INVESTIGATE")'
            ).number_format = "#,##0"
    ws.cell(row=r, column=11,
            value=f"=SUMIFS(Karigar_Ledger!$L$5:$L${KG_LAST},"
                  f"Karigar_Ledger!$B$5:$B${KG_LAST},$A{r})").number_format = GRAM
    ws.cell(row=r, column=12,
            value=f'=COUNTIFS(Karigar_Ledger!$B$5:$B${KG_LAST},$A{r},'
                  f'Karigar_Ledger!$P$5:$P${KG_LAST},"OVER NORM")'
            ).number_format = "#,##0"
    for col in range(1, 13):
        ws.cell(row=r, column=col).border = BOX
    r += 1
MONTH_LAST = r - 1

r += 1
ws.cell(row=r, column=1, value="Wastage by karigar, whole period").font = LBL
r += 1
header_row(ws, r, ["Karigar", "Trade", "Issued fine (g)", "Wastage fine (g)",
                   "Wastage %", "Norm %", "vs Norm", "Lines over norm",
                   "", "", "", ""])
ws.freeze_panes = None
r += 1
for kg in KARIGARS:
    ws.cell(row=r, column=1, value=kg).font = BODY
    ws.cell(row=r, column=2,
            value=f"=INDEX(Karigar_Ledger!$D$5:$D${KG_LAST},"
                  f"MATCH(A{r},Karigar_Ledger!$C$5:$C${KG_LAST},0))").font = BODY
    ws.cell(row=r, column=3,
            value=f"=SUMIFS(Karigar_Ledger!$H$5:$H${KG_LAST},"
                  f"Karigar_Ledger!$C$5:$C${KG_LAST},A{r})").number_format = GRAM
    ws.cell(row=r, column=4,
            value=f"=SUMIFS(Karigar_Ledger!$L$5:$L${KG_LAST},"
                  f"Karigar_Ledger!$C$5:$C${KG_LAST},A{r})").number_format = GRAM
    ws.cell(row=r, column=5, value=f"=IFERROR(D{r}/C{r},0)").number_format = PCT
    ws.cell(row=r, column=6,
            value=f"=INDEX(Setup!$B$19:$B$22,MATCH(B{r},Setup!$A$19:$A$22,0))"
            ).number_format = PCT
    ws.cell(row=r, column=7, value=f"=E{r}-F{r}").number_format = PCT
    ws.cell(row=r, column=8,
            value=f'=COUNTIFS(Karigar_Ledger!$C$5:$C${KG_LAST},A{r},'
                  f'Karigar_Ledger!$P$5:$P${KG_LAST},"OVER NORM")'
            ).number_format = "#,##0"
    for col in range(1, 9):
        ws.cell(row=r, column=col).border = BOX
    r += 1

r += 1
ws.cell(row=r, column=1,
        value="Wastage % is measured on fine gold issued, not gross weight. On a gross basis "
              "a karat mix shift looks identical to a loss."
        ).font = NOTE

wb.save(OUT)
print(f"wrote {OUT}")
print(f"rate rows 5-{RATE_LAST} | sku 5-{SKU_LAST} | karigar 5-{KG_LAST} | "
      f"sales 5-{SALES_LAST} | metal 5-{METAL_LAST} | months 5-{MONTH_LAST}")
