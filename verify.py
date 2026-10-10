"""
Check a built workbook for formula errors and for the internal consistency of the
ledger identity.

openpyxl writes formulas as strings with no cached values, so a freshly built file
has to be recalculated by a spreadsheet engine before this can read results. If every
formula cell reads back as None, open the file once in Excel or LibreOffice and save
it, or run:

    soffice --headless --convert-to xlsx --outdir . gold_reconciliation_system.xlsx

Usage:  python verify.py gold_reconciliation_system.xlsx
Exit:   0 clean, 1 problems found
"""
import sys

from openpyxl import load_workbook

ERRORS = ("#VALUE!", "#REF!", "#DIV/0!", "#N/A", "#NAME?", "#NULL!", "#NUM!")
TOL = 0.001          # grams, for the ledger identity check


def main(path):
    wb = load_workbook(path, data_only=True)
    problems = []
    cells = formulas = 0

    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if c.value is None:
                    continue
                cells += 1
                if isinstance(c.value, str) and c.value.strip() in ERRORS:
                    problems.append(f"{ws.title}!{c.coordinate}  {c.value}")

    wb_f = load_workbook(path, data_only=False)
    for ws in wb_f.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    formulas += 1

    if cells == 0:
        print("no cached values found: recalculate the file first (see docstring)")
        return 1

    # The ledger identity must hold on every row, independently of the formulas:
    #   closing = opening + purchases - issued + returned - sold + old gold
    ml = wb["Metal_Ledger"]
    checked = broken = 0
    r = 5
    while ml.cell(r, 1).value is not None:
        v = [ml.cell(r, c).value for c in (3, 4, 5, 6, 7, 8, 9)]
        if all(isinstance(x, (int, float)) for x in v):
            opening, purch, issued, returned, sold, old, closing = v
            expected = opening + purch - issued + returned - sold + old
            checked += 1
            if abs(expected - closing) > TOL:
                broken += 1
                problems.append(
                    f"Metal_Ledger row {r}: closing {closing:.3f} != "
                    f"expected {expected:.3f} (diff {closing - expected:+.4f})"
                )
        r += 1

    print(f"sheets        {len(wb.worksheets)}")
    print(f"formulas      {formulas}")
    print(f"populated     {cells}")
    print(f"ledger rows   {checked} checked, {broken} broken")

    if problems:
        print(f"\nPROBLEMS ({len(problems)}):")
        for p in problems[:40]:
            print(f"  {p}")
        if len(problems) > 40:
            print(f"  ... and {len(problems) - 40} more")
        return 1

    print("\nclean: no formula errors, ledger identity holds on every row")
    return 0


if __name__ == "__main__":
    target = sys.argv[1] if len(sys.argv) > 1 else "gold_reconciliation_system.xlsx"
    sys.exit(main(target))
