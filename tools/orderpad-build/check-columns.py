"""Prove the Order Pad's VBA column numbers match the workbook's headers.

This is the same defect class that produced the holdings bug: code reading a
column by POSITION while the data's real layout sits somewhere else. Here both
halves are ours, so drift is entirely preventable -- the .bas declares the
positions, the .xlsx declares the headings, and nothing checks they agree
unless something does.

Run this after touching either file:

    python tools/orderpad-build/check-columns.py

With no argument it reads the workbook that SHIPS, so it needs neither Excel
nor a build. Pass a path to check a freshly built one before it replaces that:

    python tools/orderpad-build/check-columns.py tools/orderpad-build/build/AutoTraderWeb-OrderPad.xlsx

Needs `openpyxl` (pip install openpyxl).
"""
import re
import sys
from pathlib import Path

try:
    import openpyxl
except ImportError:
    print("openpyxl is not installed. Run: pip install openpyxl")
    sys.exit(2)

HEADER_ROW = 2

# tools/orderpad-build/check-columns.py -> the repository root is two levels up.
REPO = Path(__file__).resolve().parents[2]
BAS = REPO / "clients/excel/current/samples/modules/AutoTraderOrderPad.bas"
BOOK = (Path(sys.argv[1]) if len(sys.argv) > 1
        else REPO / "clients/excel/current/samples/AutoTraderWeb-OrderPad.xlsm")

# constant prefix -> sheet, and what each constant is meant to point at
EXPECTED = {
    "O_": ("Orders", {
        "O_ACCOUNT": "Account", "O_EXCHANGE": "Exchange", "O_SYMBOL": "Symbol",
        "O_SIDE": "Side", "O_QTY": "Quantity", "O_ORDERTYPE": "Order Type",
        "O_PRICE": "Price", "O_TRIGGER": "Trigger Price", "O_PRODUCT": "Product",
        "O_VARIETY": "Variety", "O_VALIDITY": "Validity", "O_TARGET": "Target",
        "O_STOPLOSS": "Stoploss", "O_TRAIL": "Trailing SL",
    }),
    "A_": ("Accounts", {
        "A_ACCOUNT": "Account", "A_INCLUDE": "Include",
        "A_QTY": "Quantity", "A_WEIGHT": "Weight",
    }),
    "R_": ("Results", {
        "R_TIME": "Time", "R_ROW": "Order Row", "R_ACCOUNT": "Account",
        "R_SYMBOL": "Symbol", "R_SIDE": "Side", "R_QTY": "Quantity",
        "R_ORDERID": "Order Id", "R_STATUS": "Status", "R_REASON": "Reason",
    }),
}


def main():
    if not BAS.exists():
        print("module source not found: %s" % BAS)
        return 2
    if not BOOK.exists():
        print("workbook not found: %s" % BOOK)
        return 2

    src = BAS.read_text(encoding="utf-8", errors="replace")
    declared = {
        m.group(1): int(m.group(2))
        for m in re.finditer(r"Private Const ([OAR]_[A-Z]+) As Long = (\d+)", src)
    }

    print("%s" % BOOK.name)
    wb = openpyxl.load_workbook(BOOK)
    problems = 0
    checked = 0

    for prefix, (sheet_name, mapping) in EXPECTED.items():
        # Row 1 is the frozen button band, so the headings live on row 2. The
        # .bas FIRST_DATA_ROW must be the row after that, and this asserts it
        # rather than trusting the two files to have been edited together.
        headers = [c.value for c in wb[sheet_name][HEADER_ROW]]
        for const, expected_header in mapping.items():
            if const not in declared:
                print("  %-14s NOT DECLARED in the .bas" % const)
                problems += 1
                continue
            index = declared[const]
            actual = headers[index - 1] if index - 1 < len(headers) else None
            checked += 1
            if actual != expected_header:
                print("  %-14s -> %s column %d is %r, expected %r"
                      % (const, sheet_name, index, actual, expected_header))
                problems += 1

    # every declared constant must be covered, or the check has a blind spot
    covered = {c for _p, (_s, m) in EXPECTED.items() for c in m}
    missed = sorted(set(declared) - covered)
    if missed:
        print("  constants declared but never checked: %s" % ", ".join(missed))
        problems += 1

    first_data = re.search(r"Private Const FIRST_DATA_ROW As Long = (\d+)", src)
    if not first_data:
        print("  FIRST_DATA_ROW is not declared in the .bas")
        problems += 1
    elif int(first_data.group(1)) != HEADER_ROW + 1:
        print("  FIRST_DATA_ROW is %s, but the headings are on row %d, so the "
              "first order row is %d" % (first_data.group(1), HEADER_ROW, HEADER_ROW + 1))
        problems += 1
    else:
        print("FIRST_DATA_ROW = %d, one row below the headings" % (HEADER_ROW + 1))

    print("%d column position(s) checked against the workbook" % checked)

    # --- macro names ---------------------------------------------------------
    #
    # Two ways a macro name goes wrong without anything noticing:
    #
    #   the Help sheet names one that was never written, so the reader types it
    #   into Alt+F8 and is told it does not exist;
    #
    #   Application.OnTime invokes a macro BY STRING, so a rename that misses
    #   that string leaves a timer that arms happily and then fires nothing.
    #
    # Neither is visible to a compiler. Both are visible here.
    subs = set(re.findall(r"^Public Sub (\w+)\(", src, re.M))
    print("%d public macro(s) in the module" % len(subs))

    helpsheet = wb["Help"]
    mentioned = set()
    for row in helpsheet.iter_rows():
        for cell in row:
            if cell.value:
                mentioned.update(re.findall(r"\bOrderPad\w+", str(cell.value)))

    for name in sorted(mentioned):
        if name not in subs:
            print("  Help sheet names %r, which is not a macro in the module" % name)
            problems += 1
    print("%d macro name(s) on the Help sheet checked" % len(mentioned))

    # Same line only. A pattern that crosses lines walks on to the next string
    # literal in the file and reports it as a macro name.
    for target in re.findall(r'Application\.OnTime[^\n"]*"([^"\n]+)"', src):
        bare = target.split(".")[-1]
        if bare not in subs:
            print("  Application.OnTime calls %r, which is not a Public Sub" % target)
            problems += 1

    # Public, but not for a person to run. Application.OnTime and the
    # Workbook_BeforeClose handler can only call a Public Sub, so these have to
    # be public even though putting them on the Help sheet would only invite
    # somebody to run them by hand. Every OTHER public macro must be documented.
    INTERNAL = {"OrderPadTimerFired", "OrderPadTick", "OrderPadStopTimersQuietly"}
    orphans = sorted(subs - mentioned - INTERNAL)
    if orphans:
        print("  macros the Help sheet never mentions: %s" % ", ".join(orphans))
        problems += 1
    if problems:
        print("FAIL: %d problem(s)" % problems)
        return 1
    print("PASS: every VBA column number lands on the heading it names")
    return 0


if __name__ == "__main__":
    sys.exit(main())
