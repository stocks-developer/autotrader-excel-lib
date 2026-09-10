"""Build the Order Pad workbook.

openpyxl cannot create a macro project, so this writes the SHEETS and the
workbook is turned into an .xlsm in Excel once, with the modules imported.
Everything that can be built without Excel is built here: headers, dropdowns,
widths, frozen panes, tab colours, hover hints, the Results colouring and a
Help sheet laid out to be read rather than scanned.

🔴 The Orders sheet ships EMPTY, on purpose. Every one of the four older tools
shipped with sample orders already typed into the grid -- SBIN, INFY, a NIFTY
future -- so anyone who opened one and pressed the button placed them for real.
The examples live on the Help sheet instead, as text nothing will ever read.

🔴 HEADERS STAY ON ROW 1. The module reads data from FIRST_DATA_ROW = 2 and has
been proven live at that offset. An instruction band across the top would push
every grid down one row and put that proof back in question for something that
scrolls out of sight anyway. Guidance goes into hover notes on the headers, the
tab colours and the Help sheet instead.

🔴 Configuration and Help are the LAST two tabs, and the file opens on Orders.

This is step 1 of 3. It writes an .xlsx with no buttons and no macros; the
other two steps add those. See README.md in this directory for the whole
sequence and for what to check before shipping the result.

    python tools/orderpad-build/build-orderpad.py

Needs `openpyxl` (pip install openpyxl).
"""
import json
import math
import pathlib
import sys

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

# Everything is anchored on the repository, never on one machine's layout, so
# this runs from any clone. tools/orderpad-build/build-orderpad.py -> the
# repository root is two levels up.
REPO = pathlib.Path(__file__).resolve().parents[2]
BUILD = REPO / "tools" / "orderpad-build" / "build"

# The intermediate .xlsx is a BUILD ARTEFACT and is not committed: only the
# finished .xlsm ships. It lives under build/ so a half-finished workbook can
# never be mistaken for the one in samples/.
OUT = str(BUILD / "AutoTraderWeb-OrderPad.xlsx")

# --- palette -----------------------------------------------------------------
BRAND = "1F3864"
BAND = "D9E2F3"
GREY = "F2F2F2"

HEAD_FILL = PatternFill("solid", fgColor=BRAND)
HEAD_FONT = Font(color="FFFFFF", bold=True)
BAND_FILL = PatternFill("solid", fgColor=BAND)

TITLE_FONT = Font(bold=True, size=18, color=BRAND)
SUB_FONT = Font(italic=True, size=11, color="666666")
SECTION_FONT = Font(bold=True, size=12, color=BRAND)
LABEL_FONT = Font(bold=True, size=11)
BODY_FONT = Font(size=11)
NOTE_FONT = Font(italic=True, size=10, color="666666")
MONO_FONT = Font(name="Consolas", size=10)

# Row 1 is the frozen button band, row 2 the headers, row 3 the first order.
# 🔴 FIRST_DATA_ROW in AutoTraderOrderPad.bas MUST equal HEADER_ROW + 1.
NOTES = {}
BAND_HEIGHT = 34
HEADER_ROW = 2

WRAP = Alignment(wrap_text=True, vertical="top")
TOP = Alignment(vertical="top")

# Tab colours group the sheets by what they are for: you type into the blue
# ones, the tool writes the grey one, and the last two are set-up and reading.
TABS = {
    "Orders": BRAND,
    "Accounts": "2E75B6",
    "Results": "808080",
    "Configuration": "BF8F00",
    "Help": "548235",
}

# --- grids -------------------------------------------------------------------
# (heading, column width, hover note shown when you point at the heading)
ORDERS = [
    ("Account", 18, "Leave this blank to send the order to every account you "
                    "ticked on the Accounts sheet. Type one account name here "
                    "and the order goes to that account only."),
    ("Exchange", 10, "Pick from the list."),
    ("Symbol", 30, "The broker independent symbol. Look it up with the "
                   "instrument search on our website if you are not sure."),
    ("Side", 8, "BUY or SELL."),
    ("Quantity", 10, "How many. What each account gets also depends on the "
                     "Allocation mode on the Configuration sheet."),
    ("Order Type", 13, "MARKET takes the going price. LIMIT waits for the "
                       "price you put in the Price column."),
    ("Price", 10, "Leave this 0 for a MARKET order."),
    ("Trigger Price", 13, "Only for stop loss orders."),
    ("Product", 12, "INTRADAY is squared off the same day. DELIVERY and "
                    "NORMAL are carried forward."),
    ("Variety", 10, "REGULAR unless you are placing a bracket or cover order."),
    ("Validity", 10, "DAY stays until the market closes. IOC is filled at once "
                     "or cancelled."),
    ("Target", 10, "Bracket orders only. Leave blank otherwise."),
    ("Stoploss", 10, "Bracket orders only. Leave blank otherwise."),
    ("Trailing SL", 12, "Bracket orders only. Leave blank otherwise."),
]

ACCOUNTS = [
    ("Account", 22, "Your pseudo account name, exactly as it appears in your "
                    "AutoTrader Web account."),
    ("Include", 10, "Y to trade this account, N to skip it."),
    ("Quantity", 12, "Used only when Allocation mode is PER ACCOUNT."),
    ("Weight", 10, "Used only when Allocation mode is SPLIT. A share, not a "
                   "percentage: 2 and 1 means twice as much to the first."),
]

RESULTS = [
    ("Time", 18, "When the order was sent."),
    ("Order Row", 11, "Which row on the Orders sheet this came from."),
    ("Account", 20, None),
    ("Symbol", 28, None),
    ("Side", 8, None),
    ("Quantity", 10, None),
    ("Order Id", 22, "Your broker's own id. Look this up in your order book."),
    ("Status", 14, "Green means the order is with your broker. Red means it is "
                   "not in the market, and the Reason says why. Amber means it "
                   "MAY be live and you must check your order book."),
    ("Reason", 60, "Your broker's own words, when there are any."),
]

LISTS = {
    "B": '"NSE,BSE,MCX"',           # Exchange
    "D": '"BUY,SELL"',              # Side
    "F": '"MARKET,LIMIT,STOP_LOSS,SL_MARKET"',
    "I": '"INTRADAY,DELIVERY,NORMAL"',
    "J": '"REGULAR,BO,CO"',
    "K": '"DAY,IOC"',
}

EXAMPLE_ORDERS = [
    "Exchange  Symbol                      Side  Qty  Type    Price   Product",
    "NSE       SBIN                        BUY   1    MARKET  0       INTRADAY",
    "NSE       INFY                        SELL  5    LIMIT   1450.5  DELIVERY",
    "NSE       NIFTY_25-SEP-2026_CE_24000  BUY   75   MARKET  0       NORMAL",
    "MCX       CRUDEOIL_19-SEP-2026_FUT    BUY   1    LIMIT   5210    NORMAL",
]

# --- the Help sheet, as content rather than as formatting --------------------
HELP = [
    ("title", "AutoTrader Web  -  Order Pad"),
    ("sub", "Place orders into one account or many, and see what happened to each one."),
    ("blank",),

    ("section", "Start here"),
    ("step", "1", "Put your API key on the Configuration sheet. You will find "
                  "your key in your account settings at webx.stocksdeveloper.in."),
    ("step", "2", "Go to the Accounts sheet. Type your account names and put Y "
                  "beside each one you want to trade."),
    ("step", "3", "Run OrderPadCheckAccounts. It reads each account and shows "
                  "you its free margin. It places nothing at all, so it is safe "
                  "to run any time, and it catches a wrong account name before "
                  "you send fifty orders rather than after."),
    ("step", "4", "Go to the Orders sheet and type your orders, one per row."),
    ("step", "5", "Run OrderPadPlaceOrders. It tells you how many real orders "
                  "that is and asks you to confirm before anything is sent."),
    ("step", "6", "Read the Results sheet. Every order is there with its id, "
                  "its status and its reason."),
    ("note", "The buttons for these sit along the top of the Orders, Accounts "
             "and Results sheets, and stay put as you scroll. You can also run "
             "any of them by pressing Alt+F8 and picking the name."),
    ("blank",),

    ("section", "The sheets"),
    ("row", "Orders", "What you want to place. One row per order."),
    ("row", "Accounts", "Which accounts to place for, and how much each gets."),
    ("row", "Results", "What happened. The tool writes this. Newest run at the bottom."),
    ("row", "Configuration", "Your API key, the allocation mode, and a time to place at."),
    ("row", "Help", "This sheet."),
    ("blank",),

    ("section", "What each macro does"),
    ("row", "OrderPadCheckAccounts", "Reads each ticked account and shows its "
                                     "free margin. Places nothing. Run it first."),
    ("row", "OrderPadPlaceOrders", "Asks you to confirm, then places everything "
                                   "on the Orders sheet."),
    ("row", "OrderPadRefresh", "Re-reads Results from your broker's order book."),
    ("row", "OrderPadArmTimer", "Places automatically at the time on the "
                                "Configuration sheet."),
    ("row", "OrderPadCancelTimer", "Cancels a timer you armed."),
    ("row", "OrderPadClearResults", "Empties the Results sheet. This clears the "
                                    "sheet only. It cancels nothing at your broker."),
    ("blank",),

    ("section", "Sending one order to many accounts"),
    ("body", "Leave the Account column on the Orders sheet blank. That order "
             "then goes to every account you ticked on the Accounts sheet. Fill "
             "the Account column in and that row goes to that one account only, "
             "whatever else is ticked."),
    ("body", "How much each account gets is set by Allocation mode on the "
             "Configuration sheet:"),
    ("row", "SAME", "Every account gets the quantity on the order row."),
    ("row", "PER ACCOUNT", "Every account gets its own quantity, taken from the "
                           "Accounts sheet."),
    ("row", "SPLIT", "The order quantity is shared out across accounts in "
                     "proportion to Weight. Rounding leftovers go to the first "
                     "account, so the pieces always add back up to what you asked for."),
    ("blank",),

    ("section", "Example orders"),
    ("body", "Copy a line into the Orders sheet and change it. Nothing on this "
             "sheet is ever read by the tool, so the Orders sheet arrives empty "
             "and opening this file can never place somebody else's example."),
    ("mono", EXAMPLE_ORDERS),
    ("note", "Not sure of a symbol? Use the instrument search on our website. "
             "Always use the broker independent symbol."),
    ("blank",),

    ("section", "What Results tells you"),
    ("row", "An order id", "Your broker took it. That id is the one in your order "
                          "book. The row is green."),
    ("row", "REJECTED", "It reached your broker and your broker refused it. The "
                       "Reason column has your broker's own words. The row is red, "
                       "because there is no order in the market."),
    ("row", "NOT PLACED", "Nothing reached your broker at all. Safe to fix and run "
                         "again. The row is red."),
    ("row", "UNCONFIRMED", "It reached your broker and your broker never "
                           "answered. It MAY be live. Do not send it again "
                           "until you have checked your order book. The row is amber."),
    ("body", "NOT PLACED is rare. UNCONFIRMED is the one that matters. Treating "
             "an unconfirmed order as a failure is how the same order gets "
             "placed twice."),
    ("body", "Your broker's order book takes a few seconds to catch up, so a "
             "status read straight after placing is often still the old one. "
             "Run OrderPadRefresh a few seconds later."),
    ("blank",),

    ("section", "Keep your key safe"),
    ("body", "Anyone who has your API key can place orders in your accounts. "
             "Once your key is on the Configuration sheet, do not e-mail this "
             "workbook and do not attach it to a support request."),
]


def band(ws, columns, caption):
    """Reserve the frozen top row the buttons are placed into.

    The buttons are Form controls, which openpyxl cannot create -- add_buttons.ps1
    puts them here over Excel COM. This paints the strip and leaves the room, so
    the workbook looks right even before that step has run.
    """
    width = max(1, len(columns))
    for i in range(1, width + 1):
        ws.cell(row=1, column=i).fill = PatternFill("solid", fgColor=BAND)
    # Only where there are no buttons. On a sheet that has them the caption
    # lands underneath them and is clipped, which is what the first build did.
    if caption:
        label = ws.cell(row=1, column=1, value=caption)
        label.font = Font(bold=True, size=10, color=BRAND)
        label.alignment = Alignment(vertical="center", horizontal="left", indent=1)
    ws.row_dimensions[1].height = BAND_HEIGHT


def head(ws, columns, hint_rows=0, lists=None, header_row=HEADER_ROW):
    """Header row, widths, frozen top row, and the per-column guidance.

    🔴 THE GUIDANCE IS NOT A CELL COMMENT. A comment anchored on the header row
    hangs down across the frozen-pane split, and Excel draws that split line
    LAST, straight through the note -- a grey rule through the middle of the
    text, which is what shipped in the first build. Comments on a frozen header
    row cannot avoid it: the note always drops across the boundary.

    Data validation input messages have no such problem. They are chrome, not a
    shape on the sheet, and they appear when someone CLICKS the cell they are
    about to type in -- which is where the help is actually wanted, rather than
    on a heading they have to think to hover over.
    """
    lists = lists or {}
    for i, col in enumerate(columns, start=1):
        name, width = col[0], col[1]
        note = col[2] if len(col) > 2 else None
        letter = get_column_letter(i)

        cell = ws.cell(row=header_row, column=i, value=name)
        cell.fill = HEAD_FILL
        cell.font = HEAD_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[letter].width = width

        # 🔴 THE GUIDANCE BELONGS ON THE HEADING, NOT ON EVERY CELL. Putting it
        # in a data-validation input message meant a box popped up on every one
        # of five hundred rows the moment it was selected, which is noise, not
        # help. Only the list dropdowns stay as validation. The notes are handed
        # to add_buttons.ps1, which attaches them to the HEADING cells as real
        # comments and positions each box clear of the frozen-pane split -- the
        # line that struck through the first attempt.
        if note:
            NOTES.setdefault(ws.title, {})[letter] = note

        formula = lists.get(letter)
        if not hint_rows or formula is None:
            continue

        dv = DataValidation(type="list", formula1=formula, allow_blank=True)
        dv.error = "Pick one from the list."
        dv.errorTitle = "Not a valid value"
        dv.showErrorMessage = True
        ws.add_data_validation(dv)
        dv.add("%s%d:%s%d" % (letter, header_row + 1, letter, hint_rows))

    ws.freeze_panes = "A%d" % (header_row + 1)
    ws.row_dimensions[header_row].height = 24


def wrapped_height(text, chars_per_line, base=15):
    return base * max(1, math.ceil(len(text) / float(chars_per_line)))


def build_help(ws):
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 26
    ws.column_dimensions["C"].width = 78

    r = 2
    for item in HELP:
        kind = item[0]

        if kind == "blank":
            ws.row_dimensions[r].height = 8
            r += 1

        elif kind == "title":
            c = ws.cell(row=r, column=2, value=item[1])
            c.font = TITLE_FONT
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
            ws.row_dimensions[r].height = 30
            r += 1

        elif kind == "sub":
            c = ws.cell(row=r, column=2, value=item[1])
            c.font = SUB_FONT
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
            r += 1

        elif kind == "section":
            for col in (2, 3):
                ws.cell(row=r, column=col).fill = BAND_FILL
            c = ws.cell(row=r, column=2, value=item[1])
            c.font = SECTION_FONT
            c.alignment = Alignment(vertical="center")
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
            ws.row_dimensions[r].height = 24
            r += 1

        elif kind == "step":
            ws.cell(row=r, column=2, value=item[1]).font = LABEL_FONT
            ws.cell(row=r, column=2).alignment = Alignment(horizontal="right", vertical="top")
            c = ws.cell(row=r, column=3, value=item[2])
            c.font = BODY_FONT
            c.alignment = WRAP
            ws.row_dimensions[r].height = wrapped_height(item[2], 92)
            r += 1

        elif kind == "row":
            ws.cell(row=r, column=2, value=item[1]).font = LABEL_FONT
            ws.cell(row=r, column=2).alignment = TOP
            c = ws.cell(row=r, column=3, value=item[2])
            c.font = BODY_FONT
            c.alignment = WRAP
            ws.row_dimensions[r].height = wrapped_height(item[2], 92)
            r += 1

        elif kind == "body":
            c = ws.cell(row=r, column=3, value=item[1])
            c.font = BODY_FONT
            c.alignment = WRAP
            ws.row_dimensions[r].height = wrapped_height(item[1], 92)
            r += 1

        elif kind == "note":
            c = ws.cell(row=r, column=3, value=item[1])
            c.font = NOTE_FONT
            c.alignment = WRAP
            ws.row_dimensions[r].height = wrapped_height(item[1], 100, base=13)
            r += 1

        elif kind == "mono":
            for line in item[1]:
                c = ws.cell(row=r, column=3, value=line)
                c.font = MONO_FONT
                c.alignment = TOP
                ws.row_dimensions[r].height = 14
                r += 1

        else:
            raise SystemExit("unknown Help item: %r" % (kind,))


def main():
    out = sys.argv[1] if len(sys.argv) > 1 else OUT

    # 🔴 REFUSE TO BUILD WHEN THE TARGET IS OPEN, LOUDLY. Excel holds an
    # exclusive lock on an open workbook. Falling back to a second filename
    # would leave the CANONICAL workbook stale while the change went somewhere
    # else -- so the file being looked at would not be the file being changed,
    # and a fix that worked would look as though it had not.
    # One name, or a clear error. Never a silent second copy.
    target = pathlib.Path(out)
    target.parent.mkdir(parents=True, exist_ok=True)
    if target.exists():
        try:
            with open(target, "r+b"):
                pass
        except PermissionError:
            print("CANNOT WRITE %s -- it is open in Excel." % target)
            print("Close the workbook and run this again.")
            return 2

    wb = Workbook()

    # Orders is the sheet you work in, so the file opens on it.
    orders = wb.active
    orders.title = "Orders"
    band(orders, ORDERS, None)
    head(orders, ORDERS, hint_rows=500, lists=LISTS)

    accounts = wb.create_sheet("Accounts")
    band(accounts, ACCOUNTS, None)
    head(accounts, ACCOUNTS, hint_rows=200, lists={"B": '"Y,N"'})

    # Results is written by the tool, so it gets no input prompts: a message
    # inviting someone to type would be inviting them to type over an answer.
    results = wb.create_sheet("Results")
    band(results, RESULTS, None)
    head(results, RESULTS)

    # Ship the columns already formatted. Without this Excel picks its own
    # format for whatever the macro writes: a date loses its seconds, and an
    # order id made of digits becomes 2.6091E+13.
    for row in range(HEADER_ROW + 1, 501):
        results.cell(row=row, column=1).number_format = "yyyy-mm-dd hh:mm:ss"
        results.cell(row=row, column=7).number_format = "@"

    # The tool writes this sheet, so shade it and colour the answer. A new user
    # should be able to tell at a glance which rows need them to do something.
    #
    # 🔴 A CONDITIONAL FORMAT IS NOT A NORMAL FILL. Inside a differential format
    # (dxf) Excel reads a solid fill's colour from bgColor, and openpyxl's
    # PatternFill("solid", fgColor=...) writes fgColor ONLY -- so the rules fire
    # against no colour at all and every row stays white. Nothing warns: the
    # rules are in the file, the dxfs are in the file, and the sheet looks
    # unformatted. Caught by rendering it (check_results_colour.py), never by
    # reading the code. Set BOTH colours.
    #
    def dxf_fill(colour):
        return PatternFill(fill_type="solid", start_color=colour, end_color=colour)

    # Colour by what the row MEANS to the person reading it, not by how far the
    # request got. A REJECTED order reached the broker and was refused, so by
    # any technical measure it succeeded -- but to the reader it is an order
    # that is not in the market, which is the same thing NOT PLACED means. Red
    # for both. The Reason column is what tells them apart.
    #
    # First match wins (stopIfTrue), so the last rule can simply be "anything
    # else that is filled in" rather than a growing list of exclusions.
    for rule, colour, stop in (
        ('OR(EXACT($H3,"NOT PLACED"),EXACT($H3,"REJECTED"))', "FFFFC7CE", True),
        ('EXACT($H3,"UNCONFIRMED")', "FFFFE699", True),
        ('EXACT($H3,"CANCELLED")', "FFE7E6E6", True),
        ('$H3<>""', "FFC6EFCE", False),
    ):
        results.conditional_formatting.add(
            "A%d:I500" % (HEADER_ROW + 1),
            FormulaRule(formula=[rule], fill=dxf_fill(colour), stopIfTrue=stop),
        )

    settings = wb.create_sheet("Configuration")
    band(settings, [0, 0, 0], "Fill in the Value column:")
    head(settings, [
        ("Setting", 20, None),
        ("Value", 34, None),
        ("What it does", 74, None),
    ])
    rows = [
        ("API key", "",
         "In AutoTrader Web, go to Settings > Security. Your key is hidden, so "
         "click Reveal, type your AutoTrader password, and then click Copy. "
         "Paste it here. Anyone who has this key can place orders in your "
         "accounts, so treat it like a password and do not send this file to "
         "anyone once it is filled in."),
        ("Allocation mode", "SAME",
         "How much each account gets when you leave the Account column on an "
         "order row blank. SAME gives every account the quantity on the order "
         "row. PER ACCOUNT gives each account its own quantity from the "
         "Accounts sheet. SPLIT shares one quantity out in proportion to Weight."),
        ("Place at", "",
         "A time such as 09:20:00. Run OrderPadArmTimer and your orders go out "
         "when that time arrives. Leave this blank if you place by hand."),
    ]
    # The Value column is the only place on this sheet anyone types. Mark it,
    # so a new user is not left guessing which of three columns is theirs.
    input_fill = PatternFill("solid", fgColor="FFF2CC")
    edge = Side(style="thin", color="BF8F00")
    input_border = Border(left=edge, right=edge, top=edge, bottom=edge)
    for r, (a, b, c) in enumerate(rows, start=HEADER_ROW + 1):
        settings.cell(row=r, column=1, value=a).font = LABEL_FONT
        settings.cell(row=r, column=1).alignment = TOP
        value = settings.cell(row=r, column=2, value=b)
        value.alignment = TOP
        value.fill = input_fill
        value.border = input_border
        settings.cell(row=r, column=3, value=c).alignment = WRAP
        settings.row_dimensions[r].height = wrapped_height(c, 88)
    dv = DataValidation(type="list", formula1='"SAME,PER ACCOUNT,SPLIT"', allow_blank=True)
    settings.add_data_validation(dv)
    dv.add("B%d" % (HEADER_ROW + 2))

    helpsheet = wb.create_sheet("Help")
    build_help(helpsheet)

    for name, colour in TABS.items():
        wb[name].sheet_properties.tabColor = colour

    wb.active = 0
    wb.save(out)

    # Single source of truth for the header notes: written here, attached by
    # the COM step. Duplicating them in the PowerShell script would guarantee
    # the two drifted apart.
    notes_path = pathlib.Path(out).parent / "header-notes.json"
    notes_path.write_text(json.dumps(NOTES, indent=2), encoding="utf-8")

    print("wrote %s" % out)
    print("wrote %s (%d sheet(s))" % (notes_path, len(NOTES)))
    print("sheets: %s" % ", ".join(wb.sheetnames))
    print("opens on: %s" % wb.sheetnames[wb.active if isinstance(wb.active, int) else 0])
    return 0


if __name__ == "__main__":
    sys.exit(main())
