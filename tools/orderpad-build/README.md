# Rebuilding the Order Pad workbook

`clients/excel/current/samples/AutoTraderWeb-OrderPad.xlsm` is a **generated
file**. Nothing about it is edited by hand, and opening it in Excel to change a
column or a heading is the one thing not to do — the next rebuild discards it.

Everything the workbook contains is generated from three sources in this
repository:

| What | Where |
|---|---|
| The library modules it embeds | `direct/*.bas` |
| The Order Pad module | `clients/excel/current/samples/modules/AutoTraderOrderPad.bas` |
| Every sheet, heading, dropdown, colour rule and the Help text | `build-orderpad.py` in this directory |

## The three steps

Run them in this order, from the repository root. Steps 2 and 3 drive Excel
over COM, so they need Excel installed and run on Windows only.

```
python     tools/orderpad-build/build-orderpad.py
powershell -ExecutionPolicy Bypass -File tools/orderpad-build/add-buttons.ps1
powershell -ExecutionPolicy Bypass -File tools/orderpad-build/make-xlsm.ps1
```

1. **`build-orderpad.py`** writes `build/AutoTraderWeb-OrderPad.xlsx` — all five
   sheets, formats, dropdowns, conditional formatting and the Help sheet — plus
   `build/header-notes.json`, which is the single source of the heading notes.
   No buttons and no macros yet. Needs `openpyxl`.

2. **`add-buttons.ps1`** adds the seven Form control buttons and attaches the
   heading notes from that JSON. openpyxl cannot create a Form control, which is
   why this step needs Excel. It does *not* need the VBA trust setting.

3. **`make-xlsm.ps1`** imports the four modules, writes the `ThisWorkbook`
   close handler, and saves the result **over the shipped workbook** in
   `clients/excel/current/samples/`. This is the only step that needs *Trust
   access to the VBA project object model*; it turns that setting on, uses it,
   and puts it back.

`build/` is a scratch directory and is not committed. Only the finished `.xlsm`
ships.

## Check before shipping

Four gates. None of them needs Excel or a build, so all four can be run at any
time, on any machine:

```
python tools/orderpad-check/check-workbook.py    # embedded modules match this repo, key is the placeholder
python tools/orderpad-build/check-columns.py     # VBA column numbers match the headings; macro names resolve
python tools/orderpad-build/check-hints.py       # guidance is on the headings and nowhere else
python tools/vba-check/check-vba.py              # the library module sources themselves
```

The first three read the workbook that ships and each takes an optional path,
so a freshly built one can be checked *before* it replaces that. (`check-vba.py`
takes no argument: it reads `direct/*.bas` rather than the workbook.)

```
powershell -File tools/orderpad-build/make-xlsm.ps1 -Xlsm tools/orderpad-build/build/candidate.xlsm
python tools/orderpad-check/check-workbook.py tools/orderpad-build/build/candidate.xlsm
```

To look at the result rather than assert on it, `shot-excel.ps1` captures the
real Excel window. Use it for anything involving the buttons: `ExportAsFixedFormat`
does not render Form controls, so a PDF shows the workbook with none of them.

```
powershell -File tools/orderpad-build/shot-excel.ps1 -Sheet Orders
```

`check-vbom.ps1` reports whether *Trust access to the VBA project object model*
is currently on, and changes nothing. Step 3 turns that setting on, uses it and
puts it back — but a script that restores a setting and reports success can be
truthful about what it did and wrong about the outcome, because Excel writes its
own settings out as it shuts down. This is the independent second opinion:

```
powershell -File tools/orderpad-build/check-vbom.ps1
```

## Things that will bite

- **`FIRST_DATA_ROW` in the `.bas` must equal `HEADER_ROW` + 1.** Row 1 is the
  frozen button band and the headings sit on row 2, so the first order is row 3.
  The value is declared in three places — the `.bas`, `build-orderpad.py` and
  `add-buttons.ps1` — and `check-columns.py` asserts they agree.

- **VBA source must be plain ASCII.** Modules are stored in a code page rather
  than UTF-8, so any character above ASCII does not survive the import and the
  workbook silently stops matching its own source. `check-workbook.py` refuses
  such a file before it ever opens Excel.

- **A conditional format is not a normal fill.** Inside a `dxf`, Excel reads a
  solid fill's colour from `bgColor`. Setting only `fgColor` produces rules that
  fire against no colour, and the sheet simply looks unformatted. Set both.

- **The embedded config must carry the `<API_KEY>` placeholder.** `make-xlsm.ps1`
  refuses to run if `direct/AutoTraderConfig.bas` does not, and
  `check-workbook.py` fails if the finished workbook does not. A personalised
  download carries a real key; this repository is public.

- **`Application.OnTime` invokes its macro by string.** Renaming a macro without
  changing that literal leaves a timer that arms happily and then fires nothing.
  `check-columns.py` checks those strings resolve to a real `Public Sub`.
