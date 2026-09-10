# Put Form control buttons into the frozen band on row 1.
#
# openpyxl cannot make a Form control, so this is the one build step that needs
# Excel. It does NOT need "Trust access to the VBA project object model":
# adding a shape and setting its OnAction is ordinary object-model work. Only
# adding or removing VBA MODULES needs that setting.
#
# OnAction is stored as a macro NAME, so the button is happy to point at a macro
# that does not exist yet. In an .xlsx there are no macros at all and every
# button is inert -- which is exactly what we want for a layout preview.
#
# This is step 2 of 3. See README.md in this directory.
#
#     powershell -ExecutionPolicy Bypass -File tools/orderpad-build/add-buttons.ps1
#
# 🔴 `param` must be the FIRST statement in a PowerShell script -- only
# comments may come before it. Anything else, even a variable assignment, is a
# parse error. Paths are therefore derived below it, from $PSScriptRoot, which
# anchors them on the repository rather than on one machine's layout.
param(
    [string]$Path = "",
    [string]$Out  = ""
)

if ($Path -eq "") {
    $Path = Join-Path $PSScriptRoot "build\AutoTraderWeb-OrderPad.xlsx"
}
if (-not (Test-Path $Path)) {
    Write-Output "no workbook at $Path"
    Write-Output "Run build-orderpad.py first -- it writes the .xlsx this step decorates."
    exit 2
}

if ($Out -eq "") { $Out = $Path }

# sheet, caption, macro, slot (left to right)
$BUTTONS = @(
    @{ Sheet = "Orders";   Caption = "Check accounts"; Macro = "OrderPadCheckAccounts"; Slot = 0 },
    @{ Sheet = "Orders";   Caption = "Place orders";   Macro = "OrderPadPlaceOrders";   Slot = 1 },
    @{ Sheet = "Orders";   Caption = "Arm timer";      Macro = "OrderPadArmTimer";      Slot = 2 },
    @{ Sheet = "Orders";   Caption = "Cancel timer";   Macro = "OrderPadCancelTimer";   Slot = 3 },
    @{ Sheet = "Accounts"; Caption = "Check accounts"; Macro = "OrderPadCheckAccounts"; Slot = 0 },
    @{ Sheet = "Results";  Caption = "Refresh";        Macro = "OrderPadRefresh";       Slot = 0 },
    @{ Sheet = "Results";  Caption = "Clear results";  Macro = "OrderPadClearResults";  Slot = 1 }
)

$W = 104; $H = 24; $GAP = 6; $TOP = 4; $LEFT_MARGIN = 4

# Must match HEADER_ROW in build_orderpad.py and FIRST_DATA_ROW - 1 in the .bas.
$HEADER_ROW = 2

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $wb = $excel.Workbooks.Open($Path)

    foreach ($sheetName in ($BUTTONS | ForEach-Object { $_.Sheet } | Select-Object -Unique)) {
        $ws = $wb.Worksheets.Item($sheetName)
        # Remove any button from a previous run so this stays repeatable.
        for ($i = $ws.Buttons().Count; $i -ge 1; $i--) { $ws.Buttons().Item($i).Delete() }
    }

    foreach ($b in $BUTTONS) {
        $ws = $wb.Worksheets.Item($b.Sheet)
        # 🔴 POWERSHELL VARIABLE NAMES ARE CASE-INSENSITIVE. $left and $LEFT are
        # ONE variable, so a loop that computes $left from a $LEFT margin
        # overwrites the margin on every pass and each button lands further out
        # than the last. It reads as a layout bug and is really a name
        # collision. Keep the names distinct -- hence $LEFT_MARGIN and $x.
        $x = $LEFT_MARGIN + $b.Slot * ($W + $GAP)
        $btn = $ws.Buttons().Add($x, $TOP, $W, $H)
        $btn.Caption = $b.Caption
        $btn.OnAction = $b.Macro
        $btn.Name = "btn" + $b.Macro
        $btn.Placement = 3          # xlFreeFloating: never move or size with cells
        Write-Output ("{0,-10} {1,-16} -> {2}" -f $b.Sheet, $b.Caption, $b.Macro)
    }

    # --- header notes -------------------------------------------------------
    #
    # 🔴 A comment anchored on a frozen row is CUT BY THE PANE SPLIT LINE, which
    # Excel draws last, over floating shapes, so a grey rule strikes through the
    # note. Resizing cannot avoid it. The fix is to leave the box where Excel
    # puts it but push it DOWN, so it begins below the split rather than
    # straddling it.
    # Excel draws a leader line back to the heading, so it still reads as that
    # column's note.
    $notesFile = Join-Path (Split-Path $Path -Parent) "header-notes.json"
    if (Test-Path $notesFile) {
        $notes = Get-Content $notesFile -Raw | ConvertFrom-Json
        $attached = 0
        foreach ($sheetName in $notes.PSObject.Properties.Name) {
            $ws = $wb.Worksheets.Item($sheetName)
            $belowSplit = $ws.Rows.Item($HEADER_ROW + 1).Top + 4
            foreach ($col in $notes.$sheetName.PSObject.Properties.Name) {
                $cell = $ws.Range($col + [string]$HEADER_ROW)
                if ($null -ne $cell.Comment) { $cell.Comment.Delete() }
                $c = $cell.AddComment($notes.$sheetName.$col)
                $c.Shape.TextFrame.AutoSize = $true
                # AutoSize can produce one very long line; cap the width and let
                # it wrap, then let AutoSize settle the height.
                if ($c.Shape.Width -gt 260) {
                    $c.Shape.TextFrame.AutoSize = $false
                    $c.Shape.Width = 260
                    $c.Shape.TextFrame.AutoSize = $true
                }
                $c.Shape.Top = $belowSplit
                $c.Visible = $false
                $attached++
            }
        }
        Write-Output "attached $attached header note(s), all clear of the pane split"
    } else {
        Write-Output "NO header-notes.json beside the workbook -- no notes attached"
    }

    if ($Out -ne $Path) {
        $wb.SaveAs($Out)
    } else {
        $wb.Save()
    }
    Write-Output "saved $Out"
    $wb.Close($false)
}
finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
}
