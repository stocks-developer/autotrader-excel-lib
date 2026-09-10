# Turn the built .xlsx into the shippable, zero-import .xlsm.
#
# This is the ONE step that needs "Trust access to the VBA project object
# model". Adding buttons did not; importing MODULES does. The switch is turned
# on, used, and turned off again in the same run, and because AccessVBOM is
# normally ABSENT rather than 0, undoing it means REMOVING the value, not
# setting it to zero -- setting 0 would leave a footprint that was never there.
#
# 🔴 THE MODULES COME FROM THE REPO, NEVER FROM A DOWNLOAD. The repo copy of
# AutoTraderConfig.bas carries the <API_KEY> placeholder by design; a
# personalised download carries a REAL key, and this workbook is destined for a
# PUBLIC repository. Getting this backwards would publish a live trading key.
#
# This is step 3 of 3. See README.md in this directory.
#
#     powershell -ExecutionPolicy Bypass -File tools/orderpad-build/make-xlsm.ps1
#
# 🔴 BY DEFAULT THIS OVERWRITES THE SHIPPED WORKBOOK in
# clients/excel/current/samples/. That is deliberate -- one workbook, one name.
# Writing somewhere else "for safety" leaves a stale file at the canonical
# path, which is worse. Pass -Xlsm when you genuinely want a copy elsewhere,
# to compare against what ships before replacing it.
#
# 🔴 `param` must be the FIRST statement in a PowerShell script -- only
# comments may come before it. Paths are derived below it, from $PSScriptRoot,
# which anchors them on the repository rather than on one machine's layout.
param(
    [string]$Xlsx    = "",
    [string]$Xlsm    = "",
    [string]$LibDir  = "",
    [string]$PadBas  = ""
)

# tools/orderpad-build/ -> the repository root is two levels up.
$Repo = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent

if ($Xlsx   -eq "") { $Xlsx   = Join-Path $PSScriptRoot "build\AutoTraderWeb-OrderPad.xlsx" }
if ($Xlsm   -eq "") { $Xlsm   = Join-Path $Repo "clients\excel\current\samples\AutoTraderWeb-OrderPad.xlsm" }
if ($LibDir -eq "") { $LibDir = Join-Path $Repo "direct" }
if ($PadBas -eq "") { $PadBas = Join-Path $Repo "clients\excel\current\samples\modules\AutoTraderOrderPad.bas" }

if (-not (Test-Path $Xlsx)) {
    Write-Output "no workbook at $Xlsx"
    Write-Output "Run build-orderpad.py and add-buttons.ps1 first."
    exit 2
}

$MODULES = @(
    (Join-Path $LibDir "AutoTraderConfig.bas"),
    (Join-Path $LibDir "AutoTraderWebDirect.bas"),
    (Join-Path $LibDir "AutoTraderClientDirect.bas"),
    $PadBas
)

foreach ($m in $MODULES) {
    if (-not (Test-Path $m)) { throw "missing module: $m" }
}

# Refuse outright if the config we are about to embed is not the placeholder one.
$config = Get-Content (Join-Path $LibDir "AutoTraderConfig.bas") -Raw
if ($config -notmatch '<API_KEY>') {
    throw "AutoTraderConfig.bas does not contain the <API_KEY> placeholder. Refusing to embed a real key."
}
Write-Output "config check: <API_KEY> placeholder present, safe to embed"

# 🔴 Refuse loudly if the target is open. Silently writing somewhere else
# would leave a stale workbook at the canonical path.
if (Test-Path $Xlsm) {
    try {
        $fs = [System.IO.File]::Open($Xlsm, 'Open', 'ReadWrite', 'None')
        $fs.Close()
    } catch {
        Write-Output "CANNOT WRITE $Xlsm -- it is open in Excel."
        Write-Output "Close the workbook and run this again."
        exit 2
    }
}

$regPath = "HKCU:\Software\Microsoft\Office\16.0\Excel\Security"
$had = (Get-ItemProperty -Path $regPath -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM
Set-ItemProperty -Path $regPath -Name AccessVBOM -Value 1 -Type DWord
Write-Output "AccessVBOM enabled (was: $(if ($null -eq $had) { 'ABSENT' } else { $had }))"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $wb = $excel.Workbooks.Open($Xlsx)

    # Remove anything already there, so a re-run is not additive.
    $proj = $wb.VBProject
    for ($i = $proj.VBComponents.Count; $i -ge 1; $i--) {
        $c = $proj.VBComponents.Item($i)
        if ($c.Type -eq 1) { $proj.VBComponents.Remove($c) }   # 1 = standard module
    }

    foreach ($m in $MODULES) {
        $proj.VBComponents.Import($m) | Out-Null
        Write-Output ("imported {0}" -f (Split-Path $m -Leaf))
    }

    Write-Output ("modules now in the project: {0}" -f (
        ($proj.VBComponents | Where-Object { $_.Type -eq 1 } | ForEach-Object { $_.Name }) -join ", "))

    # A pending Application.OnTime can make Excel REOPEN this workbook by itself
    # to run the macro it promised to run. With the countdown rescheduling every
    # second that is close to certain, and a trading workbook reopening on its
    # own is alarming whether or not it places anything. The handler has to live
    # in the ThisWorkbook document module, which cannot be imported like a
    # standard module, so it is written in here.
    $thisWb = $proj.VBComponents.Item("ThisWorkbook").CodeModule
    if ($thisWb.CountOfLines -gt 0) { $thisWb.DeleteLines(1, $thisWb.CountOfLines) }
    $handler = @"
Option Explicit

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Clear any armed timer so Excel does not reopen this file on its own.
    OrderPadStopTimersQuietly
End Sub
"@
    $thisWb.AddFromString($handler)
    Write-Output "ThisWorkbook: Workbook_BeforeClose added (cancels pending timers)"

    if (Test-Path $Xlsm) { Remove-Item $Xlsm -Force }
    $wb.SaveAs($Xlsm, 52)          # 52 = xlOpenXMLWorkbookMacroEnabled
    Write-Output "saved $Xlsm"
    $wb.Close($false)
}
finally {
    $pid_ = $excel.Hwnd   # touch it before quitting, to be sure we had it
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    # 🔴 WAIT FOR EXCEL TO ACTUALLY EXIT BEFORE PUTTING THE SETTING BACK.
    # Excel writes its settings out as it shuts down, so a removal issued while
    # it is still terminating gets overwritten by the value Excel is holding in
    # memory -- which is the 1 we just set. A script that removes the value and
    # says so can therefore be truthful about what it DID and wrong about the
    # outcome, leaving the setting switched on. So: restore, then VERIFY by
    # reading it back, and say so loudly if it did not take.
    $waited = 0
    while ((Get-Process -Name EXCEL -ErrorAction SilentlyContinue) -and $waited -lt 20) {
        Start-Sleep -Milliseconds 500
        $waited++
    }
    if (Get-Process -Name EXCEL -ErrorAction SilentlyContinue) {
        Write-Output "NOTE: an Excel process is still running (probably one you opened yourself)."
    }

    for ($try = 1; $try -le 3; $try++) {
        if ($null -eq $had) {
            Remove-ItemProperty -Path $regPath -Name AccessVBOM -ErrorAction SilentlyContinue
        } else {
            Set-ItemProperty -Path $regPath -Name AccessVBOM -Value $had -Type DWord
        }
        Start-Sleep -Milliseconds 400
        $now = (Get-ItemProperty -Path $regPath -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM
        if ($now -eq $had) { break }
    }

    $now = (Get-ItemProperty -Path $regPath -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM
    $label = "ABSENT (off)"
    if ($null -ne $had) { $label = [string]$had }
    if ($now -eq $had) {
        Write-Output ("AccessVBOM restored to " + $label)
    } else {
        Write-Output ("AccessVBOM IS STILL " + [string]$now + " AND COULD NOT BE RESTORED. Turn it off by hand:")
        Write-Output "   Excel > File > Options > Trust Center > Trust Center Settings > Macro Settings"
        Write-Output "   and clear 'Trust access to the VBA project object model'."
        exit 3
    }
}
