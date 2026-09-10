# Capture the real Excel window, because the PDF path does not render Form
# controls -- the seven buttons are provably in the file (seven ctrlProps) and
# still came out of ExportAsFixedFormat invisible. A screenshot of the actual
# window is the only honest check that they look right.
#
# 🔴 PrintWindow, never CopyFromScreen. CopyFromScreen reads the SCREEN at the
# window's coordinates, so anything sitting on top is captured instead --
# that is how a WhatsApp conversation once ended up in a screenshot on this
# machine. PrintWindow asks the window to draw ITSELF into our bitmap, so
# whatever is in front of it is irrelevant.
#
# Not part of the build -- this is how you LOOK at the result. See README.md.
#
# 🔴 `param` must be the FIRST statement in a PowerShell script; paths are
# derived below it, from $PSScriptRoot.
param(
    [string]$Path  = "",
    [string]$Sheet = "Orders",
    [string]$Out   = "",
    [int]$Zoom = 70,
    # Make one cell's comment visible before capturing, so the note can be
    # looked at. A comment only shows on hover otherwise, and hover cannot be
    # staged in a screenshot.
    [string]$ShowComment = ""
)

# tools/orderpad-build/ -> the repository root is two levels up.
$Repo = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent

if ($Path -eq "") { $Path = Join-Path $Repo "clients\excel\current\samples\AutoTraderWeb-OrderPad.xlsm" }
if ($Out  -eq "") { $Out  = Join-Path $PSScriptRoot ("build\shot-" + $Sheet.ToLower() + ".png") }

if (-not (Test-Path $Path)) {
    Write-Output "no workbook at $Path"
    exit 2
}

Add-Type -AssemblyName System.Drawing
Add-Type @'
using System;
using System.Runtime.InteropServices;
public class Win {
  [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr hdc, uint flags);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
  [DllImport("user32.dll")] public static extern bool MoveWindow(IntPtr h, int x, int y, int w, int ht, bool repaint);
  [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }
}
'@

$dir = Split-Path $Out -Parent
if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }

$excel = New-Object -ComObject Excel.Application
$excel.DisplayAlerts = $false

try {
    $wb = $excel.Workbooks.Open($Path, 0, $true)
    $wb.Worksheets.Item($Sheet).Activate()
    $excel.Visible = $true
    $excel.WindowState = -4143            # xlNormal, so the size is ours to set
    $excel.ActiveWindow.Zoom = $Zoom

    if ($ShowComment -ne "") {
        $cell = $wb.Worksheets.Item($Sheet).Range($ShowComment)
        if ($null -ne $cell.Comment) {
            $cell.Comment.Visible = $true
            Write-Output ("showing comment on {0}!{1}" -f $Sheet, $ShowComment)
        } else {
            Write-Output ("NO comment on {0}!{1}" -f $Sheet, $ShowComment)
        }
    }

    $h = [IntPtr]$excel.Hwnd
    [void][Win]::MoveWindow($h, 60, 60, 1500, 620, $true)
    Start-Sleep -Milliseconds 1200        # let it lay out and repaint

    $r = New-Object Win+RECT
    [void][Win]::GetWindowRect($h, [ref]$r)
    $w = $r.Right - $r.Left
    $ht = $r.Bottom - $r.Top

    $bmp = New-Object System.Drawing.Bitmap($w, $ht)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $hdc = $g.GetHdc()
    $ok = [Win]::PrintWindow($h, $hdc, 2)   # 2 = PW_RENDERFULLCONTENT
    $g.ReleaseHdc($hdc)
    $g.Dispose()

    if ($ok) {
        $bmp.Save($Out, [System.Drawing.Imaging.ImageFormat]::Png)
        Write-Output "captured $Sheet -> $Out"
    } else {
        Write-Output "PrintWindow FAILED for $Sheet"
    }
    $bmp.Dispose()
    $wb.Close($false)
}
finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
}
