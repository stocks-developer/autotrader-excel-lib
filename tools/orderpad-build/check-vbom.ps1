# READ ONLY. Reports whether Excel's "Trust access to the VBA project object
# model" is enabled. Changes nothing.
#
# AccessVBOM is normally ABSENT rather than 0, so the correct way to undo an
# enable later is Remove-ItemProperty, not setting it back to 0.
$any = $false
Get-ChildItem 'HKCU:\Software\Microsoft\Office' -ErrorAction SilentlyContinue |
    Where-Object { $_.PSChildName -match '^\d+\.\d+$' } |
    ForEach-Object {
        $ver = $_.PSChildName
        $path = "HKCU:\Software\Microsoft\Office\$ver\Excel\Security"
        if (Test-Path $path) {
            $any = $true
            $v = (Get-ItemProperty -Path $path -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM
            if ($null -eq $v) {
                Write-Output "  Office $ver : AccessVBOM ABSENT  (normal, VBA project access is OFF)"
            } else {
                Write-Output "  Office $ver : AccessVBOM = $v"
            }
        }
    }
if (-not $any) { Write-Output "  no Excel\Security key under HKCU for any Office version" }

$xl = Get-Process -Name EXCEL -ErrorAction SilentlyContinue
if ($xl) { Write-Output "  EXCEL is RUNNING (pid $($xl.Id -join ','))" }
else { Write-Output "  EXCEL is not running" }
