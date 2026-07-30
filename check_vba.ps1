$e = New-Object -ComObject Excel.Application
$e.Visible = $false
$e.DisplayAlerts = $false
$w = $e.Workbooks.Open("C:\Users\anderson.vieira\Desktop\TESTE RTS\Sucata.xlsm")
$v = $w.VBProject

$c = $v.VBComponents("frmMain")
Write-Host "frmMain: Type=$($c.Type) Lines=$($c.CodeModule.CountOfLines)"
Write-Host "  Line1: $($c.CodeModule.Lines(1,1))"

$c2 = $v.VBComponents("mProcess")
Write-Host "mProcess: Lines=$($c2.CodeModule.CountOfLines)"
$first3 = $c2.CodeModule.Lines(1,3)
Write-Host "  Lines 1-3: $first3"

$c3 = $v.VBComponents("mMain")
Write-Host "mMain: Lines=$($c3.CodeModule.CountOfLines)"
$first3m = $c3.CodeModule.Lines(1,3)
Write-Host "  Lines 1-3: $first3m"

$w.Close($false)
$e.Quit()
[Runtime.InteropServices.Marshal]::ReleaseComObject($e) | Out-Null
Write-Host "COMPLETE"
