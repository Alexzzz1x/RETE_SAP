$e = New-Object -ComObject Excel.Application
$e.Visible = $false
$e.DisplayAlerts = $false
$w = $e.Workbooks.Open("C:\Users\anderson.vieira\Desktop\TESTE RTS\Sucata_RETE_V6.xlsm")
$v = $w.VBProject
$c = $v.VBComponents("mProcess")
$totalLines = $c.CodeModule.CountOfLines
Write-Host ("Total lines: " + $totalLines)
for ($idx = 905; $idx -le 920; $idx++) {
    $line = $c.CodeModule.Lines($idx, 1)
    Write-Host ("Line " + $idx + ": [" + $line.TrimEnd() + "]")
}
$w.Close($false)
$e.Quit()
Write-Host "DONE"
