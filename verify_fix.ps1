$e = New-Object -ComObject Excel.Application
$e.Visible = $false
$e.DisplayAlerts = $false
$w = $e.Workbooks.Open("C:\Users\anderson.vieira\Desktop\TESTE RTS\Sucata.xlsm")
$v = $w.VBProject
$c = $v.VBComponents("mProcess")
$total = $c.CodeModule.CountOfLines
Write-Host ("Total lines: " + $total)
for ($idx = 908; $idx -le 918; $idx++) {
    $line = $c.CodeModule.Lines($idx, 1)
    Write-Host ("Line " + $idx + ": " + $line.TrimEnd())
}
# Also check PreencherOperacoesIW39 exists
$full = $c.CodeModule.Lines(1, $total)
if ($full.Contains("PreencherOperacoesIW39")) {
    Write-Host "PreencherOperacoesIW39 function found!"
} else {
    Write-Host "ERROR: PreencherOperacoesIW39 NOT found!"
}
if ($full.Contains("SelectAll")) {
    Write-Host "SelectAll found!"
} else {
    Write-Host "ERROR: SelectAll NOT found!"
}
$w.Close($false)
$e.Quit()
Write-Host "VERIFIED"
