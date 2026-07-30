$e = New-Object -ComObject Excel.Application
$e.Visible = $false
$e.DisplayAlerts = $false
$w = $e.Workbooks.Open("C:\Users\anderson.vieira\Desktop\TESTE RTS\RETE_SAP\BACKUP_CONTROLE_RETE_OFICIAL.xlsm")
$v = $w.VBProject
foreach ($c in $v.VBComponents) {
    $lines = $c.CodeModule.CountOfLines
    Write-Host ("$($c.Name) - Type=$($c.Type) - $lines lines")
}
$w.Close($false)
$e.Quit()
Write-Host "DONE"
