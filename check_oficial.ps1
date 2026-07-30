$e = New-Object -ComObject Excel.Application
$e.Visible = $false
$e.DisplayAlerts = $false
$w = $e.Workbooks.Open("\\terra\conecta\arquivos\obras\Elpa - Almox CR Mauá\CONTROLES\CONTROLE_RETE ( OFICIAL ).xlsm")
$v = $w.VBProject
foreach ($c in $v.VBComponents) {
    $lines = $c.CodeModule.CountOfLines
    Write-Host ("$($c.Name) - Type=$($c.Type) - $lines lines")
}
$w.Close($false)
$e.Quit()
Write-Host "DONE"
