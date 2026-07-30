$e = New-Object -ComObject Excel.Application
$e.Visible = $false
$e.DisplayAlerts = $false
$w = $e.Workbooks.Open("C:\Users\anderson.vieira\Desktop\TESTE RTS\RETE_SAP\BACKUP_CONTROLE_RETE_OFICIAL.xlsm")
$v = $w.VBProject
$needAdd = $true
foreach ($c in $v.VBComponents) {
    if ($c.Name -eq "Módulo3") { $needAdd = $false }
}
if ($needAdd) {
    $v.VBComponents.Import("C:\Users\anderson.vieira\Desktop\TESTE RTS\RETE_SAP\_compare\DEV_Módulo3.bas")
    Write-Host "Modulo3 ADDED"
} else {
    Write-Host "Modulo3 already present"
}
$w.Save()
$w.Close()
$e.Quit()
Write-Host "DONE"
