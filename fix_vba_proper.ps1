param()

# Restore from V6 first
Copy-Item "C:\Users\anderson.vieira\Desktop\TESTE RTS\Sucata_RETE_V6.xlsm" "C:\Users\anderson.vieira\Desktop\TESTE RTS\Sucata.xlsm" -Force
Write-Host "Restored from V6 backup"

$e = New-Object -ComObject Excel.Application
$e.Visible = $false
$e.DisplayAlerts = $false
$w = $e.Workbooks.Open("C:\Users\anderson.vieira\Desktop\TESTE RTS\Sucata.xlsm")
$v = $w.VBProject

# Remove old mProcess and re-import properly
$v.VBComponents.Remove($v.VBComponents("mProcess"))
$v.VBComponents.Import("C:\Users\anderson.vieira\Desktop\TESTE RTS\RETE_SAP\mProcess.bas")
Write-Host "mProcess imported OK"

# Remove old mMain and re-import properly
$v.VBComponents.Remove($v.VBComponents("mMain"))
$v.VBComponents.Import("C:\Users\anderson.vieira\Desktop\TESTE RTS\RETE_SAP\mMain.bas")
Write-Host "mMain imported OK"

$w.Save()
$w.Close()
$e.Quit()
Write-Host "DONE"
