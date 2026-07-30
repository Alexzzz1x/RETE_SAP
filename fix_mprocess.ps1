$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open("C:\Users\anderson.vieira\Desktop\TESTE RTS\Sucata.xlsm")
$vbProj = $wb.VBProject

# Update only mProcess (standard module - AddFromString works fine)
$comp = $vbProj.VBComponents("mProcess")
$lines = $comp.CodeModule.CountOfLines
if ($lines -gt 0) { $comp.CodeModule.DeleteLines(1, $lines) }
$content = Get-Content "C:\Users\anderson.vieira\Desktop\TESTE RTS\RETE_SAP\mProcess.bas" -Raw
$comp.CodeModule.AddFromString($content)
Write-Host "mProcess updated OK"

# Also update mMain and wsRETE to ensure column consts match
$comp = $vbProj.VBComponents("mMain")
$lines = $comp.CodeModule.CountOfLines
if ($lines -gt 0) { $comp.CodeModule.DeleteLines(1, $lines) }
$content = Get-Content "C:\Users\anderson.vieira\Desktop\TESTE RTS\RETE_SAP\mMain.bas" -Raw
$comp.CodeModule.AddFromString($content)
Write-Host "mMain updated OK"

$wb.Save()
$wb.Close()
$excel.Quit()
Write-Host "DONE"
