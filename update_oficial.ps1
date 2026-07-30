$srcDir = "C:\Users\anderson.vieira\Desktop\TESTE RTS\RETE_SAP"
$oficialBkp = "$srcDir\BACKUP_CONTROLE_RETE_OFICIAL.xlsm"
$oficialDest = "\\terra\conecta\arquivos\obras\Elpa - Almox CR Mauá\CONTROLES\CONTROLE_RETE ( OFICIAL ).xlsm"

$e = New-Object -ComObject Excel.Application
$e.Visible = $false
$e.DisplayAlerts = $false

# Open backup first (network file might be locked)
$w = $e.Workbooks.Open($oficialBkp)
$v = $w.VBProject

# Remove old modules
$removeNames = @("mProcess", "mMain", "frmMain")
foreach ($name in $removeNames) {
    $comp = $v.VBComponents($name)
    if ($comp) {
        $v.VBComponents.Remove($comp)
        Write-Host ("Removed: " + $name)
    }
}

# Import new modules
# Need to temporarily copy .frm and .frx to same directory for import
$v.VBComponents.Import("$srcDir\mProcess.bas")
Write-Host "Imported mProcess.bas"

$v.VBComponents.Import("$srcDir\mMain.bas")
Write-Host "Imported mMain.bas"

$v.VBComponents.Import("$srcDir\frmMain.frm")
Write-Host "Imported frmMain.frm"

# Delete old wsRETE lines and add new
$comp = $v.VBComponents("wsRETE")
$lines = $comp.CodeModule.CountOfLines
if ($lines -gt 0) { $comp.CodeModule.DeleteLines(1, $lines) }
$content = Get-Content "$srcDir\_compare\DEV_wsRETE.bas" -Raw
$comp.CodeModule.AddFromString($content)
Write-Host "Updated wsRETE"

$w.Save()
$w.Close()
$e.Quit()
Write-Host "DONE - Backup updated"
