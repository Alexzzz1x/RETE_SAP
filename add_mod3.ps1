$srcDir = "C:\Users\anderson.vieira\Desktop\TESTE RTS\RETE_SAP"
$e = New-Object -ComObject Excel.Application
$e.Visible = $false
$e.DisplayAlerts = $false
$w = $e.Workbooks.Open("$srcDir\BACKUP_CONTROLE_RETE_OFICIAL.xlsm")
$v = $w.VBProject

# Check if Modulo3 already exists
$exists = $false
foreach ($c in $v.VBComponents) {
    if ($c.Name -eq "Módulo3") { $exists = $true; break }
}

if (-not $exists) {
    # Import from DEV version
    $v.VBComponents.Import("$srcDir\_compare\DEV_Módulo1.bas")
    Write-Host "Added Modulo1"
    
    # Import Modulo3 from the Mdulo3.bas (accent-free filename)
    # First check if there's a proper .bas file
    if (Test-Path "$srcDir\Mdulo3.bas") {
        $v.VBComponents.Import("$srcDir\Mdulo3.bas")
        Write-Host "Added Modulo3 from Mdulo3.bas"
    }
}

$w.Save()
$w.Close()
$e.Quit()
Write-Host "DONE"
