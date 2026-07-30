$e = New-Object -ComObject Excel.Application
$e.Visible = $false
$e.DisplayAlerts = $false

# Load both workbooks
$w1 = $e.Workbooks.Open("C:\Users\anderson.vieira\Desktop\TESTE RTS\Sucata.xlsm")
$w2 = $e.Workbooks.Open("C:\Users\anderson.vieira\Desktop\TESTE RTS\RETE_SAP\BACKUP_CONTROLE_RETE_OFICIAL.xlsm")

$v1 = $w1.VBProject
$v2 = $w2.VBProject

# Export all modules from both
$outDir = "C:\Users\anderson.vieira\Desktop\TESTE RTS\RETE_SAP\_compare"
New-Item -ItemType Directory -Path $outDir -Force | Out-Null

# Export from Sucata (dev)
foreach ($c in $v1.VBComponents) {
    $name = $c.Name
    $lines = $c.CodeModule.CountOfLines
    if ($lines -gt 0) {
        $path = "$outDir\DEV_$name.bas"
        $c.Export($path)
        Write-Host ("DEV $name - $lines lines -> $path")
    }
}

# Export from Oficial
foreach ($c in $v2.VBComponents) {
    $name = $c.Name
    $lines = $c.CodeModule.CountOfLines
    if ($lines -gt 0) {
        $path = "$outDir\OFICIAL_$name.bas"
        $c.Export($path)
        Write-Host ("OFICIAL $name - $lines lines -> $path")
    }
}

$w1.Close($false)
$w2.Close($false)
$e.Quit()
Write-Host "DONE - compare in $outDir"
