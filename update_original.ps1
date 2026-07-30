$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open("C:\Users\anderson.vieira\Desktop\TESTE RTS\Sucata.xlsm")
$vbProj = $wb.VBProject
$basDir = "C:\Users\anderson.vieira\Desktop\TESTE RTS\RETE_SAP"

$basFiles = @{}
$basFiles["EstaPastaDeTrabalho"] = $true
$basFiles["wsRETE"] = $true
$basFiles["frmMain"] = $true
$basFiles["mMain"] = $true
$basFiles["mProcess"] = $true

foreach ($c in $vbProj.VBComponents) {
    $name = $c.Name
    if ($basFiles.ContainsKey($name)) {
        $basPath = "$basDir\$name.bas"
        if (Test-Path $basPath) {
            $lines = $c.CodeModule.CountOfLines
            if ($lines -gt 0) { $c.CodeModule.DeleteLines(1, $lines) }
            $content = Get-Content $basPath -Raw
            $c.CodeModule.AddFromString($content)
            Write-Host "OK: $name"
        }
    }
}

# Handle accented names separately
foreach ($c in $vbProj.VBComponents) {
    $name = $c.Name
    $stripped = $name -replace '[^a-zA-Z0-9_]', ''
    if ($stripped -ne $name) {
        $basPath = "$basDir\$stripped.bas"
        if (Test-Path $basPath) {
            $lines = $c.CodeModule.CountOfLines
            if ($lines -gt 0) { $c.CodeModule.DeleteLines(1, $lines) }
            $content = Get-Content $basPath -Raw
            $c.CodeModule.AddFromString($content)
            Write-Host "OK: $name (from $stripped.bas)"
        }
    }
}

$wb.Save()
$wb.Close()
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
Write-Host "DONE!"
