Copy-Item "C:\Users\anderson.vieira\Desktop\TESTE RTS\Sucata_RETE_V6.xlsm" "C:\Users\anderson.vieira\Desktop\TESTE RTS\Sucata.xlsm" -Force
Write-Host "Restored from V6 backup"

$e = New-Object -ComObject Excel.Application
$e.Visible = $false
$e.DisplayAlerts = $false
$w = $e.Workbooks.Open("C:\Users\anderson.vieira\Desktop\TESTE RTS\Sucata.xlsm")
$v = $w.VBProject

$c = $v.VBComponents("mProcess")
$totalLines = $c.CodeModule.CountOfLines

for ($idx = 1; $idx -le $totalLines; $idx++) {
    $line = $c.CodeModule.Lines($idx, 1)
    $trimmed = $line.Trim()
    if ($trimmed -eq 'session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell -1, "GSTRP"') {
        Write-Host ("Found at line " + $idx)
        
        # Replace line with selectAll
        $newLine = '        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").SelectAll'
        $c.CodeModule.ReplaceLine($idx, $newLine)
        
        # Delete next 3 lines (SelectColumn, btn[40], sendVKey 42)
        $c.CodeModule.DeleteLines($idx + 1, 3)
        
        Write-Host ("Replace done at line " + $idx)
        break
    }
}

$w.Save()
$w.Close()
$e.Quit()
Write-Host "DONE"
