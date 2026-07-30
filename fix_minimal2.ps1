Copy-Item "C:\Users\anderson.vieira\Desktop\TESTE RTS\Sucata_RETE_V6.xlsm" "C:\Users\anderson.vieira\Desktop\TESTE RTS\Sucata.xlsm" -Force
Write-Host "Restored from V6 backup"

$e = New-Object -ComObject Excel.Application
$e.Visible = $false
$e.DisplayAlerts = $false
$w = $e.Workbooks.Open("C:\Users\anderson.vieira\Desktop\TESTE RTS\Sucata.xlsm")
$v = $w.VBProject

$c = $v.VBComponents("mProcess")
$totalLines = $c.CodeModule.CountOfLines

$oldPattern = "setCurrentCell -1, ""GSTRP"""

for ($idx = 1; $idx -le $totalLines; $idx++) {
    $line = $c.CodeModule.Lines($idx, 1)
    if ($line.Trim() -eq $oldPattern) {
        $foundLine = $idx
        Write-Host ("Found at line " + $foundLine)
        
        # Replace line with selectAll
        $newLine = 'session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").SelectAll'
        $c.CodeModule.ReplaceLine($foundLine, $newLine)
        
        # Delete next 3 lines (SelectColumn, btn[40].press, sendVKey 42)
        $c.CodeModule.DeleteLines($foundLine + 1, 3)
        
        Write-Host ("Replace done at line " + $foundLine)
        break
    }
}

$w.Save()
$w.Close()
$e.Quit()
Write-Host "DONE"
