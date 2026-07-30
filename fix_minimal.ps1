Copy-Item "C:\Users\anderson.vieira\Desktop\TESTE RTS\Sucata_RETE_V6.xlsm" "C:\Users\anderson.vieira\Desktop\TESTE RTS\Sucata.xlsm" -Force
Write-Host "Restored from V6 backup"

$e = New-Object -ComObject Excel.Application
$e.Visible = $false
$e.DisplayAlerts = $false
$w = $e.Workbooks.Open("C:\Users\anderson.vieira\Desktop\TESTE RTS\Sucata.xlsm")
$v = $w.VBProject

$c = $v.VBComponents("mProcess")
$lines = $c.CodeModule.CountOfLines

# Find the lines with the old pattern and replace them
$old1 = "setCurrentCell -1, ""GSTRP"""
$old2 = "SelectColumn ""GSTRP"""
$old3 = "btn[40].press"
$old4 = "sendVKey 42"

for ($i = 1; $i -le $lines; $i++) {
    $line = $c.CodeModule.Lines($i, 1)
    if ($line.Trim() -eq $old1) {
        Write-Host "Found line $i: $($line.Trim())"
        $c.CodeModule.ReplaceLine($i, "session.findById(""wnd[0]/usr/cntlGRID1/shellcont/shell"").SelectAll")
        # Delete the next 3 lines
        $c.CodeModule.DeleteLines($i+1, 3)
        Write-Host "Replaced at line $i"
        break
    }
}

$w.Save()
$w.Close()
$e.Quit()
Write-Host "DONE"
