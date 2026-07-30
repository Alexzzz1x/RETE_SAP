$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open("C:\Users\anderson.vieira\Desktop\TESTE RTS\Sucata.xlsm")
$vbProj = $wb.VBProject
$c = $vbProj.VBComponents("frmMain")
Write-Host "frmMain: Type=$($c.Type) Lines=$($c.CodeModule.CountOfLines)"
$first = $c.CodeModule.Lines(1, 5)
Write-Host "First 5 lines:"
Write-Host $first
$wb.Close($false)
$excel.Quit()
