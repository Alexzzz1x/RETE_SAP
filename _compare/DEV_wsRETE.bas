VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wsRETE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
    frmMain.Show
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim cell As Range
    Dim sCurrentLot As String
    Dim lines() As String
    Dim j As Long
    Dim iRow As Long
    Dim sTxt As String
    Dim sVal As String
    Dim bWasLot As Boolean

    If Intersect(Target, Me.Columns("A")) Is Nothing Then Exit Sub
    If Target.Count > 256 Then Exit Sub
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    On Error GoTo CleanUp
    
    For Each cell In Target
        sTxt = Trim(cell.Value)
        If Len(sTxt) = 0 Then GoTo NextCell
        
        If InStr(sTxt, vbLf) > 0 Or InStr(sTxt, vbCr) > 0 Then
            sTxt = Replace(sTxt, vbCrLf, vbLf)
            sTxt = Replace(sTxt, vbCr, vbLf)
            lines = Split(sTxt, vbLf)
            
            iRow = cell.Row
            For j = LBound(lines) To UBound(lines)
                sVal = Trim(lines(j))
                If Len(sVal) > 0 Then
                    If IsLotPrefix(sVal) Then
                        sCurrentLot = UCase(Left(sVal, 3))
                    Else
                        WriteSerial wsRETE, iRow, sVal, sCurrentLot
                        iRow = iRow + 1
                    End If
                End If
            Next j
            cell.ClearContents
            GoTo NextCell
        End If
        
        sVal = sTxt
        iRow = cell.Row
        bWasLot = False
        
        If IsLotPrefix(sVal) Then
            sCurrentLot = UCase(Left(sVal, 3))
            cell.ClearContents
            bWasLot = True
        ElseIf Len(sVal) > 0 Then
            WriteSerial wsRETE, iRow, sVal, sCurrentLot
        End If
NextCell:
    Next cell

CleanUp:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Private Function IsLotPrefix(ByVal sVal As String) As Boolean
    Dim c1 As String, c2 As String, c3 As String
    If Len(sVal) < 3 Then
        IsLotPrefix = False
        Exit Function
    End If
    c1 = Mid(sVal, 1, 1)
    c2 = Mid(sVal, 2, 1)
    c3 = Mid(sVal, 3, 1)
    IsLotPrefix = (c1 >= "A" And c1 <= "Z" Or c1 >= "a" And c1 <= "z") And _
                  (c2 >= "A" And c2 <= "Z" Or c2 >= "a" And c2 <= "z") And _
                  (c3 >= "0" And c3 <= "9")
End Function

Private Sub WriteSerial(ws As Worksheet, ByVal iRow As Long, ByVal sVal As String, ByVal sCurrentLot As String)
    ws.Cells(iRow, "A").Value = sVal
    If sCurrentLot <> "" Then
        ws.Cells(iRow, "C").Value = sCurrentLot
    End If
    If ws.Cells(iRow, "F").Value = "" Then
        ws.Cells(iRow, "F").Value = Format(Date, "dd/mm/yyyy")
    End If
End Sub

