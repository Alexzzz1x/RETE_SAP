Attribute VB_Name = "mMain"
Option Explicit

Global colSer, colMat, colLot, colEqp, colDep, colDat, colCen, _
        colSts, colStu, colOSM, colOper, colDoc, colDatDc, colObs As String
        
Global session As Object

Sub setInit()

    colSer = "A"
    colMat = "B"
    colLot = "C"
    colEqp = "D"
    colDep = "E"
    colDat = "F"
    colCen = "G"
    colSts = "H"
    colStu = "I"
    colOSM = "J"
    colOper = "K"
    colDoc = "L"
    colDatDc = "M"
    colObs = "N"

End Sub

Function SapConnect() As Boolean

    Static SapGuiAuto, SAPApp, SAPCon As Object

On Error GoTo NotConnected

    Set SapGuiAuto = GetObject("SAPGUI")
    Set SAPApp = SapGuiAuto.GetScriptingEngine
    Set SAPCon = SAPApp.Children(0)
    Set session = SAPCon.Children(0)
    
    SapConnect = True
    
Exit Function
NotConnected:
    SapConnect = False
    MsgBox "LOGAR NO SAP ANTES!", vbCritical

End Function

Sub comboPop(cbo As ComboBox, ws As Worksheet, col As String)

    Dim iLastCel As Long
    Dim i As Long
    Dim val As String
    Dim itemExists As Boolean
    Dim j As Long

    cbo.Clear
    iLastCel = ws.Range(col & ws.Rows.Count).End(xlUp).Row

    For i = 2 To iLastCel
        val = Trim(ws.Cells(i, col).Value)
        itemExists = False
        
        For j = 0 To cbo.ListCount - 1
            If cbo.List(j) = val Then
                itemExists = True
                Exit For
            End If
        Next j
        
        If Not itemExists And val <> "" Then
            cbo.AddItem val
        End If
    Next i

End Sub

Function GetDepEqp(sEqp As String) As String
    Dim rMatch As Range
    Dim sResult As String
    
    Set rMatch = wsConf.Range("A:A").Find(What:=sEqp, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not rMatch Is Nothing Then
        sResult = CStr(rMatch.Offset(0, 1).Value)
    Else
        sResult = ""
    End If
    
    GetDepEqp = sResult
End Function

Sub ExportSeriaisToTemp()

    Dim lastRow As Long
    Dim i As Long
    Dim filePath As String
    Dim fileNum As Integer
    Dim tempFolder As String
    
    lastRow = wsRETE.Cells(wsRETE.Rows.Count, "A").End(xlUp).Row

    tempFolder = Environ("TEMP")
    
    filePath = tempFolder & "\RETE_Export.txt"
    
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    
    For i = 2 To lastRow
        If Trim(wsRETE.Cells(i, 1).Value) <> "" Then
            Print #fileNum, wsRETE.Cells(i, 1).Value
        End If
    Next i
    
    Close #fileNum

End Sub

Function CheckSerials(serialList As String) As Boolean

    Dim colPrefixes As Collection
    Dim colFound As Collection
    Dim arrLines() As String
    Dim Serial As String
    Dim i As Long
    Dim lastRow As Long
    Dim rngFound As Range

    Set colPrefixes = New Collection
    Set colFound = New Collection

    lastRow = wsConf.Cells(wsConf.Rows.Count, "E").End(xlUp).Row
    For i = 1 To lastRow
        If Trim(wsConf.Cells(i, "E").Value) <> "" Then
            On Error Resume Next
            colPrefixes.Add UCase(Trim(wsConf.Cells(i, "E").Value)), UCase(Trim(wsConf.Cells(i, "E").Value))
            On Error GoTo 0
        End If
    Next i

    arrLines = Split(serialList, vbCrLf)

    For i = LBound(arrLines) To UBound(arrLines)
        Serial = Trim(arrLines(i))
        If Serial <> "" Then
            On Error Resume Next
            colPrefixes.Item (UCase(Serial))
            If Err.Number = 5 Then
                Err.Clear
                With wsRETE.Columns("A")
                    Set rngFound = .Find(What:=Serial, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
                    If Not rngFound Is Nothing Then
                        colFound.Add Serial
                    End If
                End With
            Else
                Err.Clear
            End If
            On Error GoTo 0
        End If
    Next i

    If colFound.Count > 0 Then
        Dim msg As String
        msg = "Serials ja cadastrados na planilha:" & vbCrLf & vbCrLf
        For i = 1 To colFound.Count
            msg = msg & colFound(i) & vbCrLf
        Next i
        MsgBox msg, vbInformation, "Matching Serials"
        CheckSerials = True
    Else
        CheckSerials = False
    End If

End Function

Function AreAllSerialsInSheet(txtSeries As Object) As Boolean
    Dim inputLines() As String
    Dim Serial As String
    Dim rng As Range
    Dim found As Range
    Dim notFound As String
    Dim i As Long
    Dim colLetter As String

    colLetter = colSer

    inputLines = Split(txtSeries.Text, vbCrLf)
    notFound = ""

    Set rng = wsRETE.Range(colLetter & "1", wsRETE.Cells(wsRETE.Rows.Count, colLetter).End(xlUp))

    For i = LBound(inputLines) To UBound(inputLines)
        Serial = Trim(inputLines(i))
        If Len(Serial) > 0 Then
            Set found = rng.Find(What:=Serial, LookIn:=xlValues, LookAt:=xlWhole)
            If found Is Nothing Then
                notFound = notFound & Serial & vbCrLf
            End If
        End If
    Next i

    If Len(notFound) > 0 Then
        MsgBox "Os seguintes seriais nao estao cadastrados:" & vbCrLf & notFound, vbExclamation
        AreAllSerialsInSheet = False
    Else
        AreAllSerialsInSheet = True
    End If
End Function

