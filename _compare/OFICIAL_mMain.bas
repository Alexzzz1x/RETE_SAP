Attribute VB_Name = "mMain"
Option Explicit

Global colSer, colMat, colLot, colEqp, colDep, colDat, colCen, _
        colSts, colStu, colOSM, colDoc, colObs As String
        
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
    colDoc = "K"
    'colDatDc = "L"
    colObs = "M"

End Sub

'Faz a Conexão com o SAP
Function SapConnect() As Boolean

    Static SapGuiAuto, SAPApp, SAPCon As Object
    'Static sMSG As String

On Error GoTo NotConnected  'Direciona Para Erro Se SAP Não Acessível

    'Obtém o SAPGUI Scripting object
    Set SapGuiAuto = GetObject("SAPGUI")
    'Obtém o SAPGUI em Execução
    Set SAPApp = SapGuiAuto.GetScriptingEngine
    'Obtém o Primeiro Sistema Atualmente Conectado
    Set SAPCon = SAPApp.Children(0)
    'Obtém a Primeira Sessão (Janela) Nesta Conexão
    Set session = SAPCon.Children(0)
    
    SapConnect = True
    
Exit Function
NotConnected:
    SapConnect = False
    MsgBox "LOGAR NO SAP ANTES!", vbCritical

End Function



' Popula os Combos (sem duplicados)
Sub comboPop(cbo As ComboBox, ws As Worksheet, col As String)

    Dim iLastCel As Long
    Dim i As Long
    Dim val As String
    Dim itemExists As Boolean
    Dim j As Long

    cbo.clear ' Limpa o ComboBox antes de popular
    iLastCel = ws.Range(col & ws.Rows.Count).End(xlUp).Row ' Última linha preenchida

    For i = 2 To iLastCel ' Começa da linha 2 (pulando o cabeçalho)
        val = Trim(ws.Cells(i, col).Value)
        itemExists = False
        
        ' Verifica se o valor já existe no ComboBox
        For j = 0 To cbo.ListCount - 1
            If cbo.List(j) = val Then
                itemExists = True
                Exit For
            End If
        Next j
        
        ' Se não existe, adiciona ao ComboBox
        If Not itemExists And val <> "" Then
            cbo.AddItem val
        End If
    Next i

End Sub

' Retorna  o depósito da equipe parametro
Function GetDepEqp(sEqp As String) As String
    Dim rMatch As Range
    Dim sResult As String
    
    Set rMatch = wsConf.Range("A:A").Find(What:=sEqp, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not rMatch Is Nothing Then
        sResult = CStr(rMatch.Offset(0, 1).Value)
    Else
        sResult = "" ' Retorna string vazia se não encontrar
    End If
    
    GetDepEqp = sResult
End Function

'Salva os seriais da planilha em arquivo txt na pasta TEMP
Sub ExportSeriaisToTemp()

    Dim lastRow As Long
    Dim i As Long
    Dim filePath As String
    Dim fileNum As Integer
    Dim tempFolder As String
    
    ' Define a última linha com conteúdo na coluna A
    lastRow = wsRETE.Cells(wsRETE.Rows.Count, "A").End(xlUp).Row

    ' Caminho da pasta TEMP do usuário
    tempFolder = Environ("TEMP")
    
    ' Define o caminho completo do arquivo
    filePath = tempFolder & "\RETE_Export.txt"
    
    ' Cria o arquivo para escrita
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    
    ' Escreve os valores da coluna A no arquivo
    For i = 2 To lastRow
        If Trim(wsRETE.Cells(i, 1).Value) <> "" Then
            Print #fileNum, wsRETE.Cells(i, 1).Value
        End If
    Next i
    
    ' Fecha o arquivo
    Close #fileNum

End Sub

Function CheckSerials(serialList As String) As Boolean

    Dim colPrefixes As Collection
    Dim colFound As Collection
    Dim arrLines() As String
    Dim serial As String
    Dim i As Long
    Dim lastRow As Long
    Dim rngFound As Range

    Set colPrefixes = New Collection
    Set colFound = New Collection

    ' Carregar siglas da coluna E da wsCONF
    lastRow = wsConf.Cells(wsConf.Rows.Count, "E").End(xlUp).Row
    For i = 1 To lastRow
        If Trim(wsConf.Cells(i, "E").Value) <> "" Then
            On Error Resume Next
            colPrefixes.Add UCase(Trim(wsConf.Cells(i, "E").Value)), UCase(Trim(wsConf.Cells(i, "E").Value))
            On Error GoTo 0
        End If
    Next i

    ' Separar linhas do texto recebido
    arrLines = Split(serialList, vbCrLf)

    ' Verificar cada linha
    For i = LBound(arrLines) To UBound(arrLines)
        serial = Trim(arrLines(i))
        If serial <> "" Then
            ' Ignorar siglas conhecidas
            On Error Resume Next
            colPrefixes.Item (UCase(serial))
            If Err.Number = 5 Then
                Err.clear
                ' Buscar na coluna A da wsRETE
                With wsRETE.Columns("A")
                    Set rngFound = .Find(What:=serial, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
                    If Not rngFound Is Nothing Then
                        colFound.Add serial
                    End If
                End With
            Else
                Err.clear
            End If
            On Error GoTo 0
        End If
    Next i

    ' Se encontrar ao menos um, retornar True
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
    Dim serial As String
    Dim rng As Range
    Dim found As Range
    Dim notFound As String
    Dim i As Long
    Dim colLetter As String

    ' Usa a variável colSer para definir a coluna
    colLetter = colSer

    ' Divide os seriais pelo caractere de nova linha
    inputLines = Split(txtSeries.Text, vbCrLf)
    notFound = ""

    ' Define a área de busca dinamicamente com base na coluna colSer
    Set rng = wsRETE.Range(colLetter & "1", wsRETE.Cells(wsRETE.Rows.Count, colLetter).End(xlUp))

    For i = LBound(inputLines) To UBound(inputLines)
        serial = Trim(inputLines(i))
        If Len(serial) > 0 Then
            Set found = rng.Find(What:=serial, LookIn:=xlValues, LookAt:=xlWhole)
            If found Is Nothing Then
                notFound = notFound & serial & vbCrLf
            End If
        End If
    Next i

    If Len(notFound) > 0 Then
        MsgBox "Os seguintes seriais não estão cadastrados:" & vbCrLf & notFound, vbExclamation
        AreAllSerialsInSheet = False
    Else
        AreAllSerialsInSheet = True
    End If
End Function

