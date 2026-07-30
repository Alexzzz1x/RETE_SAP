Attribute VB_Name = "mProcess"
Option Explicit

Sub SapRETE(txtSeries As Object)
    Dim inputLines() As String
    Dim Serial As String
    Dim i As Long
    Dim found As Range
    Dim sOSME As String
    Dim sEquipe As String
    Dim sDestino As String

    ' Define a variavel colSer para a coluna de busca
    Dim colLetter As String
    colLetter = colSer

    ' Divide os seriais por quebra de linha
    inputLines = Split(txtSeries.Text, vbCrLf)

    For i = LBound(inputLines) To UBound(inputLines)
        Serial = Trim(inputLines(i))
        If Len(Serial) > 0 Then
            ' Procura o serial na coluna definida por colSer
            Set found = wsRETE.Range(colLetter & "1", wsRETE.Cells(wsRETE.Rows.Count, colLetter).End(xlUp)) _
                          .Find(What:=Serial, LookIn:=xlValues, LookAt:=xlWhole)
            
            If Not found Is Nothing Then
                ' Le as informacoes da mesma linha do serial encontrado
                sOSME = wsRETE.Cells(found.Row, colOSM).Value
                sEquipe = wsRETE.Cells(found.Row, colEqp).Value
                sDestino = wsRETE.Cells(found.Row, colDep).Value

                ' Chama a macro MIGO com os dados capturados
                Call ExecutarMigo(sOSME, sDestino, Serial)
            End If
        End If
    Next i

    MsgBox "Lote RETE enviado com sucesso!", vbInformation
End Sub

Sub ExecutarMigo(sOSME As String, sDestino As String, Serial As String)
    Dim vbsPath As String
    Dim wsh As Object
    Dim cmd As String
    Dim exec As Object
    Dim docRetornado As String

    ' Caminho do script VBS
    vbsPath = "C:\Users\anderson.vieira\Desktop\Nova pasta\SCRIPT DE RETE.vbs"

    ' Monta o comando cscript com os parametros
    cmd = "cscript //Nologo """ & vbsPath & """ """ & sOSME & """ """ & sDestino & """ """ & Serial & """"

    Set wsh = CreateObject("WScript.Shell")
    Set exec = wsh.exec(cmd)

    ' Captura a sada do VBScript (o WScript.Echo doc)
    docRetornado = Trim(exec.StdOut.ReadAll())

    ' Verifica se a sada possui algum erro retornado pelo VBS
    If InStr(docRetornado, "ERRO:") > 0 Then
        MsgBox "Falha no SAP ao processar serial " & Serial & vbCrLf & docRetornado, vbCritical
        Exit Sub
    End If

    ' Se retornou um documento valido, atualiza a planilha
    If Len(docRetornado) > 0 Then
        ' Procura a linha do serial na planilha e grava o documento retornado na coluna L (colDoc)
        Dim found As Range
        Set found = wsRETE.Range(colSer & "1", wsRETE.Cells(wsRETE.Rows.Count, colSer).End(xlUp)) _
                        .Find(What:=Serial, LookIn:=xlValues, LookAt:=xlWhole)
        If Not found Is Nothing Then
            wsRETE.Cells(found.Row, colDoc).Value = docRetornado
        End If
    End If

End Sub

Sub add2Plan(ws1 As Worksheet, txtBox As Object, _
            Optional sEqp As String = "", Optional sLot As String = "", Optional sDep As String = "", _
            Optional sOSM As String = "", Optional sDoc As String = "", Optional sObs As String = "")
             
    Dim sSERIE() As String
    Dim iLn As Long
    Dim iEnd As Long
    Dim sDat As Date
    Dim sSig As String
    Dim sCurrentLot As String
    Dim sVal As String
    Dim sSerial As String
    Dim c1 As String, c2 As String, c3 As String
    
    sDat = Format(Date, "dd/mm/yyyy")
    sSERIE = Split(txtBox.Text, vbCrLf)
    sCurrentLot = ""
       
    For iLn = LBound(sSERIE) To UBound(sSERIE)
        sVal = Trim(sSERIE(iLn))
        If Len(sVal) = 0 Then GoTo NextLoop
        
        If Len(sVal) >= 3 Then
            c1 = Mid(sVal, 1, 1)
            c2 = Mid(sVal, 2, 1)
            c3 = Mid(sVal, 3, 1)
            
            If (c1 >= "A" And c1 <= "Z" Or c1 >= "a" And c1 <= "z") And _
               (c2 >= "A" And c2 <= "Z" Or c2 >= "a" And c2 <= "z") And _
               (c3 >= "0" And c3 <= "9") Then
                
                sCurrentLot = UCase(Left(sVal, 3))
                sSerial = Trim(Mid(sVal, 4))
                
                If Len(sSerial) > 0 Then
                    iEnd = ws1.Cells(ws1.Rows.Count, colSer).End(xlUp).Row + 1
                    ws1.Cells(iEnd, colSer).Value = sSerial
                    ws1.Cells(iEnd, colLot).Value = sCurrentLot
                    ws1.Cells(iEnd, colDat).Value = sDat
                    If sEqp <> "" Then
                        ws1.Cells(iEnd, colEqp).Value = sEqp
                        ws1.Cells(iEnd, colDep).Value = GetDepEqp(sEqp)
                    End If
                    If sOSM <> "" Then ws1.Cells(iEnd, colOSM).Value = sOSM
                    If sDoc <> "" Then ws1.Cells(iEnd, colDoc).Value = sDoc
                    If sObs <> "" Then ws1.Cells(iEnd, colObs).Value = sObs
                End If
                GoTo NextLoop
            End If
        End If
        
        iEnd = ws1.Cells(ws1.Rows.Count, colSer).End(xlUp).Row + 1
        ws1.Cells(iEnd, colSer).Value = sVal
        ws1.Cells(iEnd, colDat).Value = sDat
        If sCurrentLot <> "" Then
            ws1.Cells(iEnd, colLot).Value = sCurrentLot
        End If
        If sEqp <> "" Then
            ws1.Cells(iEnd, colEqp).Value = sEqp
            ws1.Cells(iEnd, colDep).Value = GetDepEqp(sEqp)
        End If
        If sOSM <> "" Then ws1.Cells(iEnd, colOSM).Value = sOSM
        If sDoc <> "" Then ws1.Cells(iEnd, colDoc).Value = sDoc
        If sObs <> "" Then ws1.Cells(iEnd, colObs).Value = sObs
        
NextLoop:
    Next iLn
    
    MsgBox "Dados adicionados com sucesso!", vbInformation
End Sub

Sub SapIQ09()

Dim sSapSer, sSapMat, sSapStS, sSapStU, sSapCen, sSapRP As String
Dim x As Integer

    ' Faz a conexao com o SAP
    If SapConnect = False Then
        Exit Sub
    End If
    
    ' Exporta os seriais para arquivo TXT na pasta TEMP do usuario
    Call ExportSeriaisToTemp

    ' Abre a transacao IQ09
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/NIQ09"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/btn%_SERNR_%_APP_%-VALU_PUSH").press

    ' Indica o diretorio e o nome do arquivo temporario
    session.findById("wnd[1]/tbar[0]/btn[23]").press
    session.findById("wnd[2]/usr/ctxtDY_PATH").Text = Environ("TEMP")
    session.findById("wnd[2]/usr/ctxtDY_FILENAME").Text = "\RETE_Export.txt"
    session.findById("wnd[2]/tbar[0]/btn[0]").press

    ' Confirma a entrada e processa
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    
    ' Retorna o numero total de linhas no grid do SAP
    Dim iSapRow, iVisibleRows, iCurrentRow
    iSapRow = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").RowCount
    iVisibleRows = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").VisibleRowCount
    
    ' Inicia no topo
    iCurrentRow = 0
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").FirstVisibleRow = iCurrentRow
    
    ' Faz a rolagem por paginas
    Do While iCurrentRow + iVisibleRows < iSapRow
        iCurrentRow = iCurrentRow + iVisibleRows
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").FirstVisibleRow = iCurrentRow
    Loop

    ' Garante que a ultima pagina foi carregada
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").FirstVisibleRow = 0
    
    ' Configura a barra de progresso antes de iniciar o loop
    With frmMain.pgbMain
        .Min = 0
        .Max = iSapRow
        .Value = 0
    End With
    
    ' Percorre as linhas do grid SAP
    For x = 0 To iSapRow - 1
    
        ' Atualiza a barra de progresso
        frmMain.pgbMain.Value = x + 1
        DoEvents

        sSapSer = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(x, "SERNR")
        sSapMat = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(x, "MATNR")
        sSapStS = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(x, "STTXT")
        sSapStU = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(x, "USTXT")
        sSapCen = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(x, "WERK")
        
        Dim iRow As Long
        Dim iLastRow As Long

        iLastRow = wsRETE.Cells(wsRETE.Rows.Count, colSer).End(xlUp).Row
        
        For iRow = 1 To iLastRow
            If Trim(wsRETE.Cells(iRow, 1).Value) = Trim(sSapSer) Then
                wsRETE.Cells(iRow, colMat).Value = sSapMat
                wsRETE.Cells(iRow, colCen).Value = sSapCen
                wsRETE.Cells(iRow, colSts).Value = sSapStS
                wsRETE.Cells(iRow, colStu).Value = sSapStU
                Exit For
            End If
        Next iRow
    
    Next
    
    frmMain.pgbMain.Value = frmMain.pgbMain.Max
    MsgBox "Processo Concluido"

End Sub

Sub SapOSME()

    ' Faz a conexao com o SAP
    If SapConnect = False Then
        Exit Sub
    End If

Dim i, x, y As Integer
Dim sSERIE, sORDEM, sEquipe As String
Dim tree, coll As Object
Dim objClipBrd As MSForms.DataObject

    With frmMain.pgbMain
        .Min = 0
        .Max = wsRETE.UsedRange.Rows.Count
        .Value = 0
    End With
      
    For i = 2 To wsRETE.UsedRange.Rows.Count
    
        frmMain.pgbMain.Value = i + 1
        DoEvents
    
        If wsRETE.Cells(i, colStu).Value = "ELAB" Then
            If wsRETE.Cells(i, colOSM).Value = "" Then

                sSERIE = Trim(CStr(wsRETE.Cells(i, colSer).Value))
                If Len(sSERIE) < 1 Then
                    MsgBox "Final da Consulta ou Linha de Serial em Branco encontrada"
                    Exit For
                End If
            
                session.findById("wnd[0]/tbar[0]/okcd").Text = "/NIQ09"
                session.findById("wnd[0]").sendVKey 0
                session.findById("wnd[0]/usr/txtSERNR-LOW").Text = sSERIE
                session.findById("wnd[0]/tbar[1]/btn[8]").press
                session.findById("wnd[0]/usr/tabsTABSTRIP/tabpT").Select
                session.findById("wnd[0]/usr/tabsTABSTRIP/tabpT/ssubSUB_DATA:SAPLITO0:0122/subSUB_0122B:SAPLITO0:1221/btn%_AUTOTEXT002").press
        
                Set tree = session.findById("wnd[0]/usr/cntlTREE_CONTAINER/shellcont/shell")
                Set coll = tree.GetAllNodeKeys()
                
                For x = 1 To coll.Length - 1
                    sORDEM = coll.ElementAt(x)
                    sEquipe = coll.ElementAt(x)
                    sORDEM = tree.GetItemText(sORDEM, "2")
                    sEquipe = tree.GetItemText(sEquipe, "3")
                    If Left(sORDEM, 4) = "1000" Then
                        If Len(sORDEM) = 12 Then
                            wsRETE.Cells(i, colOSM).Value = sORDEM
                            Exit For
                        End If
                    End If
                Next x
                
            End If
        End If
    Next i
    
    frmMain.pgbMain.Value = frmMain.pgbMain.Max
    MsgBox "Processo concluido."
    
End Sub

Sub SetLot(txtSeries As Object, cboLote As Object)

    Dim inputLines() As String
    Dim Serial As String
    Dim i As Long
    Dim found As Range
    Dim colSerial As String, colLote As String
    Dim ws As Worksheet

    Set ws = wsRETE
    colSerial = colSer
    colLote = colLot

    inputLines = Split(txtSeries.Text, vbCrLf)

    For i = LBound(inputLines) To UBound(inputLines)
        Serial = Trim(inputLines(i))
        If Len(Serial) > 0 Then
            Set found = ws.Range(colSerial & "1", ws.Cells(ws.Rows.Count, colSerial).End(xlUp)) _
                          .Find(What:=Serial, LookIn:=xlValues, LookAt:=xlWhole)
            
            If Not found Is Nothing Then
                ws.Cells(found.Row, colLote).Value = cboLote.Text
            End If
        End If
    Next i

    MsgBox "Lote atribuido.", vbInformation
    
End Sub

Sub SendLotRETE(txtSeries As Object, txtDoc As Object)

    Dim inputLines() As String
    Dim Serial As String
    Dim i As Long
    Dim found As Range
    Dim destRow As Long

    inputLines = Split(txtSeries.Text, vbCrLf)

    For i = UBound(inputLines) To LBound(inputLines) Step -1
        Serial = Trim(inputLines(i))
        If Len(Serial) > 0 Then
            Set found = wsRETE.Range(colSer & "1", wsRETE.Cells(wsRETE.Rows.Count, colSer).End(xlUp)) _
                            .Find(What:=Serial, LookIn:=xlValues, LookAt:=xlWhole)
            If Not found Is Nothing Then
                wsRETE.Cells(found.Row, colDoc).Value = frmMain.txtDoc.Text
                
                If Len(frmMain.txtDat) > 0 Then
                    wsRETE.Cells(found.Row, colDatDc).Value = frmMain.txtDat.Text
                End If

                destRow = wsSEND.Cells(wsSEND.Rows.Count, 1).End(xlUp).Row + 1

                wsRETE.Rows(found.Row).Copy Destination:=wsSEND.Rows(destRow)
                wsRETE.Rows(found.Row).Delete
                
            End If
        End If
    Next i

    MsgBox "Envio de lote cadastrado!", vbInformation
    
End Sub



' =====================================================================
' RETIRADA RETE - Saida de mercadorias para centro RETE (Z77)
' =====================================================================



Sub SapPreencherOSME()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long, iCount As Long, iTotal As Long
    Dim sEquipe As String, sStatus As String, sSubStatus As String, sOSME As String
    Dim dictOSME As Object
    
    Call setInit
    Set ws = wsRETE
    Set dictOSME = CreateObject("Scripting.Dictionary")
    lastRow = ws.Cells(ws.Rows.Count, colSer).End(xlUp).Row
    
    If Not SapConnect() Then Exit Sub
    
    iTotal = 0
    For i = 2 To lastRow
        sStatus = Trim(ws.Cells(i, colSts).Value)
        sSubStatus = Trim(ws.Cells(i, colStu).Value)
        sOSME = Trim(ws.Cells(i, colOSM).Value)
        If sStatus = "LIDI" And sSubStatus = "ELAB" And sOSME = "" Then
            iTotal = iTotal + 1
        End If
    Next i
    
    If iTotal = 0 Then
        MsgBox "Nenhum serial LIDI/ELAB sem OSME encontrado.", vbInformation
        Exit Sub
    End If
    
    With frmMain.pgbMain
        .Min = 0
        .Max = iTotal
        .Value = 0
    End With
    
    iCount = 0
    For i = 2 To lastRow
        sStatus = Trim(ws.Cells(i, colSts).Value)
        sSubStatus = Trim(ws.Cells(i, colStu).Value)
        sEquipe = Trim(ws.Cells(i, colEqp).Value)
        sOSME = Trim(ws.Cells(i, colOSM).Value)
        
        If sStatus = "LIDI" And sSubStatus = "ELAB" And sEquipe <> "" And sOSME = "" Then
            iCount = iCount + 1
            frmMain.pgbMain.Value = iCount
            DoEvents
            
            If Not dictOSME.Exists(sEquipe) Then
                dictOSME(sEquipe) = GetOSMEValida(sEquipe)
            End If
            sOSME = dictOSME(sEquipe)
            If sOSME <> "" Then
                ws.Cells(i, colOSM).Value = sOSME
            End If
        End If
    Next i
    
    frmMain.pgbMain.Value = frmMain.pgbMain.Max
    MsgBox "OSMEs preenchidas! " & iCount & " seriais atualizados.", vbInformation
End Sub





Function GetOSMEValida(sEquipe As String) As String
    Dim grid As Object, x As Integer, iRow As Integer
    Dim sOSME As String, sDataFim As String
    Dim dDataFim As Date, dMelhorData As Date
    Dim sMelhorOSME As String
    
    dMelhorData = DateSerial(1900, 1, 1)
    sMelhorOSME = ""
    
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/NIW39"
    session.findById("wnd[0]").sendVKey 0
    
    On Error Resume Next
    session.findById("wnd[0]/usr/chkDY_MAB").Selected = vbTrue
    session.findById("wnd[0]/usr/chkDY_HIS").Selected = vbTrue
    session.findById("wnd[0]/usr/cmbDY_PARVW").Key = "Z2"
    On Error GoTo 0
    
    session.findById("wnd[0]/usr/ctxtDY_PARNR").Text = sEquipe
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    
    Application.Wait (Now + TimeValue("0:00:03"))
    
    On Error Resume Next
    Set grid = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
    If grid Is Nothing Then
        GetOSMEValida = ""
        Exit Function
    End If
    On Error GoTo 0
    
    iRow = grid.RowCount
    If iRow = 0 Then
        GetOSMEValida = ""
        Exit Function
    End If
    
    grid.SelectColumn "GLTRP"
    session.findById("wnd[0]/tbar[1]/btn[40]").press
    
    For x = 0 To iRow - 1
        sOSME = grid.GetCellValue(x, "AUFNR")
        sDataFim = grid.GetCellValue(x, "GLTRP")
        If sDataFim <> "" Then
            dDataFim = DateSerial(CInt(Mid(sDataFim, 7, 4)), CInt(Mid(sDataFim, 4, 2)), CInt(Left(sDataFim, 2)))
            If dDataFim > Date And dDataFim > dMelhorData Then
                dMelhorData = dDataFim
                sMelhorOSME = sOSME
            End If
        End If
    Next x
    
    GetOSMEValida = sMelhorOSME
End Function

Sub SapRetiradaReteZ77()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long, iCount As Long, iTotal As Long
    Dim sSerial As String, sMaterial As String, sDep As String, sEquipe As String
    Dim sOSME As String, sStatus As String, sSubStatus As String
    Dim sDoc As String, sCen As String
    
    Call setInit
    Set ws = wsRETE
    lastRow = ws.Cells(ws.Rows.Count, colSer).End(xlUp).Row
    
    If Not SapConnect() Then Exit Sub
    
    iTotal = 0
    For i = 2 To lastRow
        sSerial = Trim(ws.Cells(i, colSer).Value)
        sStatus = Trim(ws.Cells(i, colSts).Value)
        sSubStatus = Trim(ws.Cells(i, colStu).Value)
        sOSME = Trim(ws.Cells(i, colOSM).Value)
        sDoc = Trim(ws.Cells(i, colDoc).Value)
        sCen = Trim(ws.Cells(i, colCen).Value)
        If sSerial <> "" And sStatus = "LIDI" And sSubStatus = "ELAB" And sOSME <> "" And sDoc = "" And sCen = "" Then
            iTotal = iTotal + 1
        End If
    Next i
    
    If iTotal = 0 Then
        MsgBox "Nenhum serial elegivel para Retirada Z77.", vbInformation
        Exit Sub
    End If
    
    With frmMain.pgbMain
        .Min = 0
        .Max = iTotal
        .Value = 0
    End With
    
    iCount = 0
    For i = 2 To lastRow
        sSerial = Trim(ws.Cells(i, colSer).Value)
        sMaterial = Trim(ws.Cells(i, colMat).Value)
        sDep = Trim(ws.Cells(i, colDep).Value)
        sStatus = Trim(ws.Cells(i, colSts).Value)
        sSubStatus = Trim(ws.Cells(i, colStu).Value)
        sOSME = Trim(ws.Cells(i, colOSM).Value)
        sEquipe = Trim(ws.Cells(i, colEqp).Value)
        If UCase(Left(sEquipe, 3)) = "RCL" Then
            sDep = "CL04"
        ElseIf UCase(Left(sEquipe, 3)) = "CSC" Or UCase(Left(sEquipe, 3)) = "CDC" Then
            sDep = "CL03"
        End If
        sDoc = Trim(ws.Cells(i, colDoc).Value)
        sCen = Trim(ws.Cells(i, colCen).Value)
        
        If sSerial <> "" And sStatus = "LIDI" And sSubStatus = "ELAB" And sOSME <> "" And sDoc = "" And sCen = "" Then
            iCount = iCount + 1
            frmMain.pgbMain.Value = iCount
            DoEvents
            
            On Error Resume Next
            sDoc = ExecutarMIGO_Z77(sOSME, sMaterial, sSerial, sDep)
            If Err.Number <> 0 Then
                Err.Clear
                sDoc = ""
            End If
            On Error GoTo 0
            
            If sDoc <> "" Then
                ws.Cells(i, colDoc).Value = sDoc
                ws.Cells(i, colDatDc).Value = Format(Date, "dd/mm/yyyy")
                ws.Cells(i, colObs).Value = "RETE OK"
            Else
                ws.Cells(i, colObs).Value = "Erro, Verificar"
            End If
        End If
    Next i
    
    frmMain.pgbMain.Value = frmMain.pgbMain.Max
    MsgBox "Retirada Z77 concluida! " & iCount & " seriais processados.", vbInformation
End Sub

Function ExecutarMIGO_Z77(sOSME As String, sMaterial As String, sSerial As String, sDep As String) As String
    Dim x As Integer, sSapMat As String, sNDoc() As String
    Dim sRecebedor As String, sPontoDescarga As String
    Dim iTotalRows As Integer
    
    If sDep = "CL03" Then
        sRecebedor = "CL03"
        sPontoDescarga = "RETE 0CL03"
    ElseIf sDep = "CL04" Then
        sRecebedor = "CL04"
        sPontoDescarga = "RETE 0CL04"
    Else
        ExecutarMIGO_Z77 = ""
        Exit Function
    End If
    
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/NMIGO"
    session.findById("wnd[0]").sendVKey 0
    
    On Error Resume Next
    If Not session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_ITEMDETAIL:SAPLMIGO:0302/btnBUTTON_DETAIL") Is Nothing Then
        session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_ITEMDETAIL:SAPLMIGO:0302/btnBUTTON_DETAIL").press
    End If
    On Error GoTo 0
    
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-ACTION").Key = "A07"
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-REFDOC").Key = "R08"
    
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_FIRSTLINE:SAPLMIGO:0010/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2070/ctxtGODYNPRO-ORDER_NUMBER").Text = sOSME
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_FIRSTLINE:SAPLMIGO:0010/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2070/btnMIGO_OK_GO").press
    
    Application.Wait (Now + TimeValue("0:00:03"))
    
    Dim iFoundRow As Integer
    iFoundRow = -1
    
    For x = 0 To 50
        On Error Resume Next
        sSapMat = Trim(session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM/ctxtGOITEM-MATNR[1," & x & "]").Text)
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            Exit For
        End If
        On Error GoTo 0
        
        If sSapMat = Trim(sMaterial) Then
            iFoundRow = x
            Exit For
        End If
    Next x
    
    If iFoundRow = -1 Then
        ExecutarMIGO_Z77 = ""
        Exit Function
    End If
    
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM/btnGOITEM-ZEILE[0," & iFoundRow & "]").SetFocus
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM/btnGOITEM-ZEILE[0," & iFoundRow & "]").press
    
    Application.Wait (Now + TimeValue("0:00:01"))
    
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/subSUB_DETAIL_TAKE:SAPLMIGO:0304/chkGODYNPRO-DETAIL_TAKE").Selected = vbTrue
    
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_QUANTITIES").Select
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_QUANTITIES/ssubSUB_TS_GOITEM_QUANTITIES:SAPLMIGO:0315/txtGOITEM-ERFMG").Text = "1"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_DESTINAT.").Select
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_DESTINAT./ssubSUB_TS_GOITEM_DESTINATION:SAPLMIGO:0325/txtGOITEM-WEMPF").Text = sRecebedor
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_DESTINAT./ssubSUB_TS_GOITEM_DESTINATION:SAPLMIGO:0325/txtGOITEM-ABLAD").Text = sPontoDescarga
    
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_SERIAL").Select
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_SERIAL/ssubSUB_TS_GOITEM_SERIAL:SAPLMIGO:0360/tblSAPLMIGOTV_GOSERIAL/txtGOSERIAL-SERIALNO[0,0]").Text = sSerial
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_SERIAL/ssubSUB_TS_GOITEM_SERIAL:SAPLMIGO:0360/tblSAPLMIGOTV_GOSERIAL/txtGOSERIAL-SERIALNO[0,0]").SetFocus
    session.findById("wnd[0]").sendVKey 0
    
    session.findById("wnd[0]/tbar[1]/btn[7]").press
    
    If Not session.findById("wnd[1]", False) Is Nothing Then
        ExecutarMIGO_Z77 = ""
        Exit Function
    End If
    
    session.findById("wnd[0]/tbar[1]/btn[23]").press
    
    On Error Resume Next
    If Not session.findById("wnd[1]", False) Is Nothing Then
        session.findById("wnd[1]/usr/btnBUTTON_2").press
    End If
    On Error GoTo 0
    
    sNDoc = Split(session.findById("wnd[0]/sbar").Text, " ")
    If UBound(sNDoc) >= 1 Then
        ExecutarMIGO_Z77 = sNDoc(1)
    Else
        ExecutarMIGO_Z77 = ""
    End If
End Function





Sub SapRetiradaReteZ77_Alt()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long, iCount As Long, iTotal As Long
    Dim sSerial As String, sMaterial As String, sDep As String, sEquipe As String
    Dim sOSME As String, sStatus As String, sSubStatus As String
    Dim sDoc As String, sCen As String, sOper As String
    
    Call setInit
    Set ws = wsRETE
    lastRow = ws.Cells(ws.Rows.Count, colSer).End(xlUp).Row
    
    If Not SapConnect() Then Exit Sub
    
    iTotal = 0
    For i = 2 To lastRow
        sSerial = Trim(ws.Cells(i, colSer).Value)
        sStatus = Trim(ws.Cells(i, colSts).Value)
        sSubStatus = Trim(ws.Cells(i, colStu).Value)
        sOSME = Trim(ws.Cells(i, colOSM).Value)
        sEquipe = Trim(ws.Cells(i, colEqp).Value)
        sDoc = Trim(ws.Cells(i, colDoc).Value)
        sCen = Trim(ws.Cells(i, colCen).Value)
        If sSerial <> "" And sStatus = "LIDI" And sSubStatus = "ELAB" And sOSME <> "" And sDoc = "" And sCen = "" Then
            iTotal = iTotal + 1
        End If
    Next i
    
    If iTotal = 0 Then
        MsgBox "Nenhum serial elegivel para Retirada Z77 Alternativa.", vbInformation
        Exit Sub
    End If
    
    With frmMain.pgbMain
        .Min = 0
        .Max = iTotal
        .Value = 0
    End With
    
    iCount = 0
    For i = 2 To lastRow
        sSerial = Trim(ws.Cells(i, colSer).Value)
        sEquipe = Trim(ws.Cells(i, colEqp).Value)
        sMaterial = Trim(ws.Cells(i, colMat).Value)
        sDep = Trim(ws.Cells(i, colDep).Value)
        sStatus = Trim(ws.Cells(i, colSts).Value)
        If UCase(Left(sEquipe, 3)) = "RCL" Then
            sDep = "CL04"
        ElseIf UCase(Left(sEquipe, 3)) = "CSC" Or UCase(Left(sEquipe, 3)) = "CDC" Then
            sDep = "CL03"
        End If
        sSubStatus = Trim(ws.Cells(i, colStu).Value)
        sOSME = Trim(ws.Cells(i, colOSM).Value)
        sOper = Trim(ws.Cells(i, colOper).Value)
        If sOper = "" Then sOper = "0100"
        sDoc = Trim(ws.Cells(i, colDoc).Value)
        sCen = Trim(ws.Cells(i, colCen).Value)
        
        If sSerial <> "" And sStatus = "LIDI" And sSubStatus = "ELAB" And sOSME <> "" And sDoc = "" And sCen = "" Then
            iCount = iCount + 1
            frmMain.pgbMain.Value = iCount
            DoEvents
            
            On Error Resume Next
            sDoc = ExecutarMIGO_Z77_Alt(sOSME, sMaterial, sSerial, sDep, sOper)
            If Err.Number <> 0 Then
                Err.Clear
                sDoc = ""
            End If
            On Error GoTo 0
            
            If sDoc <> "" Then
                ws.Cells(i, colDoc).Value = sDoc
                ws.Cells(i, colDatDc).Value = Format(Date, "dd/mm/yyyy")
                ws.Cells(i, colObs).Value = "RETE OK"
            Else
                ws.Cells(i, colObs).Value = "Erro, Verificar"
            End If
        End If
    Next i
    
    frmMain.pgbMain.Value = frmMain.pgbMain.Max
    MsgBox "Retirada Z77 Alternativa concluida! " & iCount & " seriais processados.", vbInformation
End Sub

Function ExecutarMIGO_Z77_Alt(sOSME As String, sMaterial As String, sSerial As String, sDep As String, Optional sOper As String = "0100") As String
    Dim sRecebedor As String, sPontoDescarga As String, sNDoc() As String
    
    If sDep = "CL03" Then
        sRecebedor = "CL03"
        sPontoDescarga = "RETE 0CL03"
    ElseIf sDep = "CL04" Then
        sRecebedor = "CL04"
        sPontoDescarga = "RETE 0CL04"
    Else
        ExecutarMIGO_Z77_Alt = ""
        Exit Function
    End If
    
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/NMIGO"
    session.findById("wnd[0]").sendVKey 0
    
    On Error Resume Next
    If Not session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_ITEMDETAIL:SAPLMIGO:0302/btnBUTTON_DETAIL") Is Nothing Then
        session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_ITEMDETAIL:SAPLMIGO:0302/btnBUTTON_DETAIL").press
    End If
    On Error GoTo 0
    
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-ACTION").Key = "A07"
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-REFDOC").Key = "R10"
    
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_MATERIAL/ssubSUB_TS_GOITEM_MATERIAL:SAPLMIGO:0310/ctxtGOITEM-MAKTX").Text = sMaterial
    session.findById("wnd[0]").sendVKey 0
    
    Application.Wait (Now + TimeValue("0:00:02"))
    
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_QUANTITIES").Select
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_QUANTITIES/ssubSUB_TS_GOITEM_QUANTITIES:SAPLMIGO:0315/txtGOITEM-ERFMG").Text = "1"
    session.findById("wnd[0]").sendVKey 0
    
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_DESTINAT.").Select
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_DESTINAT./ssubSUB_TS_GOITEM_DESTINATION:SAPLMIGO:0325/ctxtGOITEM-NAME1").Text = "RETE"
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_DESTINAT./ssubSUB_TS_GOITEM_DESTINATION:SAPLMIGO:0325/ctxtGOITEM-LGOBE").Text = sDep
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_DESTINAT./ssubSUB_TS_GOITEM_DESTINATION:SAPLMIGO:0325/txtGOITEM-WEMPF").Text = sRecebedor
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_DESTINAT./ssubSUB_TS_GOITEM_DESTINATION:SAPLMIGO:0325/txtGOITEM-ABLAD").Text = sPontoDescarga
    session.findById("wnd[0]").sendVKey 0
    
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_ACCOUNT").Select
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_ACCOUNT/ssubSUB_TS_GOITEM_ACCOUNT:SAPLMIGO:0345/ssubSUB_ACCOUNTINGBLOCK:SAPLKACB:1003/ctxtCOBL-AUFNR").Text = sOSME
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_ACCOUNT/ssubSUB_TS_GOITEM_ACCOUNT:SAPLMIGO:0345/ssubSUB_ACCOUNTINGBLOCK:SAPLKACB:1003/ctxtCOBL-VORNR_AUF").Text = sOper
    session.findById("wnd[0]").sendVKey 0
    
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_SERIAL").Select
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_SERIAL/ssubSUB_TS_GOITEM_SERIAL:SAPLMIGO:0360/tblSAPLMIGOTV_GOSERIAL/txtGOSERIAL-SERIALNO[0,0]").Text = sSerial
    session.findById("wnd[0]").sendVKey 0
    
    session.findById("wnd[0]/tbar[1]/btn[7]").press
    
    If Not session.findById("wnd[1]", False) Is Nothing Then
        ExecutarMIGO_Z77_Alt = ""
        Exit Function
    End If
    
    session.findById("wnd[0]/tbar[1]/btn[23]").press
    
    On Error Resume Next
    If Not session.findById("wnd[1]", False) Is Nothing Then
        session.findById("wnd[1]/usr/btnBUTTON_2").press
    End If
    On Error GoTo 0
    
    sNDoc = Split(session.findById("wnd[0]/sbar").Text, " ")
    If UBound(sNDoc) >= 1 Then
        ExecutarMIGO_Z77_Alt = sNDoc(1)
    Else
        ExecutarMIGO_Z77_Alt = ""
    End If
End Function

Sub PreencherOperacoesIW39()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long, j As Long, x As Long
    Dim sEquipe As String, sMaterial As String, sOper As String
    Dim sCen As String, sOSME As String, sStatus As String, sSubStatus As String
    Dim dictEquipes As Object, dictOper As Object
    Dim sKey As Variant
    Dim iTotal As Long, iCount As Long
    Dim sDebug As String
    Dim tbl As Object, ScrollBar As Object
    Dim sFound As Boolean, sEqpPlan As String, sMatPlan As String, sOperPlan As String
    Dim iMatch As Long
    Dim iRowCount As Long
    
    Call setInit
    Set ws = wsRETE
    lastRow = ws.Cells(ws.Rows.Count, colSer).End(xlUp).Row
    
    If Not SapConnect() Then Exit Sub
    
    Set dictEquipes = CreateObject("Scripting.Dictionary")
    Set dictOper = CreateObject("Scripting.Dictionary")
    
    ' Collect unique equipes that need operations
    For i = 2 To lastRow
        sStatus = Trim(ws.Cells(i, colSts).Value)
        sSubStatus = Trim(ws.Cells(i, colStu).Value)
        sEquipe = Trim(ws.Cells(i, colEqp).Value)
        sOper = Trim(ws.Cells(i, colOper).Value)
        
        If sStatus = "LIDI" And sSubStatus = "ELAB" And sEquipe <> "" And sOper = "" Then
            If Not dictEquipes.Exists(sEquipe) Then
                dictEquipes(sEquipe) = True
            End If
        End If
    Next i
    
    If dictEquipes.Count = 0 Then
        MsgBox "Nenhuma equipe pendente para preencher operacoes.", vbInformation
        Exit Sub
    End If
    
    With frmMain.pgbMain
        .Min = 0
        .Max = dictEquipes.Count
        .Value = 0
    End With
    
    iCount = 0
    For Each sKey In dictEquipes.Keys
        iCount = iCount + 1
        frmMain.pgbMain.Value = iCount
        DoEvents
        
        sEquipe = sKey
        
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/NIW39"
        session.findById("wnd[0]").sendVKey 0
        
        Application.Wait (Now + TimeValue("0:00:02"))
        
        session.findById("wnd[0]/usr/cmbDY_PARVW").Key = "Z2"
        session.findById("wnd[0]/usr/ctxtDY_PARNR").Text = sEquipe
        session.findById("wnd[0]").sendVKey 0
        
        Application.Wait (Now + TimeValue("0:00:02"))
        
        session.findById("wnd[0]/tbar[1]/btn[8]").press
        
        Application.Wait (Now + TimeValue("0:00:02"))
        
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectAll
        
        Application.Wait (Now + TimeValue("0:00:02"))
        
        session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpMUEB").Select
        
        Application.Wait (Now + TimeValue("0:00:01"))
        
        sDebug = ""
        iMatch = 0
        ' Read all rows from componentes grid - scroll in batches
        On Error Resume Next
        Set tbl = session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020")
        Set ScrollBar = tbl.VerticalScrollBar
        ScrollBar.Position = ScrollBar.Max
        Application.Wait (Now + TimeValue("0:00:01"))
        ScrollBar.Position = 0
        Application.Wait (Now + TimeValue("0:00:01"))
        On Error GoTo 0
        
        For x = 0 To 99
            On Error Resume Next
            sMaterial = Trim(session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/ctxtRESBD-MATNR[1," & x & "]").Text)
            If Err.Number <> 0 Then
                Err.Clear
                On Error GoTo 0
                Exit For
            End If
            On Error GoTo 0
            
            If sMaterial = "" Then Exit For
            
            sCen = Trim(session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/ctxtRESBD-WERKS[9," & x & "]").Text)
            sOper = Trim(session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1107/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020/txtRESBD-VORNR[10," & x & "]").Text)
            
            If sCen = "RETE" Then
                sDebug = sDebug & "Row " & x & ": Mat=" & sMaterial & " Cen=" & sCen & " Oper=" & sOper & vbCrLf
            End If
            If sCen = "RETE" And sOper <> "" And sOper <> "0010" Then
                ' Fill ALL rows with this material (any equipe)
                For j = 2 To lastRow
                    If Trim(ws.Cells(j, colMat).Value) = sMaterial Then
                        sStatus = Trim(ws.Cells(j, colSts).Value)
                        sSubStatus = Trim(ws.Cells(j, colStu).Value)
                        sOperPlan = Trim(ws.Cells(j, colOper).Value)
                        If sOperPlan = "" And sStatus = "LIDI" And sSubStatus = "ELAB" Then
                            ws.Cells(j, colOper).Value = sOper
                            iMatch = iMatch + 1
                        End If
                    End If
                Next j
            End If
        Next x
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/NIW39"
        If sDebug <> "" Then MsgBox sDebug, vbInformation, "RETE components for " & sEquipe
        session.findById("wnd[0]").sendVKey 0
    Next sKey
    
    ' Mark rows that still have no OPER as Erro, Verificar
    For i = 2 To lastRow
        sStatus = Trim(ws.Cells(i, colSts).Value)
        sSubStatus = Trim(ws.Cells(i, colStu).Value)
        If ws.Cells(i, colOper).Value = "" And _
           sStatus = "LIDI" And sSubStatus = "ELAB" And _
           Trim(ws.Cells(i, colOSM).Value) <> "" Then
            ws.Cells(i, colObs).Value = "Erro, Verificar"
        End If
    Next i
    
    frmMain.pgbMain.Value = frmMain.pgbMain.Max
    MsgBox "Operacoes preenchidas! " & dictEquipes.Count & " equipes processadas. Total matches: " & iMatch, vbInformation
End Sub


