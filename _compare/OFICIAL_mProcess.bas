Attribute VB_Name = "mProcess"
Option Explicit

Sub add2Plan(ws1 As Worksheet, txtBox As Object, _
            Optional sEqp As String = "", Optional sLot As String = "", Optional sDep As String = "", _
            Optional sOSM As String = "", Optional sDoc As String = "", Optional sObs As String = "")
             
    Dim sSERIE() As String
    Dim iLn As Long
    Dim iEnd As Long
    Dim sDat As Date
    Dim rngFind As Range
    Dim sSig As String
    Dim bLnTp As Boolean
    
    ' Range dos tipos de equipamentos
    Set rngFind = wsConf.Range("E:E")
        
    ' Obter a data atual no formato dd/mm/aaaa
    sDat = Format(Date, "dd/mm/yyyy")
    
    ' Separar o texto em linhas com base no caractere de quebra de linha
    sSERIE = Split(txtBox.Text, vbCrLf)
       
    ' Encontrar a última linha na coluna de seriais (Rotina inserida no loop para evitar linha em branco na identificação de TIPO/SIGLA de material)
    'iEnd = ws1.Cells(ws1.Rows.Count, colSer).End(xlUp).Row + 1
    
    ' Adicionar os dados à planilha
    For iLn = LBound(sSERIE) To UBound(sSERIE)
        If sSERIE(iLn) <> "" Then ' Verifica se a linha não está vazia
            
            ' Verifica se a linha é sigla
            sSig = UCase(Left(sSERIE(iLn), 2))
            If Not rngFind.Find(What:=UCase(sSig), LookIn:=xlValues, LookAt:=xlPart) Is Nothing Then
                sLot = UCase(Left(sSERIE(iLn), 3))
                GoTo NextLoop
            End If
            
    ' Encontrar a última linha na coluna de seriais
    iEnd = ws1.Cells(ws1.Rows.Count, colSer).End(xlUp).Row + 1

            ' Adicionar serial e data nas colunas especificadas
            ws1.Cells(iEnd + 0, colSer).Value = sSERIE(iLn) 'Serial
            ws1.Cells(iEnd + 0, colDat).Value = sDat 'Data
            
            ' Adicionar equipe e localiza o depósito
            If sEqp <> "" Then
                ws1.Cells(iEnd + 0, colEqp).Value = sEqp
                sDep = GetDepEqp(sEqp)
            End If

            ' Adicionar o restante das informações nas colunas especificadas
            If sLot <> "" Then
                ws1.Cells(iEnd + 0, colLot).Value = sLot
            End If
            
            If sDep <> "" Then
                ws1.Cells(iEnd + 0, colDep).Value = sDep
            End If
            
            If sOSM <> "" Then
                ws1.Cells(iEnd + 0, colOSM).Value = sOSM
            End If
            
            If sDoc <> "" Then
                ws1.Cells(iEnd + 0, colDoc).Value = sDoc
            End If
            
            If sObs <> "" Then
                ws1.Cells(iEnd + 0, colObs).Value = sObs
            End If
            
        End If
        
NextLoop:
    
    Next iLn
    
    MsgBox "Dados adicionados com sucesso!", vbInformation

End Sub

Sub SapIQ09()

Dim sSapSer, sSapMat, sSapStS, sSapStU, sSapCen, sSapRP As String
Dim x As Integer

    ' Faz a conexão com o SAP
    If SapConnect = False Then
        Exit Sub
    End If
    
    ' Exporta os seriais para arquivo TXT na pasta TEMP do usuário
    Call ExportSeriaisToTemp

    ' Abre a transação IQ09
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/NIQ09"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/btn%_SERNR_%_APP_%-VALU_PUSH").press

    ' Indica o diretório e o nome do arquivo temporario
    session.findById("wnd[1]/tbar[0]/btn[23]").press
    session.findById("wnd[2]/usr/ctxtDY_PATH").Text = Environ("TEMP") ' Caminho da pasta TEMP do usuário
    session.findById("wnd[2]/usr/ctxtDY_FILENAME").Text = "\RETE_Export.txt" ' Nome do arquivo TXT
    session.findById("wnd[2]/tbar[0]/btn[0]").press

    ' Confirma a entrada e processa
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    
    ' Retorna o número total de linhas no grid do SAP
    Dim iSapRow, iVisibleRows, iCurrentRow
    iSapRow = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").RowCount
    iVisibleRows = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").VisibleRowCount
    
    ' Inicia no topo
    iCurrentRow = 0
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").FirstVisibleRow = iCurrentRow
    
    ' Faz a rolagem por páginas (senão o loop abaixo não coleta todos os valores)
    Do While iCurrentRow + iVisibleRows < iSapRow
        iCurrentRow = iCurrentRow + iVisibleRows
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").FirstVisibleRow = iCurrentRow
    Loop

    ' Garante que a última página foi carregada
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
        DoEvents ' Permite que a interface seja atualizada

        sSapSer = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(x, "SERNR") 'Coluna Série
        sSapMat = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(x, "MATNR") 'Coluna Material
        sSapStS = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(x, "STTXT") 'Coluna StatSis
        sSapStU = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(x, "USTXT") 'Coluna StatUsr
        sSapCen = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(x, "WERK")  'Coluna Centro
        
        ' Loop na planilha wsRETE para encontrar a linha com o mesmo número de série
        Dim iRow As Long
        Dim iLastRow As Long

        iLastRow = wsRETE.Cells(wsRETE.Rows.Count, colSer).End(xlUp).Row
        
        For iRow = 1 To iLastRow
            If Trim(wsRETE.Cells(iRow, 1).Value) = Trim(sSapSer) Then
                ' Preenche as colunas...
                wsRETE.Cells(iRow, colMat).Value = sSapMat
                wsRETE.Cells(iRow, colCen).Value = sSapCen
                wsRETE.Cells(iRow, colSts).Value = sSapStS
                wsRETE.Cells(iRow, colStu).Value = sSapStU
                Exit For ' Encontrado, sai do loop interno
            End If
        Next iRow
    
    Next
    
    frmMain.pgbMain.Value = frmMain.pgbMain.Max ' Garante barra cheia no fim
    MsgBox "Processo Concluído"

End Sub

Sub SapOSME()

    ' Faz a conexão com o SAP
    If SapConnect = False Then
        Exit Sub
    End If

Dim i, x, y As Integer
Dim sSERIE, sORDEM, sEQUIPE As String
Dim tree, coll As Object
Dim objClipBrd As MSForms.DataObject


    ' Configura a barra de progresso antes de iniciar o loop
    With frmMain.pgbMain
        .Min = 0
        .Max = wsRETE.UsedRange.Rows.Count
        .Value = 0
    End With
      
    For i = 2 To wsRETE.UsedRange.Rows.Count        'LER PLANILHA A PARTIR DA LINHA 2
    
        ' Atualiza a barra de progresso
        frmMain.pgbMain.Value = i + 1
        DoEvents ' Permite que a interface seja atualizada
    
        If wsRETE.Cells(i, colStu).Value = "ELAB" Then ' Atualiza apenas se for ELAB
            If wsRETE.Cells(i, colOSM).Value = "" Then  'Atualiza apenas se não houver OSME preenchida

                sSERIE = Trim(CStr(wsRETE.Cells(i, colSer).Value))        'A VARIÁVEL SERIE
                If Len(sSERIE) < 1 Then
                    MsgBox "Final da Consulta ou Linha de Serial em Branco encontrada"
                    Exit For
                End If
            
                session.findById("wnd[0]/tbar[0]/okcd").Text = "/NIQ09"
                session.findById("wnd[0]").sendVKey 0
                session.findById("wnd[0]/usr/txtSERNR-LOW").Text = sSERIE
                session.findById("wnd[0]/tbar[1]/btn[8]").press
                session.findById("wnd[0]/usr/tabsTABSTRIP/tabpT\06").Select
                session.findById("wnd[0]/usr/tabsTABSTRIP/tabpT\06/ssubSUB_DATA:SAPLITO0:0122/subSUB_0122B:SAPLITO0:1221/btn%_AUTOTEXT002").press
        
                Set tree = session.findById("wnd[0]/usr/cntlTREE_CONTAINER/shellcont/shell")
                Set coll = tree.GetAllNodeKeys()
                
                For x = 1 To coll.Length - 1
                    sORDEM = coll.ElementAt(x)
                    sEQUIPE = coll.ElementAt(x)
                    sORDEM = tree.GetItemText(sORDEM, "2")
                    sEQUIPE = tree.GetItemText(sEQUIPE, "3")
                    If Left(sORDEM, 4) = "1000" Then
                        If Len(sORDEM) = 12 Then
                            wsRETE.Cells(i, colOSM).Value = sORDEM   'PREENCHE NA PLAN ORDEM ENCONTRADA NO SAP
                            'wsRete.Cells(i, colEqp).Value = sEQUIPE  'PREENCHE NA PLAN EQUIPE ENCONTRADA NO SAP
                            Exit For
                        End If
                    End If
                Next x
                
            End If
        End If
    Next i
    
    frmMain.pgbMain.Value = frmMain.pgbMain.Max ' Garante barra cheia no fim
    MsgBox "Processo concluído."
    
End Sub

Sub SetLot(txtSeries As Object, cboLote As Object)

    Dim inputLines() As String
    Dim serial As String
    Dim i As Long
    Dim found As Range
    Dim colSerial As String, colLote As String
    Dim ws As Worksheet

    Set ws = wsRETE
    colSerial = colSer
    colLote = colLot

    ' Divide o conteúdo do TextBox por quebras de linha
    inputLines = Split(txtSeries.Text, vbCrLf)

    For i = LBound(inputLines) To UBound(inputLines)
        serial = Trim(inputLines(i))
        If Len(serial) > 0 Then
            ' Procura o serial na coluna definida por colSer
            Set found = ws.Range(colSerial & "1", ws.Cells(ws.Rows.Count, colSerial).End(xlUp)) _
                          .Find(What:=serial, LookIn:=xlValues, LookAt:=xlWhole)
            
            If Not found Is Nothing Then
                ' Escreve o valor do comboBox na coluna colLot, mesma linha
                ws.Cells(found.Row, colLote).Value = cboLote.Text
            End If
        End If
    Next i

    MsgBox "Lote atribuído.", vbInformation
    
End Sub

Sub SendLotRETE(txtSeries As Object, txtDoc As Object)

    Dim inputLines() As String
    Dim serial As String
    Dim i As Long
    Dim found As Range
    Dim destRow As Long

    ' Divide os seriais por quebra de linha
    inputLines = Split(txtSeries.Text, vbCrLf)

    ' Percorre de trás pra frente para deletar com segurança
    For i = UBound(inputLines) To LBound(inputLines) Step -1
        serial = Trim(inputLines(i))
        If Len(serial) > 0 Then
            Set found = wsRETE.Range(colSer & "1", wsRETE.Cells(wsRETE.Rows.Count, colSer).End(xlUp)) _
                            .Find(What:=serial, LookIn:=xlValues, LookAt:=xlWhole)
            If Not found Is Nothing Then
                ' Atualiza a coluna de documento
                wsRETE.Cells(found.Row, colDoc).Value = frmMain.txtDoc.Text
                
                ' Atualiza a coluna de data
                If Len(frmMain.txtDat) > 0 Then
                    wsRETE.Cells(found.Row, "L").Value = frmMain.txtDat.Text
                End If

                ' Localiza a próxima linha vazia em wsSEND
                destRow = wsSEND.Cells(wsSEND.Rows.Count, 1).End(xlUp).Row + 1

                ' Copia a linha da wsRETE para wsSEND
                wsRETE.Rows(found.Row).Copy Destination:=wsSEND.Rows(destRow)

                ' Remove a linha original da wsRETE
                wsRETE.Rows(found.Row).Delete
                
            End If
        End If
    Next i

    MsgBox "Envio de lote cadastrado!", vbInformation
    
End Sub

