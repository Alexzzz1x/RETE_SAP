Attribute VB_Name = "Módulo3"
Sub ProcessarEmailSucata(Item As Object)
    ' =====================================================================
    ' 1. DECLARAÇÃO DE VARIÁVEIS
    ' =====================================================================
    Dim TextoCorpo As String
    Dim Palavras() As String
    Dim Token As String
    Dim i As Integer
    Dim LoteAtual As String
    Dim Equipe As String
    Dim PrefixoEquipe As String
    Dim Deposito As String
    Dim LotesValidos As String
    Dim DataEntrada As Date
    
    ' Variáveis do Excel
    Dim xlApp As Object
    Dim MeuArquivo As Object
    Dim ws As Object
    Dim UltimaLinha As Long
    Dim ExcelJaEstavaAberto As Boolean
    Dim ArquivoJaEstavaAberto As Boolean
    
    On Error GoTo TrataErro

    ' =====================================================================
    ' 2. EXTRAIR EQUIPE DO ASSUNTO E DEFINIR O DEPÓSITO
    ' =====================================================================
    Dim AssuntoLimpo As String
    Dim PalavrasAssunto() As String
    AssuntoLimpo = Trim(UCase(Item.Subject))
    PalavrasAssunto = Split(AssuntoLimpo, " ")
    
    ' Pega a última palavra do assunto (ex: CSC300)
    If UBound(PalavrasAssunto) >= 0 Then
        Equipe = PalavrasAssunto(UBound(PalavrasAssunto))
    End If
    
    ' REGRA DO DEPÓSITO: Pega as 3 primeiras letras da Equipe
    PrefixoEquipe = Left(Equipe, 3)
    
    If PrefixoEquipe = "CCL" Or PrefixoEquipe = "CCS" Or PrefixoEquipe = "CSC" Or PrefixoEquipe = "CDC" Then
        Deposito = "CL03"
    ElseIf PrefixoEquipe = "RCL" Then
        Deposito = "CL04"
    Else
        Deposito = "" ' Deixa em branco se for alguma sigla diferente
    End If

    ' =====================================================================
    ' 3. LIMPEZA PROFUNDA DO CORPO DO E-MAIL (Fim do erro de lotes)
    ' =====================================================================
    TextoCorpo = UCase(Item.Body)
    
    ' Destrói qualquer tipo de quebra de linha ou espaço invisível do Outlook
    TextoCorpo = Replace(TextoCorpo, vbCrLf, " ")
    TextoCorpo = Replace(TextoCorpo, vbCr, " ")
    TextoCorpo = Replace(TextoCorpo, vbLf, " ")
    TextoCorpo = Replace(TextoCorpo, vbTab, " ")
    TextoCorpo = Replace(TextoCorpo, Chr(160), " ")
    
    While InStr(TextoCorpo, "  ") > 0
        TextoCorpo = Replace(TextoCorpo, "  ", " ")
    Wend
    
    ' Corta todo o e-mail em palavras isoladas
    Palavras = Split(TextoCorpo, " ")
    LotesValidos = "/SM2/DT3/PL1/VD1/GD1/"
    DataEntrada = Date

    ' =====================================================================
    ' 4. CONECTAR AO EXCEL (SEM TRAVAR O SEU PC)
    ' =====================================================================
    On Error Resume Next
    Set xlApp = GetObject(, "Excel.Application")
    On Error GoTo TrataErro ' Volta a ligar o alerta de erros
    
    If xlApp Is Nothing Then
        Set xlApp = CreateObject("Excel.Application")
        xlApp.Visible = False
        ExcelJaEstavaAberto = False
    Else
        ExcelJaEstavaAberto = True
    End If
    
    ArquivoJaEstavaAberto = False
    For Each MeuArquivo In xlApp.Workbooks
        If InStr(1, MeuArquivo.Name, "Sucata", vbTextCompare) > 0 Then
            ArquivoJaEstavaAberto = True
            Exit For
        End If
    Next MeuArquivo
    
    ' ABRE O SEU ARQUIVO NO CAMINHO CORRETO E COM A EXTENSÃO
    If Not ArquivoJaEstavaAberto Then
        Set MeuArquivo = xlApp.Workbooks.Open("C:\Users\anderson.vieira\Desktop\Nova pasta\Sucata.xlsm")
    End If
    
    ' Puxa a aba exata onde a tabela está
    Set ws = MeuArquivo.Sheets("RETE")

    ' =====================================================================
    ' 5. LER OS SERIAIS E PREENCHER TODAS AS COLUNAS
    ' =====================================================================
    LoteAtual = ""
    
    For i = 0 To UBound(Palavras)
        Token = Trim(Palavras(i))
        
        If Token <> "" Then
            ' Identifica se a palavra é um Lote
            If InStr(LotesValidos, "/" & Token & "/") > 0 Then
                LoteAtual = Token
            
            ' Se não for lote, MAS a memória já tiver um lote gravado, é um Serial!
            ElseIf LoteAtual <> "" Then
                
                ' Evita ler palavras muito curtas da assinatura
                If Len(Token) > 3 Then
                    UltimaLinha = ws.Cells(ws.Rows.Count, "A").End(-4162).Row + 1
                    
                    ' Preenche exatamente na ordem da sua tabela
                    ws.Range("A" & UltimaLinha).Value = Token        ' SERIAL
                    ws.Range("C" & UltimaLinha).Value = LoteAtual    ' LOTE
                    ws.Range("D" & UltimaLinha).Value = Equipe       ' EQUIPE
                    ws.Range("E" & UltimaLinha).Value = Deposito     ' DEPÓSITO
                    ws.Range("F" & UltimaLinha).Value = DataEntrada  ' DATA
                End If
                
            End If
        End If
    Next i

    ' =====================================================================
    ' 6. SALVAR E FINALIZAR CORRETAMENTE
    ' =====================================================================
    If Not ArquivoJaEstavaAberto Then
        MeuArquivo.Close SaveChanges:=True
    Else
        MeuArquivo.Save
    End If
    
    If Not ExcelJaEstavaAberto Then
        xlApp.Quit
    End If

    Set ws = Nothing
    Set MeuArquivo = Nothing
    Set xlApp = Nothing
    Exit Sub

TrataErro:
    MsgBox "Erro inesperado: " & Err.Description, vbCritical, "Erro"
    Set ws = Nothing
    Set MeuArquivo = Nothing
    Set xlApp = Nothing
End Sub
