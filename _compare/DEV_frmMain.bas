VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMain 
   Caption         =   "Entrada de Seriais"
   ClientHeight    =   7860
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3060
   OleObjectBlob   =   "DEV_frmMain.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancelar_Click()

    Call clearAll
    frmMain.Hide

End Sub

Private Sub btnOk_Click()

    ' Sem selecao de acao
    If cboAcao.ListIndex = -1 Then
        MsgBox "Selecione um processo valido"
        Exit Sub
    End If

    ' Adicionar seriais
    If cboAcao.ListIndex = 0 Then
        If CheckSerials(frmMain.txtSeries.Text) = True Then
            Exit Sub
        End If
        Call add2Plan(wsRETE, txtSeries, cboEquipe.Text, cboLote.Text)
        Call clearAll
        Exit Sub
    End If

    ' Consultar status IQ09
    If cboAcao.ListIndex = 1 Then
        Call SapIQ09
        Exit Sub
    End If

    ' Consultar OSME
    If cboAcao.ListIndex = 2 Then
        Call SapOSME
        Exit Sub
    End If
    
    ' Fechar lote RETE
    If cboAcao.ListIndex = 3 Then
        If AreAllSerialsInSheet(txtSeries) = False Then
            Exit Sub
        End If
        Call SetLot(txtSeries, cboLote)
        Call clearAll
        Exit Sub
    End If
    
    ' Enviar Lote RETE
    If cboAcao.ListIndex = 4 Then
        If AreAllSerialsInSheet(txtSeries) = False Then
            Exit Sub
        End If
        Call SendLotRETE(txtSeries, txtDoc)
        Call clearAll
        Exit Sub
    End If
    
    ' Executar RETE MIGO
    If cboAcao.ListIndex = 5 Then
        Call SapRETE(txtSeries)
        Call clearAll
        Exit Sub
    End If
    
    ' Preencher OSME (IW39)
    If cboAcao.ListIndex = 6 Then
        Call SapPreencherOSME
        Exit Sub
    End If
    
    ' Preencher Operacoes IW39
    If cboAcao.ListIndex = 7 Then
        Call PreencherOperacoesIW39
        Exit Sub
    End If
    
    ' Retirada RETE Z77 (MIGO)
    If cboAcao.ListIndex = 8 Then
        Call SapRetiradaReteZ77
        Exit Sub
    End If
    
    ' Retirada RETE Z77 Alternativa
    If cboAcao.ListIndex = 9 Then
        Call SapRetiradaReteZ77_Alt
        Exit Sub
    End If
    
End Sub

Private Sub UserForm_Activate()

    'Load Combo
    Call comboPop(frmMain.cboEquipe, wsConf, "A")
    Call comboPop(frmMain.cboLote, wsRETE, "C")
    
    ' Popula o combo de acao
    With cboAcao
        .AddItem "Entrada de Seriais"
        .AddItem "Consultar Status SAP"
        .AddItem "Consultar OSME"
        .AddItem "Fechar Lote RETE"
        .AddItem "Enviar Lote RETE"
        .AddItem "Executar RETE MIGO"
        .AddItem "Preencher OSME (IW39)"
        .AddItem "Preencher Operacoes IW39"
        .AddItem "Retirada RETE Z77"
        .AddItem "Retirada RETE Z77 Alternativa"
    End With

End Sub

Private Sub clearAll()

    cboAcao.Text = ""
    cboEquipe.Text = ""
    cboLote.Text = ""
    txtDoc.Text = ""
    txtSeries.Text = ""

End Sub
