VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmPrevalidacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub btnSalvar_Click()
    Call salvaValor("Continuar_Importacao", -1)
    Call DoCmd.Close(acForm, Me.Name, acSaveYes)
End Sub

Private Sub chkSomenteErros_Click()
    If chkSomenteErros.value Then Call salvaValor("TelaPrevalidacao_FiltroStatus", "erro") Else Call salvaValor("TelaPrevalidacao_FiltroStatus", "")
    Call FiltrarStatus
End Sub

Private Sub btnSelecionarPasta_Click()
    Dim vArquivos As Variant
    vArquivos = CaixaDeDialogo(msoFileDialogFilePicker, _
                               "Escolha os arquivos para a importação dos dados", _
                               True, _
                               "Importar Arquivos", _
                               "Arquivos de Dados;*.txt,*.xls*,*.*db|Todos os Arquvos;*.*")
    If Not VBA.IsEmpty(vArquivos) Then
        vResultValidacao = preValidacaoRegraArquivos(vArquivos, True)
        Call verificaResultadoValidacao(vResultValidacao)
        Me.Reload
    End If
End Sub

Private Sub Fechar_Click()
    Call salvaValor("Continuar_Importacao", 0)
    DoCmd.Close acForm, Me.Name, acSaveYes
End Sub

Private Sub Form_Load()
    Reload
End Sub

Sub Reload()
    Call salvaValor("Continuar_Importacao", 0)
    Call DisplayImage(Me.LocalImagem, PegaEnderecoConfiguracoes & "\" & pegaValor("imagemStatusPrevalidacao") & ".png")
    lblAviso.Caption = pegaValor("avisoPrevalidacao")
    Me.btnSelecionarPasta.Enabled = pegaValor("bEditarPasta") = "-1"
    If pegaValor("TelaPrevalidacao_FiltroStatus") <> "" Then Quadro194.value = 1: Quadro194_AfterUpdate
    Me.txtTotalAnalisados.Requery
    Me.txtTotalIdentificados.Requery
    Me.txtTotalNaoIdentificados.Requery
End Sub

Sub FiltrarStatus()
    If pegaValor("TelaPrevalidacao_FiltroStatus") <> "" Then
        Me.sfPreValidacao.Form.Filter = "status = '" & pegaValor("TelaPrevalidacao_FiltroStatus") & "'"
        Me.sfPreValidacao.Form.FilterOn = True
    Else
        Me.sfPreValidacao.Form.FilterOn = False
    End If
End Sub

Private Sub Quadro194_AfterUpdate()
    Call salvaValor("TelaPrevalidacao_FiltroStatus", Quadro194.Controls("opt" & Quadro194.value).tag)
    Call FiltrarStatus
End Sub

Private Sub Seleção148_Click()
    If Seleção148.value Then Call salvaValor("TelaPrevalidacao_FiltroStatus", "ok") Else Call salvaValor("TelaPrevalidacao_FiltroStatus", "")
    Call FiltrarStatus
End Sub
