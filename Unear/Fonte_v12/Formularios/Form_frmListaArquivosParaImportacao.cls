VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmListaArquivosParaImportacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub btnSalvar_Click()
    Call salvaValor("Continuar_Importacao", -1)
    Call DoCmd.Close(acForm, Me.Name, acSaveYes)
End Sub

Private Sub btnSelecionarPasta_Click()
    Dim vArquivos As Variant
    vArquivos = CaixaDeDialogo(msoFileDialogFilePicker, _
                               "Escolha os arquivos para a importação dos dados", _
                               True, _
                               "Importar Arquivos", _
                               "Arquivos de Dados;*.xml,*.txt,*.xls*,*.*db|Todos os Arquvos;*.*")
    If Not VBA.IsEmpty(vArquivos) Then
        Me.sfPreValidacao.Form.AllowAdditions = True
        vResultValidacao = AdicionaArquivosSelecionadoParaImportacao(vArquivos)
        Me.sfPreValidacao.Requery
        Me.sfPreValidacao.Form.AllowAdditions = False
        'Me.Reload
    End If
End Sub

Private Sub chkMarcarTodos_Click()
    Call CurrentDb.Execute("UPDATE tblListaArquivosParaImportacao SET selecionar = " & VBA.CInt(Nz(chkMarcarTodos.value, 0)))
    Me.sfPreValidacao.Requery
End Sub

Private Sub Fechar_Click()
    Call salvaValor("Continuar_Importacao", 0)
    DoCmd.Close acForm, Me.Name, acSaveYes
End Sub

Private Sub Form_Load()
    Me.sfPreValidacao.Form.AllowAdditions = False
    Reload
End Sub

Sub Reload()
    Call CurrentDb.Execute("DELETE FROM tblListaArquivosParaImportacao")
    Me.sfPreValidacao.Requery
    Call salvaValor("Continuar_Importacao", 0)
    'Call DisplayImage(Me.LocalImagem, PegaEnderecoConfiguracoes & "\" & pegaValor("imagemStatusPrevalidacao") & ".png")
    'lblAviso.Caption = pegaValor("avisoPrevalidacao")
    Me.btnSelecionarPasta.Enabled = -1 'pegaValor("bEditarPasta") = "-1"
End Sub
