VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmCOnfiguracoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public ActivePage As Page
Private LastPage As Page
Private lngLastPage As Integer


Sub refreshStatusIcones(Optional ArquivoSumario As Integer = 1, Optional bOK As Boolean = False, Optional complemento_Icone As String)
40    Me.Controls("imgErro" & complemento_Icone).Visible = Not bOK
50    Me.Controls("imgAviso" & complemento_Icone).Visible = bOK
End Sub

Private Sub cmdSelecionarSumario_Click()
   Dim strEndereco As String
   ' strEndereco = PegaEndereco("Selecionar Pasta de Salvamento", msoFileDialogFolderPicker, False)
    If Not strEndereco = "Vazio" Then
        Call salvaValor("LocalSalvamentoArquivosTransferPricing", strEndereco)
    End If
    Call refreshStatusIcones(1, FolderExists(pegaValor("LocalSalvamentoArquivosTransferPricing")), "TP")
End Sub

Private Sub Fechar_Click()
    If Me.sfContainer.SourceObject = "frmModeloFluxoProcesso" Then
        If Me.sfContainer.Form.mpProcesso.Pages(mpProcesso.value).Name = Me.sfContainer.Form.pgSubForm.Name Then
            Me.sfContainer.Form.sfContainer.SourceObject = ""
            If Not Me.sfContainer.Form.ActivePage Is Nothing Then
                Fechar.Caption = "&Fechar"
                Me.sfContainer.Form.ActivePage.SetFocus
            Else
                Call FecharForm
            End If
        Else
            Call FecharForm
        End If
    Else
        Call FecharForm
    End If
End Sub

Sub FecharForm()
    Call DoCmd.Close(acForm, Me.Name, acSaveNo)
    If Not AuxForm.FormularioEstaAberto("frmPrincipal") Then
        Call AuxForm.AbrirFormulario("frmPrincipal")
    End If
End Sub
Private Sub Form_Close()
    Call AuxAplicacao.salvaValor("CaptionForm", "")
End Sub

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : Form_frmProc_ReceitasBBI.Form_Load()
' TIPO             : Sub
' DATA/HORA        : 29/08/2014 15:14
' CONSULTOR        : (ADELSON)/ TECNUN - Adelson Rosendo Marques da Silva
' CONTATO          : adelson@tecnun.com.br
' DESCRIÇÃO        : Inicializa as variáveis do formulário
'---------------------------------------------------------------------------------------
Private Sub Form_Load()
    'Tarja e imagem do logo
    'cxTarja.BackColor = TFWCliente.pegarVariavelAplicacao("color_bar")
    'Call TFWCliente.AuxAplicacao.DisplayImage(imgLogo, TFWCliente.pegarVariavelAplicacao("logo_file_path"))
    Me.Reload
    'Call AuxAplicacao.DisplayImage(Me.sfHeaderBar.Form.imgLogo, Nz(pegarVariavelAplicacao("logo_file_path")))
'    Call Me.sfHeaderBar.Form.Reload
    
80  On Error GoTo 0
90  Exit Sub
Form_Load_Error:
100 If VBA.Err <> 0 Then Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "Form_frmProc_ReceitasBBI.Form_Load()", VBA.Erl)
110 Exit Sub
120 Resume
End Sub
'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : Form_frmFluxoProcesso.activateTab()
' TIPO             : Sub
' DATA/HORA        : 18/11/2014 15:29
' CONSULTOR        : (ADELSON)/ TECNUN - Adelson Rosendo Marques da Silva
' CONTATO          : adelson@tecnun.com.br
' DESCRIÇÃO        : Torna a tab indicada ativa
'---------------------------------------------------------------------------------------

Sub Reload()
'    Call InicializarTFW
    Set frmCurrentForm = Me
    Me.Progresso.Width = 0
    mpProcesso.value = 1
    Call ExibirSubForm(Me.sfContainer, AuxAplicacao.pegaValor("subform"))
    'Me.Cabeçalho_Automático0.Value = AuxAplicacao.pegaValor("CaptionForm")
    Set currentParentControls = mpProcesso.Pages(mpProcesso.value)
    'Call CarregarImagensStatus(pgProcesso1, "CommandButton", Me)
    'Me.lblCompetencia.Requery
End Sub
