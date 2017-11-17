VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAppConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim strFile    As String

Private Type tChooseColor
    lStructSize As LongPtr
    hwndOwner As LongPtr
    hInstance As LongPtr
    rgbResult As LongPtr
    lpCustColors As String
    flags As LongPtr
    lCustData As LongPtr
    lpfnHook As LongPtr
    lpTemplateName As String
End Type

#If VBA7 Then
    Private Declare PtrSafe Function ChooseColorA Lib "comdlg32.dll" (pChoosecolor As tChooseColor) As LongPtr
#Else
    Private Declare Function ChooseColorA Lib "comdlg32.dll" (pChoosecolor As tChooseColor) As LongPtr
#End If

Private Sub btnAlterarColor_Click()
    Dim lngColor As Long
    lngColor = CLng(ShowColor)
    If lngColor <> -1 Then
        If color_bar.value <> lngColor Then
            color_bar.value = lngColor
            DoCmd.RunCommand acCmdSaveRecord
            Call ReloadAllForms
            color_bar.BackColor = lngColor
        End If
    End If
End Sub

Private Sub btnAlterarVersao_Click()
    If salvarVersaoAplicacao() = True Then
        VBA.MsgBox "Versão Alterada !. É necessário reiniciar a aplicação", VBA.vbInformation
        'Call FinalizaAccess
        VBA.Shell "explorer """ & CurrentDb.Name & """"
    End If
    Form_Load
    Call ModoProducao(PegarVersao() = Producao)
End Sub

Private Sub btnEditarRibbon_Click()
    AbrirRibbonEditor
End Sub

Private Sub btnHistorico_Click()
    Call FecharObjeto(Me.Name, acSaveNo)
    Call AbrirSubFormulario("tblVersao", "Historico de Versão")
End Sub

Private Sub btnMensagens_Click()
    Call FecharObjeto(Me.Name, acSaveNo)
    Call AbrirSubFormulario("tblMensagens", "Congiruações de Mensagens do Sistema", "frmPrincipal", True, True)
End Sub

Private Sub btnSelecionarLogo_Click()
    Dim strNewPath As String
    strFile = CaixaDeDialogo(msoFileDialogOpen, "Selecionar o arquivo de Logo da empresa Cliente", False, "Escolher", "Arquivos de Imagens;*.jpg,*.png,*.bmp")
    strNewPath = PegaEnderecoConfiguracoes & "\logo.jpg"
    If strFile = "" Then Exit Sub
    logo_file_path.value = strNewPath
    DoCmd.RunCommand acCmdSaveRecord
    Call VBA.FileCopy(strFile, strNewPath)
    'Anexa o arquivo na ferramenta
    If VBA.Dir(strFile) <> "" Then
        'Copia para o local de configurações
        'Nao anexar mais o arquivo
        Call AnexaArquivoApp("tblApp", strNewPath)
        Call DisplayImage(imgLogo, strNewPath)
        Call ReloadAllForms
    End If
End Sub

Private Sub btnSelecionarIcone_Click()
    Dim strNewPath As String
    strFile = CaixaDeDialogo(msoFileDialogOpen, "Selecionar o arquivo de Icone", False, "Escolher", "Icones;*.ico")
    If strFile = "" Then Exit Sub
    strNewPath = PegaEnderecoConfiguracoes & "\icone.ico"
    icon_file_path.value = strNewPath
    DoCmd.RunCommand acCmdSaveRecord
    Call VBA.FileCopy(strFile, strNewPath)
    'Anexa o arquivo na ferramenta
    If VBA.Dir(strFile) <> "" Then
        'Copia para o local de configurações
        Call AnexaArquivoApp("tblApp", strNewPath)
        Call DisplayImage(imgIcone, strNewPath)
        Call ReloadAllForms
    End If
End Sub

Private Sub Comando70_Click()
    Call AbrirSubFormulario("frmOpcoes", "Configurações")
End Sub

Private Sub Comando81_Click()
    Call CurrentDb.Execute("DELETE FROM tblVariaveis")
    Call VBA.MsgBox("Todas as variaveis Locais foram excluidas. Reiniciar a aplicação !", vbExclamation)
End Sub

Private Sub Form_Close()
    Call ReloadAllForms
End Sub

Private Sub Form_Load()
     Me.lblVersao = NomeVersaoAtual
     color_bar.BackColor = Nz(color_bar.value, 0)
     Call DisplayImage(Me.imgLogo, Nz(logo_file_path.value))
     Call DisplayImage(Me.imgIcone, Nz(icon_file_path.value))
     'Call DisplayImage(Me.imgBGInicio, Nz(img_bg_inico.Value))
    grpOpcaoForm.value = pegaValor("grpOpcaoForm_Value", 1)
    Me.cmbForms.value = pegaValor("sfContainer_Inicial", "")
    Call grpOpcaoForm_AfterUpdate
     Me.appVersion.Requery
End Sub

Private Sub Texto71_AfterUpdate()
    Call CurrentDb.Execute("UPDATE tblVersao set versao = Left(versao, len(versao)-2) & '." & Left(Texto71.value, 1) & "' WHERE versao = '" & NumeroVersaoAtual & "'")
    Me.appVersion.Requery
    Call ModoProducao(PegarVersao() = Producao)
    Call Form_Load
End Sub

Private Sub cmbForms_AfterUpdate()
    Call salvaValor("sfContainer_Inicial", cmbForms.value)
End Sub

Private Sub grpOpcaoForm_AfterUpdate()
    Select Case grpOpcaoForm.value
    Case 1:
        Call salvaValor("sfContainer_Inicial", "frmPrincipalTFW")
    Case 2:
        cmbForms_AfterUpdate
    End Select
    Call salvaValor("grpOpcaoForm_Value", grpOpcaoForm.value)
End Sub

'Exibe a paleta de cores do windows
Private Function ShowColor() As Long
    Dim tColor As tChooseColor
    Dim lThisColor As Long
    Dim abytCustomColors(0 To 16 * 4 - 1) As Byte
    
    For lThisColor = LBound(abytCustomColors) To UBound(abytCustomColors)
        abytCustomColors(lThisColor) = 0
    Next

    tColor.lStructSize = Len(tColor)
    tColor.hwndOwner = 0
    tColor.hInstance = 1
    'Converte para Unicode
    tColor.lpCustColors = StrConv(abytCustomColors, VBA.vbUnicode)
    tColor.flags = 0

    'Exibe a caixa de dialogo
    If ChooseColorA(tColor) <> 0 Then
        ShowColor = tColor.rgbResult
    Else
        ShowColor = -1
    End If
End Function

