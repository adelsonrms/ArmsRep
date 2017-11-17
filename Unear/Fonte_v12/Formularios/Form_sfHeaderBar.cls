VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfHeaderBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Reload()
    Dim sCaption As String
    sCaption = pegaValor("CaptionForm")
    If sCaption = "" Then sCaption = pegarVariavelAplicacao("appName")
    Me.cbTitulo.value = sCaption
    cxTarja.BackColor = pegarVariavelAplicacao("color_bar")
    Call DisplayImage(Me.imgLogo, Nz(pegarVariavelAplicacao("logo_file_path")))
    Me.btnAppConfig.Caption = pegarVariavelAplicacao("appVersion")
    Me.btnAppConfig.Enabled = -1 'PegarVersao() <> 2
End Sub

Private Sub btnAppConfig_Click()
    Call AbrirFormulario("frmAppConfig", , acDialog)
    Reload
End Sub

Private Sub Form_Load()
    Reload
End Sub


