VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim vtm As Date

Private Sub Form_Load()
    vtm = VBA.Time
    'cxTarja.BackColor = TFWCliente.pegarVariavelAplicacao("color_bar")
   ' Call TFWCliente.AuxAplicacao.DisplayImage(imgLogo, TFWCliente.pegarVariavelAplicacao("logo_file_path"))
   If Me.OpenArgs <> "" Then
        If Me.OpenArgs = True Then
            Me.TimerInterval = 0
        End If
   End If
   
   On Error Resume Next
   Me.Controls("ctrAviso").Visible = False
   Me.lblAviso.Caption = "Aguardando inicialização...."
End Sub

Private Sub Form_Timer()
    Dim datafutura, contagemRegressiva
    datafutura = VBA.Date + (vtm + VBA.TimeSerial(0, 0, 3))
    contagemRegressiva = (datafutura - VBA.Now)
    If contagemRegressiva < 0 Then
        Me.lblAviso.Caption = "Contagem finalizada !"
        Me.TimerInterval = 0
        Call DoCmd.Close(acForm, Me.Name, acSaveNo)
    Else
        Me.lblAviso.Caption = "Inicializando a ferramenta, Fechamento automatico em...." & VBA.Format(contagemRegressiva, "hh:nn:ss")
    End If
End Sub

