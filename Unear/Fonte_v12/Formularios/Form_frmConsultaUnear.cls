VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmConsultaUnear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private WithEvents cUnear As cBL_Unear
Attribute cUnear.VB_VarHelpID = -1

Private bTimer As Date

Private Sub btnExecutar_Click()
    Call TestaExecutar_CampanhasSelecionadas(Executar)
End Sub

Private Sub btnTestar_Click()
    Call TestaExecutar_CampanhasSelecionadas(Testar)
End Sub

Private Sub Comando579_Click()
    mBL_CRMUnear.AtualizarStatusDaLista
End Sub

Private Sub Form_Close()
    Me.TimerInterval = 0
End Sub

Private Sub Form_Load()
    Reload
End Sub

Function Reload()
    Call LimparEtapas
    sfEtapas.Form.Requery
    txtNomeSelecao = pegaValor("txtNomeSelecao")
    RefreshStatusCRM
End Function

Private Sub Form_Timer()
    Call RefreshStatusCRM
End Sub

Sub RefreshStatusCRM()
    If cUnear Is Nothing Then Set cUnear = New cBL_Unear
    cUnear.AtualizaInstancia
End Sub

Sub StartScheduler(pMsg As String, Optional pCor As Long)
    bTimer = Now
    Me.TimerInterval = 100
End Sub
