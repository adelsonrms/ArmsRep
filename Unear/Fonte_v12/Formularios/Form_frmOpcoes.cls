VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmOpcoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    grpOpcaoForm.value = pegaValor("grpOpcaoForm_Value", 1)
    Me.cmbForms.value = pegaValor("sfContainer_Inicial", "")
    Call grpOpcaoForm_AfterUpdate
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
