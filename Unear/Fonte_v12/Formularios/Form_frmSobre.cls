VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmSobre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'Private Sub CabeçalhoDoFormulário_Click()
'Detalhe_Click
'End Sub
'
'Private Sub Detalhe_Click()
'DoCmd.Close acForm, Me.Name, acSaveNo
'End Sub

Private Sub Form_Click()
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub Form_Load()
On Error GoTo TratarErro
    Call Inicializar_Globais
    Call AuxForm.Atualizar_Campo_Versao(Me.lblVersao)
On Error GoTo 0
Exit Sub
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "Form_frmEscolherPeriodo.Form_Load", Erl)
End Sub

'Private Sub RodapéDoFormulário_Click()
'Detalhe_Click
'End Sub
