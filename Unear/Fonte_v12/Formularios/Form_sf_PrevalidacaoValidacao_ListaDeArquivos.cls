VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sf_PrevalidacaoValidacao_ListaDeArquivos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Anexo_FileData_AttachmentCurrent()

End Sub

Private Sub Anexo_FileData_Click()
    VBA.MsgBox Me.Código.value
End Sub

Private Sub Anexo_FileData_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    'VBA.MsgBox
End Sub

Private Sub Anexo_FileData_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    VBA.MsgBox Button
    VBA.MsgBox X
    VBA.MsgBox y
    
End Sub
