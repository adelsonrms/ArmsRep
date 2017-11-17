VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfStatusBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub btnDataBase_Click()
    Call AbrirSubFormulario("frmCaminhoBE", "Configuração Back-End")
End Sub

Private Sub Form_Load()
    Call Me.Reload
End Sub

Public Sub Reload()
    'btnDataBase.enabled = PegarVersao() = Desenvolvimento
   ' Call AtualizaFlagStatusArquivo(Me, FolderExists(Me.txtPastaDB.Value))
End Sub


 
 
