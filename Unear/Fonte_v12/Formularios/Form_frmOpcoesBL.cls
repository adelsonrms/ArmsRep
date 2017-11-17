VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmOpcoesBL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Comando17_Click()
    If VBA.Dir(PegaEnderecoConfiguracoes() & "\PosicaoMouse.exe") = "" Then
        MsgBox "Aplicativo 'MousePos.exe' não localizado na pasta '\BDs\Configurações' !", vbExclamation
        Exit Sub
    End If
    Call VBA.Shell(PegaEnderecoConfiguracoes() & "\PosicaoMouse.exe", vbNormalFocus)
End Sub

Private Sub Comando23_Click()
    Me.PerfilPadraoMousePointer = VBA.Environ("ComputerName")
End Sub

Private Sub Comando27_Click()
    PastaRaiz.value = PegaEndereco_Programa()
End Sub

Private Sub Form_Load()
    If Nz(PastaDownloads_IE.value) = "" Then
        PastaDownloads_IE.value = AuxFileSystem.PegarPasta(AuxFileSystem.PegarPasta(AuxFileSystem.PegarPasta(VBA.Environ("temp")))) & "\Downloads\"
    End If
End Sub
