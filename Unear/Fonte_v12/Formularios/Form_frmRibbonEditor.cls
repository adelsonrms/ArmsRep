VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmRibbonEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Load()
    Me.sfTabs.Form.NavigationButtons = False
    Me.sfGrupo.Form.NavigationButtons = False
    Me.sfControles.Form.NavigationButtons = False
End Sub
