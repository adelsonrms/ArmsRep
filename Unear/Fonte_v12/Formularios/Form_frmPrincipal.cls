VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub Reload()
    Dim sForm As String
    If pegaValor("TrocarForm") = "-1" Then
        Call ExibirSubForm(Me.sfContainer_Inicial, pegaValor("subform"))
    Else
        sForm = pegaValor("sfContainer_Inicial")
        If sForm <> "" Then
            If FormsExists(sForm) Then
                Me.sfContainer_Inicial.SourceObject = sForm
                If Me.sfContainer_Inicial.SourceObject = "" Then
                    salvaValor "sfContainer_Inicial", "frmPrincipalTFW"
                    Me.sfContainer_Inicial.SourceObject = "frmPrincipalTFW"
                End If
            End If
        End If
        
    End If
End Sub

Private Sub Form_Load()
    Reload
End Sub
