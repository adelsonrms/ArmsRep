VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmProgresso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Modulo....: Form_frmProgresso/ Documento VBA
' Rotina....: Form_Load / Sub
' Autor.....: Victor Félix
' Contato...: victor.santos@mondial.com.br
' Data......: 29/5/2013
' Empresa...: Mondial Tecnologia em Informática LTDA.
' Descrição.: Evento de Load do Form. Zera o comprimento do controle Imagem
'---------------------------------------------------------------------------------------
Private Sub Form_Load()
    Me.Progresso.Width = 0
End Sub
