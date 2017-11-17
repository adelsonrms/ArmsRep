VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DEPARAs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private CBD_Atual                       As ConexaoDB
Private m_Gerentes_PorCod               As Object 'Scripting.Dictionary
Private m_Produtos_PorCod               As Object 'Scripting.Dictionary

Public Property Get Gerentes_PorCod() As Object 'Scripting.Dictionary
    Set Gerentes_PorCod = m_Gerentes_PorCod
End Property
Public Property Set Gerentes_PorCod(ByRef valor As Object)
    Set m_Gerentes_PorCod = valor
End Property

Public Property Get Produtos_PorCod() As Object 'Scripting.Dictionary
    Set Produtos_PorCod = m_Produtos_PorCod
End Property
Public Property Set Produtos_PorCod(ByRef valor As Object)
    Set m_Produtos_PorCod = valor
End Property

Private Sub Class_Initialize()
On Error GoTo TratarErro
    Set CBD_Atual = New ConexaoDB
    Me.PopularInformacoesDicionario
On Error GoTo 0
Exit Sub
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "Propriedades.Class_Initialize", Erl)
End Sub

'---------------------------------------------------------------------------------------
' Modulo....: Propriedades / Módulo de classe
' Rotina....: PopularInformacoesDicionario / Sub
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 19/08/2013
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina que inicializa os dicionarios preenchendo-os com as devidas informações
'             dos DEPARAS
'---------------------------------------------------------------------------------------
Public Sub PopularInformacoesDicionario()
On Error GoTo TratarErro

    With CBD_Atual
        Set Me.Gerentes_PorCod = .PegarDicionario("Pegar_DEPARA_Gerente_PorCod")
        Set Me.Produtos_PorCod = .PegarDicionario("Pegar_DEPARA_Produtos_PorCod")
    End With

On Error GoTo 0
Exit Sub
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "Propriedades.PopularInformacoesDicionario", Erl)
End Sub
