Attribute VB_Name = "VariaveisEConstantes"
Option Compare Database
Option Explicit

'********************VARIAVEIS GLOBAIS********************
Public Conexao          As ConexaoDB
Public Relatorio        As Relatorios
Public Resultado        As VBA.VbMsgBoxResult
Public versao           As String

Public m_objCurrentDB_Cliente            As Object
Public m_objCurrentProject_Cliente       As Access.CurrentProject
Public m_objConnection_Cliente           As Object 'ADODB.Connection

Public Const appName = "Nome do Aplicativo"

'********************ENUM********************
Public Enum TipoVerificacao
    TabelasLocais = 1
    TabelasCriadas = 2
End Enum

Public Enum TipoData
    DataInicial = 0
    DataFinal = 1
End Enum

Public Enum tipoEndereco
    Pastas = 0
    arquivos = 1
End Enum


Public Enum OutlookAcao
    Rascunho = 1
    Enviar = 2
End Enum

'********************VARIAVEIS GLOBAIS********************
Private m_FalhaNoLink        As Boolean

Public Property Get FalhaNoLink() As Boolean
    FalhaNoLink = m_FalhaNoLink
End Property
Public Property Let FalhaNoLink(ByVal valor As Boolean)
    m_FalhaNoLink = valor
End Property
'*********************************************************

Public Property Get CurrentDB_Cliente() As Object
    If m_objCurrentDB_Cliente Is Nothing Then Set m_objCurrentDB_Cliente = Access.Application.CurrentDb
    Set CurrentDB_Cliente = m_objCurrentDB_Cliente
End Property

Public Property Set CurrentDB_Cliente(objCurrentDB_Cliente As Object)
    Set m_objCurrentDB_Cliente = objCurrentDB_Cliente
End Property

Public Property Get CurrentProject_Cliente() As Access.CurrentProject
    If CurrentProject_Cliente Is Nothing Then Set m_objCurrentProject_Cliente = Access.Application.CurrentProject
    Set CurrentProject_Cliente = m_objCurrentProject_Cliente
End Property

Public Property Set CurrentProject_Cliente(objCurrentProject_Cliente As Access.CurrentProject)
    Set m_objCurrentProject_Cliente = objCurrentProject_Cliente
End Property

Public Property Get Connection_Cliente() As Object ' DAO.Connection
    If CurrentProject_Cliente Is Nothing Then Set m_objConnection_Cliente = Access.Application.CurrentProject.Connection
    Set Connection_Cliente = m_objConnection_Cliente
End Property

Public Property Set Connection_Cliente(objConnection_Cliente As Object)
    Set m_objConnection_Cliente = objConnection_Cliente
End Property

Function ReiniciarClasseExcel() As cTFW_Excel
    If oExcel Is Nothing Then Set oExcel = New cTFW_Excel
End Function
