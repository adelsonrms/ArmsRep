VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTFW_Aplicacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Private m_sIDAplicacao As String
Private m_sNomeAplicacao As String
Private m_sVersaoAplicacao As String
Private m_sCliente As String
Private m_lCorTarja As Long
Private m_sArquivoIcone As String
Private m_sArquivoLogo As String
Private m_objCurrentDatabase As Object
Private m_sNomeRibbonPrincipal As String

Private Sub Class_Initialize()
    Set m_objCurrentDatabase = Access.Application.CodeDb
End Sub

Private Sub Class_Terminate()
    Set m_objCurrentDatabase = Nothing
End Sub

Public Property Get CurrentDatabase() As Object
    Set CurrentDatabase = m_objCurrentDatabase
End Property

Public Property Set CurrentDatabase(objCurrentDatabase As Object)
    Set m_objCurrentDatabase = objCurrentDatabase
End Property

Public Property Get idAplicacao() As String
    idAplicacao = m_sIDAplicacao
End Property

Public Property Let idAplicacao(ByVal sIDAplicacao As String)
    m_sIDAplicacao = sIDAplicacao
End Property

Public Property Get NomeAplicacao() As String
    NomeAplicacao = m_sNomeAplicacao
End Property

Public Property Let NomeAplicacao(ByVal sNomeAplicacao As String)
    m_sNomeAplicacao = sNomeAplicacao
End Property

Public Property Get VersaoAplicacao() As String
    VersaoAplicacao = m_sVersaoAplicacao
End Property

Public Property Let VersaoAplicacao(ByVal sVersaoAplicacao As String)
    m_sVersaoAplicacao = sVersaoAplicacao
End Property

Public Property Get Cliente() As String
    Cliente = m_sCliente
End Property

Public Property Let Cliente(ByVal sCliente As String)
    m_sCliente = sCliente
End Property

Public Property Get CorTarja() As Long
    CorTarja = m_lCorTarja
End Property

Public Property Let CorTarja(ByVal lCorTarja As Long)
    m_lCorTarja = lCorTarja
End Property

Public Property Get ArquivoIcone() As String
    ArquivoIcone = m_sArquivoIcone
End Property

Public Property Let ArquivoIcone(ByVal sArquivoIcone As String)
    m_sArquivoIcone = sArquivoIcone
End Property

Public Property Get ArquivoLogo() As String
    ArquivoLogo = m_sArquivoLogo
End Property

Public Property Let ArquivoLogo(ByVal sArquivoLogo As String)
    m_sArquivoLogo = sArquivoLogo
End Property

Public Property Get NomeRibbonPrincipal() As String
    NomeRibbonPrincipal = m_sNomeRibbonPrincipal
End Property

Public Property Let NomeRibbonPrincipal(ByVal sNomeRibbonPrincipal As String)
    m_sNomeRibbonPrincipal = sNomeRibbonPrincipal
End Property
