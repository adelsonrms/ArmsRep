Private m_StrNome As String
Private m_StrTipo As eProcedureType
Private m_LonLinhasDeCodigo As Long
Private m_StrCódigo As String
Private m_Variaveis As Variaveis
Private m_bComtémIncompatibilidade As Boolean
Private m_sModificador As eTypeModifier
Private m_objArgumentos As Argumentos
Private m_objTipoPropriedade As eTipoPropriedade
Private m_sDeclaração As String
Private m_sRetorna As String
Private m_objVariaveis As Variaveis
Private m_sTipoMetodo As String
Private m_objParent As Componente
Private m_sVBNetCode As String
Private m_sTypeName As String
Private m_sTypeNamo As String

Private Sub Class_Terminate()
    Set m_objArgumentos = Nothing
    Set m_objVariaveis = Nothing
End Sub

Public Property Get ComtémIncompatibilidade() As Boolean
    ComtémIncompatibilidade = m_bComtémIncompatibilidade
End Property

Public Property Let ComtémIncompatibilidade(ByVal bValue As Boolean)
    m_bComtémIncompatibilidade = bValue
End Property

Public Property Get Código() As String
    Código = m_StrCódigo
End Property

Public Property Let Código(ByVal StrValue As String)
    m_StrCódigo = StrValue
End Property

Public Property Get LinhasDeCodigo() As Long
    LinhasDeCodigo = m_LonLinhasDeCodigo
End Property

Public Property Let LinhasDeCodigo(ByVal LonValue As Long)
    m_LonLinhasDeCodigo = LonValue
End Property

Public Property Get Tipo() As eProcedureType
    Tipo = m_StrTipo
End Property

Public Property Let Tipo(ByVal StrValue As eProcedureType)
    m_StrTipo = StrValue
End Property

Public Property Get nome() As String
    nome = m_StrNome
End Property

Public Property Let nome(ByVal StrValue As String)
    m_StrNome = StrValue
End Property

Public Property Get Modificador() As eTypeModifier
    Modificador = m_sModificador
End Property

Public Property Let Modificador(ByVal sModificador As eTypeModifier)
    m_sModificador = sModificador
End Property

Public Property Get Argumentos() As Argumentos
    Set Argumentos = m_objArgumentos
End Property

Public Property Set Argumentos(objArgumentos As Argumentos)
    Set m_objArgumentos = objArgumentos
End Property

Public Property Get TipoPropriedade() As eTipoPropriedade
    TipoPropriedade = m_objTipoPropriedade
End Property

Public Property Let TipoPropriedade(objTipoPropriedade As eTipoPropriedade)
    m_objTipoPropriedade = objTipoPropriedade
End Property

Public Property Get Declaração() As String
    Declaração = m_sDeclaração
End Property

Public Property Let Declaração(ByVal sDeclaração As String)
    m_sDeclaração = sDeclaração
End Property

Public Property Get Retorna() As String

    Retorna = m_sRetorna

End Property

Public Property Let Retorna(ByVal sRetorna As String)

    m_sRetorna = sRetorna

End Property


Public Property Get Variaveis() As Variaveis

    Set Variaveis = m_objVariaveis

End Property

Public Property Set Variaveis(objVariaveis As Variaveis)

    Set m_objVariaveis = objVariaveis

End Property



Public Property Get TipoMetodo() As String

    TipoMetodo = m_sTipoMetodo

End Property

Public Property Let TipoMetodo(ByVal sTipoMetodo As String)

    m_sTipoMetodo = sTipoMetodo

End Property

Public Property Get parent() As Componente
    Set parent = m_objParent
End Property

Public Property Set parent(objParent As Componente)
    Set m_objParent = objParent
End Property


Public Property Get VBNetCode() As String
    VBNetCode = m_sVBNetCode
End Property

Public Property Let VBNetCode(ByVal sVBNetCode As String)
    m_sVBNetCode = sVBNetCode
End Property

Public Property Get TypeNamo() As String

    TypeNamo = m_sTypeNamo

End Property

Public Property Let TypeNamo(ByVal sTypeNamo As String)

    m_sTypeNamo = sTypeNamo

End Property
