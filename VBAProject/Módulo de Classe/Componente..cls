
Private m_StrNome As String
Private m_StrTipo As eComponentType
Private m_VBPVBComponentParent As VBComponent
Private m_Procedimentos As Procedimentos
Private m_sDeclarationCode As String
Private m_objVariaveisGlobais As Variaveis
Private m_sVBNetCode As String

Private m_sCodigo As String

Private m_sExtensao As String

Private m_sTipoNome As String




Private Sub Class_Terminate()
    Set m_objVariaveisGlobais = Nothing
End Sub

Public Property Get Procedimentos() As Procedimentos
    Set Procedimentos = m_Procedimentos
End Property

Public Property Set Procedimentos(ByVal ClsValue As Procedimentos)
    Set m_Procedimentos = ClsValue
End Property

Public Property Get VBComponentParent() As VBComponent
    Set VBComponentParent = m_VBPVBComponentParent
End Property

Public Property Get DeclarationCode() As String
    DeclarationCode = m_sDeclarationCode
End Property

Public Property Let DeclarationCode(ByVal sDeclarationCode As String)
    m_sDeclarationCode = sDeclarationCode
End Property

Public Property Set VBComponentParent(ByVal VBPValue As VBComponent)
    Set m_VBPVBComponentParent = VBPValue
End Property

Public Property Get Tipo() As eComponentType
    Tipo = m_StrTipo
End Property

Public Property Get TipoNome() As String
    TipoNome = m_sTipoNome
End Property

Public Property Let TipoNome(ByVal sNewValue As String)
    m_sTipoNome = sNewValue
End Property

Public Property Get Extensao() As String
    Extensao = m_sExtensao
End Property

Public Property Let Extensao(ByVal sNewValue As String)
    m_sExtensao = sNewValue
End Property

Public Property Let Tipo(ByVal StrValue As eComponentType)
    m_StrTipo = StrValue
End Property

Public Property Get nome() As String
    nome = m_StrNome
End Property

Public Property Let nome(ByVal StrValue As String)
    m_StrNome = StrValue
End Property

Public Property Get VariaveisGlobais() As Variaveis
    Set VariaveisGlobais = m_objVariaveisGlobais
End Property

Public Property Set VariaveisGlobais(objVariaveisGlobais As Variaveis)
    Set m_objVariaveisGlobais = objVariaveisGlobais
End Property

Public Property Get Codigo() As String
    With m_VBPVBComponentParent.CodeModule
        If .CountOfLines > 0 Then
            m_sCodigo = .Lines(1, .CountOfLines)
        End If
    End With
    Codigo = m_sCodigo
End Property
