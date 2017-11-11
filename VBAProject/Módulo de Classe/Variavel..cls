Private m_StrNome As String
Private m_StrTipoDeDados As String
Private m_StrObjectName As String
Private m_StrModifier As String
Private m_bIsObject As Boolean
Private m_objProcParent As Procedimento
Private m_objParent As Procedimento

Private Sub Class_Terminate()

    Set m_objProcParent = Nothing

End Sub


Public Property Get IsObject() As Boolean
    IsObject = m_bIsObject
End Property

Public Property Let IsObject(ByVal bValue As Boolean)
    m_bIsObject = bValue
End Property


Public Property Get Modifier() As String
    Modifier = m_StrModifier
End Property

Public Property Let Modifier(ByVal StrValue As String)
    m_StrModifier = StrValue
End Property

Public Property Get ObjectName() As String
    ObjectName = m_StrObjectName
End Property

Public Property Let ObjectName(ByVal StrValue As String)
    m_StrObjectName = StrValue
End Property



Public Property Get TipoDeDados() As String
    TipoDeDados = m_StrTipoDeDados
End Property

Public Property Let TipoDeDados(ByVal StrValue As String)
    m_StrTipoDeDados = StrValue
End Property

Public Property Get nome() As String
    nome = m_StrNome
End Property

Public Property Let nome(ByVal StrValue As String)
    m_StrNome = StrValue
End Property

Public Property Get procParent() As Procedimento

    Set procParent = m_objProcParent

End Property

Public Property Set procParent(objProcParent As Procedimento)

    Set m_objProcParent = objProcParent

End Property

Public Property Get parent() As Procedimento
    Set parent = m_objParent
End Property

Public Property Set parent(objParent As Procedimento)
    Set m_objParent = objParent
End Property


