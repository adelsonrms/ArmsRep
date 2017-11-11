Private m_StrName As String
Private m_StrReferenceType As String
Private m_bIsOptional As Boolean
Private m_StrDataType As String
Private m_VarDefaulValue As Variant
Private m_objParent As Procedimento


Public Property Get DefaulValue() As Variant
    DefaulValue = m_VarDefaulValue
End Property

Public Property Let DefaulValue(ByVal VarValue As Variant)
    m_VarDefaulValue = VarValue
End Property

Public Property Get DataType() As String
    DataType = m_StrDataType
End Property

Public Property Let DataType(ByVal StrValue As String)
    m_StrDataType = StrValue
End Property

Public Property Get IsOptional() As Boolean
    IsOptional = m_bIsOptional
End Property

Public Property Let IsOptional(ByVal bValue As Boolean)
    m_bIsOptional = bValue
End Property

Public Property Get ReferenceType() As String
    ReferenceType = m_StrReferenceType
End Property

Public Property Let ReferenceType(ByVal StrValue As String)
    m_StrReferenceType = StrValue
End Property

Public Property Get Name() As String
    Name = m_StrName
End Property

Public Property Let Name(ByVal StrValue As String)
    m_StrName = StrValue
End Property


Public Property Get parent() As Procedimento
    Set parent = m_objParent
End Property

Public Property Set parent(objParent As Procedimento)
    Set m_objParent = objParent
End Property


