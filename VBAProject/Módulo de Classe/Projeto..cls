Private m_StrName As String
Private m_objParent As VBEditor
Private m_objProjeto As VBProject
Private m_objComponentes As Componentes
Private m_eProtectionMode As String

Private Sub Class_Terminate()
    Set m_objProjeto = Nothing
    Set m_objComponentes = Nothing
End Sub

Public Property Get Name() As String
    Name = m_StrName
End Property

Public Property Let Name(ByVal StrValue As String)
    m_StrName = StrValue
End Property

Public Property Get parent() As VBEditor
    Set parent = m_objParent
End Property

Public Property Get ProtectionMode() As String
    ProtectionMode = m_eProtectionMode
End Property

Public Property Let ProtectionMode(ByVal eProtectionMode As String)
    m_eProtectionMode = eProtectionMode
End Property

Public Property Set parent(objParent As VBEditor)
    Set m_objParent = objParent
End Property

Public Property Get Projeto() As VBProject
    Set Projeto = m_objProjeto
End Property

Public Property Set Projeto(objProjeto As VBProject)
    Set m_objProjeto = objProjeto
End Property

Public Property Get Componentes() As Componentes
    Set Componentes = m_objComponentes
End Property

Public Property Set Componentes(objComponentes As Componentes)
    Set m_objComponentes = objComponentes
End Property
