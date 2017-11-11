
Private m_sFileConfig As String

Public Property Get FileConfig() As String
    FileConfig = m_sFileConfig
End Property

Public Property Let FileConfig(ByVal sNewValue As String)
    m_sFileConfig = sNewValue
End Property

Public Function BackupProject(Projeto As Projeto) As Boolean
    Dim bResult As Boolean
    bResult = False
    Dim strPasta As String
    strPasta = ""
    
    Dim Comp As Componente
    
    For Each Comp In Projeto.Componentes
        'Comp.Tipo
    Next
        
    BackupProject = bResult
End Function

