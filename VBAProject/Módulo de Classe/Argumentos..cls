' ------------------------------------------------------
' Nome do Projeto     : Argumentos
' Tipo                : Class Module
' Objetivo            : Relaciona propriedades comuns em uma Collection
' Autor               : Adelson
' Data / Hora         : 10/11/2017 21:56
' ------------------------------------------------------
Private m_ColArgumentos As Collection

Private Sub Class_Initialize()
    Set m_ColArgumentos = New Collection
End Sub
' Adiciona um item na coleção
Sub add(VBAItem As Argumento, Optional VarKey As Variant)
    m_ColArgumentos.add VBAItem, VarKey
End Sub
'Remove um item na collecion
Sub remove(VarIndex As Variant)
    On Error Resume Next 'O metodo .Remove gera um erro caso o item nao exista
    m_ColArgumentos.remove VarIndex
End Sub
' Retorna um item atraves do indice ou da chave
Function Item(VarIndex As Variant) As Argumento
    On Error Resume Next 'O metodo .Item gera um erro caso o item nao exista
    Set Item = m_ColArgumentos.Item(VarIndex)
End Function
' Retorna o numero de itens
Property Get Count() As Long
    Count = m_ColArgumentos.Count
End Property
'Limpa a Colletion
Sub Clear()
    Set m_ColArgumentos = New Collection
End Sub
' Esta função habilita o suporte a enumeração dos items atraves de um "For Each". Deve haver um DispID-4.
Function NewEnum() As IUnknown
    Set NewEnum = m_ColArgumentos.[_NewEnum]
End Function
'Método que verifica se um item existe na collection
Function Exists(vIndex As Variant) As Boolean
    Dim vItem As Object
    On Error Resume Next
    Set vItem = m_ColArgumentos.Item(vIndex)
    Exists = Not vItem Is Nothing
End Function
