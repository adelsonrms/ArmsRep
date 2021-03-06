' ------------------------------------------------------
' Nome do Variavel     : Variavels
' Tipo                : Class Module
' Objetivo            : Relaciona propriedades comuns em uma Collection
' Autor               : Adelson
' Data / Hora         : 10/11/2017 21:56
' ------------------------------------------------------
Private m_ColVariavels As Collection

Private Sub Class_Initialize()
    Set m_ColVariavels = New Collection
End Sub
' Adiciona um item na cole��o
Sub add(VBAItem As Variavel, Optional VarKey As Variant)
    m_ColVariavels.add VBAItem, VarKey
End Sub
'Remove um item na collecion
Sub remove(VarIndex As Variant)
    On Error Resume Next 'O metodo .Remove gera um erro caso o item nao exista
    m_ColVariavels.remove VarIndex
End Sub
' Retorna um item atraves do indice ou da chave
Function Item(VarIndex As Variant) As Variavel
    On Error Resume Next 'O metodo .Item gera um erro caso o item nao exista
    Set Item = m_ColVariavels.Item(VarIndex)
End Function
' Retorna o numero de itens
Property Get Count() As Long
    Count = m_ColVariavels.Count
End Property
'Limpa a Colletion
Sub Clear()
    Set m_ColVariavels = New Collection
End Sub
' Esta fun��o habilita o suporte a enumera��o dos items atraves de um "For Each". Deve haver um DispID-4.
Function NewEnum() As IUnknown
    Set NewEnum = m_ColVariavels.[_NewEnum]
End Function
'M�todo que verifica se um item existe na collection
Function Exists(vIndex As Variant) As Boolean
    Dim vItem As Object
    On Error Resume Next
    Set vItem = m_ColVariavels.Item(vIndex)
    Exists = Not vItem Is Nothing
End Function


