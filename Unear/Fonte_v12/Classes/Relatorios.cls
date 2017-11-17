VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Relatorios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'********************PROPRIEDADES********************
Private m_Rels          As VBA.Collection
Private m_CDB_Apoio     As ConexaoDB

Private Property Get Rels() As Collection
    Set Rels = m_Rels
End Property
Private Property Set Rels(ByRef valor As VBA.Collection)
    Set m_Rels = valor
End Property

Public Property Get CDB_Apoio() As ConexaoDB
    Set CDB_Apoio = m_CDB_Apoio
End Property
Public Property Set CDB_Apoio(ByRef valor As ConexaoDB)
    Set m_CDB_Apoio = valor
End Property
'''Public Property Get StringConexao() As String
'''    StringConexao = m_StringConexao
'''End Property
'''Public Property Let StringConexao(ByVal valor As String)
'''    m_StringConexao = valor
'''End Property


Public Property Get count() As Long
    count = Rels.count
End Property

'********************INICIALIZACAO********************
Private Sub Class_Initialize()
    Set Rels = New VBA.Collection
    Set CDB_Apoio = New ConexaoDB
    Call PreencheRelatorio
End Sub
Private Sub Class_Terminate()
    Call Publicas.RemoverObjetosMemoria(CDB_Apoio, Rels)
End Sub

'********************SUB-ROTINAS********************

Public Sub PreencheRelatorio()
On Error GoTo TrataErro
Dim RelDetalhe          As RelatorioDetalhe
Dim rsRelatorio         As Object 'ADODB.Recordset
Dim Posicao             As Integer
Dim contador            As Integer
Dim chave               As Variant
Dim RetornoValidacao    As Variant
    
    'Incluido para validade a consistencia de tabelas vinculadas envolvidas no processo
    'A tabela 'tblRelatorios_Criados' precisa existir nesse processo
    RetornoValidacao = CDB_Apoio.VerificaConsistenciaLink("tblRelatorios_Criados", False, "CDB_Apoio.PreencheRelatorio()")
    
    If Not RetornoValidacao(0) = True Then
'        Call auxBackEnd.MostrarAvisoInconsistenciaLinks(False, CStr(RetornoValidacao(1)))
        Exit Sub
    End If
    
    Set rsRelatorio = CDB_Apoio.PegarRS("Pegar_Relatorio")

    If Not rsRelatorio Is Nothing Then

        With rsRelatorio
            .MoveFirst
            Do While Not .EOF
                Posicao = Posicao + 1
                chave = .Fields(1).value
                If Not Me.Exists(chave) Then
                    Set RelDetalhe = New RelatorioDetalhe
                    Call RelDetalhe.SetPropriedades("Posicao", Posicao)
                    For contador = 0 To .Fields.count - 1 Step 1
                        Call RelDetalhe.SetPropriedades(.Fields(contador).Name, .Fields(contador).value)
                    Next contador
                    Call Rels.Add(RelDetalhe, chave)
                End If
                .MoveNext
            Loop
        End With
    End If
    Call Publicas.RemoverObjetosMemoria(rsRelatorio, RelDetalhe)
Exit Sub
TrataErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "Relatorio.PreencheRelatorio()")
    Exit Sub
    Resume
End Sub

Public Sub Add(ByVal Referencia As String, Optional ByVal NomeRelatorio As String, Optional ByVal NomeArquivo As String, _
               Optional ByVal Tabela As String, Optional ByVal ParentID As Integer, Optional ByVal Identificacao As String, _
               Optional ByVal Regex As String, Optional ByVal TipoRelatorio As String, Optional ByVal Visivel As Boolean, _
               Optional ByVal NomeExibicao As String, Optional ByVal Vinculada As Boolean = True)
On Error GoTo TrataErro
Dim RelDetalhe      As New RelatorioDetalhe
Dim RelID           As Long

    RelID = CDB_Apoio.InserirRegistros("Inserir_Relatorio", Referencia, NomeRelatorio, NomeArquivo, Tabela, _
                                       ParentID, Identificacao, Regex, TipoRelatorio, Visivel, NomeExibicao, _
                                       Vinculada)
    With RelDetalhe
        .Posicao = Rels.count + 1
        .ID = RelID
        .Referencia = Referencia
        .NomeRelatorio = NomeRelatorio
        .NomeArquivo = NomeArquivo
        Set .Tabela = New Tabelas
        .Tabela.Name = Tabela
        .ParentID = ParentID
        .Identificacao = Identificacao
        .Regex = Regex
        .TipoRelatorio = TipoRelatorio
        .Visivel = Visivel
        .NomeExibicao = NomeExibicao
        .Vinculada = Vinculada
    End With
    Call Rels.Add(RelDetalhe, Referencia)
Exit Sub
TrataErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "Relatorios.Add()", Erl)
End Sub

Public Sub Clear()
On Error GoTo TrataErro
    Set Rels = New Collection
Exit Sub
TrataErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "Relatorio.Clear()", Erl)
End Sub

'********************FUNCOES********************

Public Function item(ByVal index As Variant) As RelatorioDetalhe
On Error GoTo TrataErro
    If VBA.IsNumeric(index) Then
        Set item = Rels.item(PegaID(index))
    Else
        Set item = Rels.item(index)
    End If
Exit Function
TrataErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "Relatorio.Item()", Erl)
End Function

'---------------------------------------------------------------------------------------
' Modulo....: Relatorio / Módulo de classe
' Rotina....: Exists() / Function
' Autor.....: Jefferson
' Contato...: jefferson@tecnun.com.br
' Data......: 09/11/2012
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Retorna "True" se o valor da chave (NÃO O INDEX) existe na coleção
'---------------------------------------------------------------------------------------
Public Function Exists(ByVal key As String) As Boolean
On Error GoTo TrataErro
Dim Rel     As RelatorioDetalhe

    Set Rel = Rels.item(key)
    Call Publicas.RemoverObjetosMemoria(Rel)
    Exists = True

Exit Function
TrataErro:
    If VBA.Err.Number <> 5 Then
        Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "Relatorio.Exists()", Erl)
    End If
    Exists = False
End Function

'---------------------------------------------------------------------------------------
' Modulo....: AuxTabela / Módulo
' Rotina....: PegaID() / Function
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 21/12/2012 - 11:22
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina para checar o indice do dicionario de relatorios e comparar o
'             ID contra o enumerador
'---------------------------------------------------------------------------------------
Private Function PegaID(ByVal RelID As Variant) As Integer
On Error GoTo TrataErro
Dim iIndex As Integer

    For iIndex = 1 To Rels.count Step 1
        If Rels.item(iIndex).ID = RelID Then
           PegaID = iIndex
           Exit Function
        End If
    Next iIndex
    'So vai rodar essa parte caso não seja encontrado, pois assim vai utilizar a posição
    For iIndex = 1 To Rels.count Step 1
        If Rels.item(iIndex).Posicao = RelID Then
            PegaID = iIndex    'Rels.Item(iIndex).ID
            Exit Function
        End If
    Next iIndex
Exit Function
TrataErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "Relatorios.PegaID()", Erl)
End Function


