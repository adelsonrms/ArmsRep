VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConexaoDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

Private m_ConexaoBanco          As Object
Private m_ProceduresModelo      As Object ' Scripting.Dictionary
Private m_StringConexao         As String
Private m_objCodeDb_Instance    As Object ' database

'********************PROPRIEDADES********************
Public Property Get ConexaoBanco() As Object
    Set ConexaoBanco = m_ConexaoBanco
End Property
Public Property Set ConexaoBanco(ByRef valor As Object)
    Set m_ConexaoBanco = valor
End Property

Public Property Get ProceduresModelo() As Object ' Scripting.Dictionary
    Set ProceduresModelo = m_ProceduresModelo
End Property
Public Property Set ProceduresModelo(ByRef valor As Object)
    Set m_ProceduresModelo = valor
End Property

Public Property Get StringConexao() As String
    StringConexao = m_StringConexao
End Property
Public Property Let StringConexao(ByVal valor As String)
    m_StringConexao = valor
End Property

'********************ROTINAS DE INICIALIZACAO********************
Private Sub Class_Initialize()
    Set ConexaoBanco = VBA.CreateObject("ADODB.Connection")
    Set m_objCodeDb_Instance = Access.Application.CodeDb
    Set ProceduresModelo = VBA.CreateObject("Scripting.Dictionary")
    Call CarregarProceduresModelo
End Sub

Private Sub Class_Terminate()
    Call Publicas.RemoverObjetosMemoria(ConexaoBanco)
    Set m_objCodeDb_Instance = Nothing
End Sub

'********************SUB-ROTINAS********************
'---------------------------------------------------------------------------------------
' Modulo    : ConexaoDB / Módulo de classe
' Rotina    : Conecta() / Sub
' Autor     : Jefferson
' Data      : 07/11/2012 - 15:15
' Proposta  : Conecta ao banco de dados de acordo com a conection string
'---------------------------------------------------------------------------------------
'ARMS - 14/10/2015  Incluído condição para criar a conexão somente em casos que a mesma ja nao estiver conecatada.
'                   Aproveitar uma conexão existente atraves da variável 'm_ConexaoBanco'
'---------------------------------------------------------------------------------------
Private Sub Conectar()
    On Error GoTo TratarErro
    If Not m_ConexaoBanco Is Nothing Then
        If m_ConexaoBanco.State = 1 Then
            Exit Sub
        End If
        GoTo NovaConexao
    End If
NovaConexao:
    Set m_ConexaoBanco = VBA.CreateObject("ADODB.Connection")
    If Not m_StringConexao = VBA.vbNullString Then
        Call m_ConexaoBanco.Open(StringConexao)
    Else
        Set m_ConexaoBanco = Access.Application.CurrentProject.Connection
    End If
    Exit Sub
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "ConexaoDB.Conecta()", Erl)
    Exit Sub
Resume
End Sub

Public Sub DesConectar()
On Error GoTo TratarErro
    If Not ConexaoBanco Is Nothing Then
        If ConexaoBanco.State = 1 Then
            Call ConexaoBanco.Close
        End If
        Call Publicas.RemoverObjetosMemoria(ConexaoBanco)
    End If
Exit Sub
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "ConexaoDB.DesConecta()", Erl)
End Sub

'---------------------------------------------------------------------------------------
' Modulo....: ConexaoDB / Módulo de classe
' Rotina....: CarregarProceduresModelo / Sub
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 27/08/2013
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina para carregar todas as rotinas modelos no Dicionario para facilitar
'             e principalmente melhorar a performance do uso do ADOX em sua utilização
'---------------------------------------------------------------------------------------
Public Sub CarregarProceduresModelo()
On Error GoTo TratarErro
Dim mCat                As Object 'ADOX.Catalog
Dim cmdComando          As Object 'ADODB.Command
Dim adoxProc            As Object 'ADOX.Procedure
Dim adoxView            As Object 'ADOX.View
Dim NomeProcedure       As String
    
    '---------------------------------------------------------------------------------------------------------
    'ARMS 02/11/2015 - Observado alguns casos de performance e ocorrencia crash ao utilizar a biblioteca ADOX
    '---------------------------------------------------------------------------------------------------------
    'Utilizado a coleção QueryDefs devido a problemas com o uso dos objetos .Procedures e .Views.
    'Demora excessiva no carregamento da coleção
    'Não identifica os motivos dos erros
    'QueryDefs nativo retorna a informação necessários (Nome e SQL)
    If m_ProceduresModelo Is Nothing Then
        Set m_ProceduresModelo = VBA.CreateObject("Scripting.Dictionary")
    End If
    For Each adoxProc In CurrentDb.QueryDefs
        Call VerificarEInserirNoDicionario(adoxProc.Name, adoxProc.SQL)
    Next adoxProc
    
    Call Publicas.RemoverObjetosMemoria(adoxProc)
    
    'Finalizar a função aqui
    Exit Sub
    '----------------------------------------------------------------------
    ' ** OBSERVADO
    '----------------------------------------------------------------------
    Call Conectar
    Set mCat = VBA.CreateObject("ADOX.Catalog")
    Set mCat.ActiveConnection = m_ConexaoBanco
    
    For Each adoxProc In mCat.Procedures
        With adoxProc
            NomeProcedure = .Name
            Set cmdComando = .Command
        End With
        Call VerificarEInserirNoDicionario(NomeProcedure, cmdComando.CommandText)
        Call Publicas.RemoverObjetosMemoria(adoxProc, cmdComando)
    Next adoxProc
    
    For Each adoxView In mCat.Views
        With adoxView
            NomeProcedure = .Name
            Set cmdComando = .Command
        End With
        Call VerificarEInserirNoDicionario(NomeProcedure, cmdComando.CommandText)
        Call Publicas.RemoverObjetosMemoria(adoxProc, cmdComando)
    Next adoxView

    Call Me.DesConectar
    Call Publicas.RemoverObjetosMemoria(mCat)
On Error GoTo 0
Exit Sub
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "ConexaoDB.CarregarProceduresModelo", Erl)
    Resume
End Sub

'---------------------------------------------------------------------------------------
' Modulo....: ConexaoDB / Módulo de classe
' Rotina....: VerificarEInserirNoDicionario / Sub
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 27/08/2013
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina de apoio à rotina CarregarProceduresModelo para diminir a quantidade
'             de codigo
'---------------------------------------------------------------------------------------
Private Sub VerificarEInserirNoDicionario(ByVal Procedure As String, ByVal comando As String)
On Error GoTo TratarErro
    If AuxTexto.IsLinhaMatch(VBA.LCase(Procedure), "^modelo_") Then
        If ProceduresModelo Is Nothing Then Set ProceduresModelo = VBA.CreateObject("Scripting.Dictionary")
        If Not ProceduresModelo.Exists(Procedure) Then
            Call ProceduresModelo.Add(Procedure, comando)
        Else
            ProceduresModelo.item(Procedure) = comando
        End If
    End If
On Error GoTo 0
Exit Sub
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "ConexaoDB.VerificarEInserirNoDicionario", Erl)
End Sub

'---------------------------------------------------------------------------------------
' Modulo    : ConexaoDB / Módulo de classe
' Rotina    : AlterarRegistros() / Sub
' Autor     : Jefferson
' Data      : 07/11/2012 - 15:27
' Proposta  : Rotina para alterar os registros atraves de uma proc
'---------------------------------------------------------------------------------------
Public Sub AlterarRegistros(ByVal strProc As String, ParamArray arrParametros() As Variant)
On Error GoTo TratarErro
    arrParametros = AuxArray.Acertar_Array_Parametros(arrParametros)
    Call RodarComando(strProc, arrParametros)
Exit Sub
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "ConexaoDB.AlterarRegistros()", Erl)
End Sub

'---------------------------------------------------------------------------------------
' Modulo    : ConexaoDB / Módulo de classe
' Rotina    : DeletarRegistros() / Sub
' Autor     : Jefferson
' Data      : 07/11/2012 - 16:26
' Proposta  : Rotina para deletar registro atraves de uma proc
'---------------------------------------------------------------------------------------
Public Sub DeletarRegistros(ByVal strProc As String, ParamArray arrParametros() As Variant)
On Error GoTo TratarErro
    arrParametros = AuxArray.Acertar_Array_Parametros(arrParametros)
    Call RodarComando(strProc, arrParametros)
Exit Sub
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "ConexaoDB.DeletarRegistros()", Erl)
End Sub

Public Function PegarSQLCommand(ByVal strQueryOrigem As String, ParamArray arrParametros() As Variant)
    On Error GoTo TratarErro
    Dim mCat As Object           'ADOX.Catalog
    Dim intContador As Integer
    Dim strComando As String

    strComando = CurrentDb.QueryDefs(strQueryOrigem).SQL

    For intContador = 0 To UBound(arrParametros) Step 2
        strComando = VBA.Replace(strComando, arrParametros(intContador), _
                                 arrParametros(intContador + 1))
    Next intContador

    PegarSQLCommand = strComando
    Exit Function
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "ConexaoDB.ModificarConsulta()", Erl)
End Function

'---------------------------------------------------------------------------------------
' Modulo    : ConexaoDB / Módulo de classe
' Rotina    : ModificarConsulta() / Sub
' Autor     : Jefferson
' Data      : 07/11/2012 - 15:48
' Proposta  : Rotina para criar consultas usando ADOX em tempo de execução a partir de
'             uma consulta modelo. Essa rotina usa um conceito de carregamento dos modelos
'             em um dicionario para melhorar a performance do ADOX
'---------------------------------------------------------------------------------------
Public Sub ModificarConsulta(ByVal strQueryModelo As String, ParamArray arrParametros() As Variant)
On Error GoTo TratarErro
Dim mCat            As Object 'ADOX.Catalog
Dim cmdComando      As Object 'ADODB.Command
Dim intContador     As Integer
Dim strComando      As String

    Call Conectar
    If Not ProceduresModelo.Exists(strQueryModelo) Then
        Set mCat = VBA.CreateObject("ADOX.Catalog")
        Set cmdComando = VBA.CreateObject("ADODB.Command")

        Set mCat.ActiveConnection = ConexaoBanco
        Set cmdComando = mCat.Procedures(strQueryModelo).Command
        strComando = cmdComando.CommandText
        Call VerificarEInserirNoDicionario(strQueryModelo, strComando)
    Else
        strComando = ProceduresModelo.item(strQueryModelo)
    End If
    For intContador = 0 To UBound(arrParametros) Step 2
        strComando = VBA.Replace(strComando, arrParametros(intContador), _
                                 arrParametros(intContador + 1))
    Next intContador

    
    Call Me.RemoveERecriaConsulta(Me.PegarQueryNome(strQueryModelo), strComando)
    Call DesConectar
    Call Publicas.RemoverObjetosMemoria(cmdComando, mCat)
    
Exit Sub
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "ConexaoDB.ModificarConsulta()", Erl)
End Sub

'---------------------------------------------------------------------------------------
' Modulo....: ConexaoDB / Módulo de classe
' Rotina....: RemoveERecriaConsulta / Sub
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 08/08/2013
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina para remover uma consulta pelo nome passado, usando DAO.
'---------------------------------------------------------------------------------------
Public Sub RemoveERecriaConsulta(ByVal nomeConsulta As String, ByVal comando As String)
On Error GoTo TrataErro
Dim dbCorrente      As Object

    If StringConexao = VBA.vbNullString Then
        Set dbCorrente = Access.CodeDb
    Else
        Set dbCorrente = Access.DBEngine.OpenDatabase(StringConexao)
    End If
    
    If Me.ObjetoExiste(dbCorrente, Access.AcObjectType.acQuery, nomeConsulta) Then
        Call Access.DoCmd.DeleteObject(Access.AcObjectType.acQuery, nomeConsulta)
    End If
    Call dbCorrente.CreateQueryDef(nomeConsulta, comando)
    Call Publicas.RemoverObjetosMemoria(dbCorrente)
Exit Sub
TrataErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxTabelas.RemoveERecriaConsulta()", Erl)
End Sub

'---------------------------------------------------------------------------------------
' Modulo....: ConexaoDB / Módulo de classe
' Rotina....: PegaComandoPorQueryModelo / Function
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 08/08/2013
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina para pegar o conteudo de uma Query usando ADOX
'---------------------------------------------------------------------------------------
Public Function PegaComandoPorQueryModelo(ByVal QueryModelo As String) As String
On Error GoTo TratarErro
Dim mCat            As Object 'ADOX.Catalog
Dim cmdComando      As Object 'ADODB.Command
Dim comando         As String

    Call Conectar

    Set mCat = VBA.CreateObject("ADOX.Catalog")
    Set cmdComando = VBA.CreateObject("ADODB.Command")

    Set mCat.ActiveConnection = ConexaoBanco
    Set cmdComando = mCat.Procedures(QueryModelo).Command

    comando = cmdComando.CommandText
    PegaComandoPorQueryModelo = comando
Exit Function
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "ConexaoDB.PegaComandoPorQueryModelo()", Erl)
End Function

'---------------------------------------------------------------------------------------
' Modulo....: ConexaoDB / Módulo de classe
' Rotina....: RemoveQuery() / Sub
' Autor.....: Jefferson
' Contato...: jefferson@tecnun.com.br
' Data......: 12/12/2012
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina que verifica se uma query é uma procedure ou uma view e deleta ela
'             Não tem tratamento de erro pq esperasse que um erro possa ocorrer
'---------------------------------------------------------------------------------------
Public Sub RemoveQuery(ByVal Query As String, Optional ByVal Catalogo As Object)
On Error Resume Next
Dim QAux        As String
Dim mCat        As Object

    Set mCat = VBA.CreateObject("ADOX.Catalog")
    If Not Catalogo Is Nothing Then
        Set mCat = Catalogo
    Else
        mCat.ActiveConnection = ConexaoBanco
    End If
    QAux = mCat.Procedures.item(Query).Name
    If VBA.Err.Number = 0 Then
        Call mCat.Procedures.Delete(Query)
        GoTo Fim
    End If
    VBA.Err.Clear
    QAux = mCat.Views.item(Query).Name
    If VBA.Err.Number = 0 Then Call mCat.Views.Delete(Query)
Fim:
    VBA.Err.Clear
End Sub

'---------------------------------------------------------------------------------------
' Modulo....: ConexaoDB / Módulo de classe
' Rotina....: DeletaTabela() / Sub
' Autor.....: Jefferson
' Contato...: jefferson@tecnun.com.br
' Data......: 12/12/2012
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina para excluir uma tabela usando ADOX
'---------------------------------------------------------------------------------------
Public Sub RemoveTabela(ByVal Tabela As String, Optional ByVal Catalogo As Object)
On Error Resume Next
Dim TabAux      As String
Dim mCat        As Object

    Set mCat = VBA.CreateObject("ADOX.Catalog")

    If Not Catalogo Is Nothing Then
        Set mCat = Catalogo
    Else
        If Not Conectado Then Conectar
        mCat.ActiveConnection = ConexaoBanco
    End If
    TabAux = mCat.Tables.item(Tabela).Name
    If VBA.Err.Number = 0 Then
        Call mCat.Tables.Delete(Tabela)
        'mCat.Tables.Refresh
    End If
    VBA.Err.Clear
End Sub

'********************FUNCOES********************
'---------------------------------------------------------------------------------------
' Modulo    : ConexaoDB / Módulo de classe
' Rotina    : Conectado() / Function
' Autor     : Jefferson
' Data      : 07/11/2012 - 16:25
' Proposta  : Verifica se a conexao está ativa ou não
'---------------------------------------------------------------------------------------
Private Function Conectado() As Boolean
On Error GoTo TratarErro
Dim blResultado     As Boolean

    blResultado = False
    If Not ConexaoBanco Is Nothing Then
        If ConexaoBanco.State = 1 Then 'ADODB.ObjectStateEnum.adStateOpen
            blResultado = True
        End If
    End If
    Conectado = blResultado
Exit Function
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "ConexaoDB.IsConectado()", Erl)
End Function

'---------------------------------------------------------------------------------------
' Modulo    : ConexaoDB / Módulo de classe
' Rotina    : PegaComando() / Function
' Autor     : Jefferson
' Data      : 07/11/2012 - 16:01
' Proposta  : Rotina para criar um comando para ser usado posteriormente
'---------------------------------------------------------------------------------------

'ByVal TipoComando As ADODB.CommandTypeEnum
Private Function PegaComando(ByVal strProc As String, ByVal TipoComando As Integer, _
                             ByVal arrParametros As Variant) As Object 'ADODB.Command
On Error GoTo TratarErro
Dim adoComando      As Object 'New ADODB.Command
Dim adoParametro    As Object 'New ADODB.Parameter
Dim btContador      As Byte

    Set adoComando = VBA.CreateObject("ADODB.Command")

    With adoComando
        .ActiveConnection = ConexaoBanco
        .CommandType = TipoComando
        .CommandText = strProc

        If UBound(arrParametros) >= 0 Then
            If UBound(arrParametros) = 0 Then
                If VBA.IsMissing(arrParametros(0)) Then GoTo Fim
            End If
            For btContador = 0 To UBound(arrParametros) Step 1
                Set adoParametro = VBA.CreateObject("ADODB.Parameter")
                With adoParametro
                    .value = VBA.IIf(VBA.IsEmpty(arrParametros(btContador)), Null, arrParametros(btContador))
                    .Type = 200 'ADODB.DataTypeEnum.adVarChar
                    .size = 255
                    .Direction = 1 'ADODB.ParameterDirectionEnum.adParamInput
                End With
                .Parameters.Append adoParametro
                Set adoParametro = Nothing
            Next btContador
        End If
    End With
Fim:
    Set PegaComando = adoComando
Exit Function
TratarErro:
    If Not strProc = "Insere_LogErro" Then 'Se for "Insere_LogErro" então o problema esta na conexao com o banco
        Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "ConexaoDB.PegaComando()", Erl)
        Call VBA.Err.Raise(VBA.Err.Number, VBA.Err.Description)
    End If
End Function

'---------------------------------------------------------------------------------------
' Modulo    : ConexaoDB / Módulo de classe
' Rotina    : RodarComando() / Function
' Autor     : Jefferson
' Data      : 07/11/2012 - 16:25
' Proposta  : Rotina para rodar o comando depois de setado
'---------------------------------------------------------------------------------------
'ByVal Cursor As ADODB.CursorTypeEnum = ADODB.CursorTypeEnum.adOpenKeyset,
'ByVal LockType As ADODB.LockTypeEnum = ADODB.LockTypeEnum.adLockUnspecified,
'ByVal TipoComando As ADODB.CommandTypeEnum = ADODB.CommandTypeEnum.adCmdStoredProc) As ADODB.Recordset

Private Function RodarComando(ByVal strProc As String, ByVal arrParametros As Variant, _
                              Optional ByVal Cursor As Integer = 1, _
                              Optional ByVal LockType As Integer = -1, _
                              Optional ByVal Desconectado As Boolean, _
                              Optional ByVal TipoComando As Integer = 4) As Object 'ADODB.Recordset
On Error GoTo TratarErro
Dim adoRS       As Object 'ADODB.Recordset
Dim adoComando  As Object 'New ADODB.Command
Dim HoraInicio  As Date, TempoInicial As Double
Dim HoraFinal   As Date, TempoFinal As Double

    If Not Conectado() Then Conectar
    'Set adoComando = VBA.CreateObject("ADODB.Command")
    Set adoComando = PegaComando(strProc, TipoComando, arrParametros)

    If Not adoComando Is Nothing Then
        Set adoRS = VBA.CreateObject("ADODB.Recordset")
        If Desconectado Then
            adoRS.CursorLocation = 3 'ADODB.CursorLocationEnum.adUseClient
        End If
        
        HoraInicio = VBA.Now(): TempoInicial = VBA.Timer()
        Call adoRS.Open(adoComando, , Cursor, LockType)
        HoraFinal = VBA.Now():  TempoFinal = VBA.Timer()
        
        If Desconectado Then
            Set adoRS.ActiveConnection = Nothing
            Call DesConectar
        End If
    End If

    Set RodarComando = adoRS
Fim:
    Call Publicas.RemoverObjetosMemoria(adoComando, adoRS)

Exit Function
TratarErro:
    If Not strProc = "Insere_LogErro" Then    'Se for "Insere_LogErro" então o problema esta na conexao com o banco
        If bStop = True Then
            Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "ConexaoDB.RodarComando()", Erl)
        Else
            errNumber = VBA.Err.Number
            errDescricao = VBA.Err.Description
        End If

        Call VBA.Err.Raise(VBA.Err.Number, VBA.Err.Description)
        Resume
    End If
    Call DesConectar
    GoTo Fim
End Function

'---------------------------------------------------------------------------------------
' Modulo    : ConexaoDB / Módulo de classe
' Rotina    : InserirRegistros() / Function
' Autor     : Jefferson
' Data      : 07/11/2012 - 16:23
' Proposta  : Rotina para inserir registros em uma tabela atraves de uma proc, retornando
'             o id do registro inserido
'---------------------------------------------------------------------------------------
Public Function InserirRegistros(ByVal strProc As String, ParamArray arrParametros() As Variant) As Long
On Error GoTo TratarErro
Dim adoRS       As Object 'ADODB.Recordset
Dim lngID       As Long

    Set adoRS = VBA.CreateObject("ADODB.Recordset")

    arrParametros = AuxArray.Acertar_Array_Parametros(arrParametros)
    Call RodarComando(strProc, arrParametros)
    Set adoRS = ConexaoBanco.Execute("Select @@IDENTITY AS NovoID")
    If Not adoRS.BOF And Not adoRS.EOF Then
        lngID = VBA.CLng(adoRS.Fields(0).value)
    End If
    InserirRegistros = lngID
    Call Publicas.RemoverObjetosMemoria(adoRS)
Exit Function
TratarErro:
    If bStop = False Then
        errNumber = VBA.Err.Number
        errDescricao = VBA.Err.Description
        Exit Function
    End If
    If Not strProc = "Insere_LogErro" Then    'Se for "Insere_LogErro" então o problema esta na conexao com o banco
        Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "ConexaoDB.InserirRegistros()", Erl)
    End If
    Resume

End Function

'---------------------------------------------------------------------------------------
' Modulo    : ConexaoDB / Módulo de classe
' Rotina    : PegarRS() / Function
' Autor     : Jefferson
' Data      : 07/11/2012 - 16:22
' Proposta  : Rotina que atraves de uma proc retorna um recordset que pode ser usado em
'             outras funcoes
'---------------------------------------------------------------------------------------
Public Function PegarRS(ByVal strProc As String, ParamArray arrParametros() As Variant) As Object 'ADODB.Recordset
On Error GoTo TratarErro
Dim rsAux       As Object 'ADODB.Recordset

    Set rsAux = VBA.CreateObject("ADODB.Recordset")
    If AuxTexto.IsLinhaMatch(VBA.UCase(strProc), "^SELECT") Then
        Set rsAux = Me.ExecutarQuery(strProc)
    Else
        arrParametros = AuxArray.Acertar_Array_Parametros(arrParametros)
        Set rsAux = RodarComando(strProc, arrParametros)
    End If
    Set PegarRS = rsAux
Exit Function
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "ConexaoDB.PegarRS()", Erl)
End Function

'---------------------------------------------------------------------------------------
' Modulo    : ConexaoDB / Módulo de classe
' Rotina    : PegarArray() / Function
' Autor     : Jefferson
' Data      : 07/11/2012 - 15:57
' Proposta  : Rotina para criar um array com base em um recordset
'---------------------------------------------------------------------------------------
Public Function PegarArray(ByVal Objeto As Variant, ParamArray Parametros() As Variant) As Variant
On Error GoTo TratarErro
Dim arrResultado        As Variant
Dim rs                  As Object 'ADODB.Recordset

    Set rs = VBA.CreateObject("ADODB.Recordset")
    If VBA.TypeName(Objeto) = "String" Then
        Parametros = AuxArray.Acertar_Array_Parametros(Parametros)
        'ADODB.CursorTypeEnum.adOpenForwardOnly
        'ADODB.LockTypeEnum.adLockReadOnly
        Set rs = RodarComando(Objeto, Parametros, 0, 1, False)
    Else
        Set rs = Objeto
    End If

    If Not rs Is Nothing Then
        With rs
            If Not .BOF And Not .EOF Then
                arrResultado = .GetRows(VBA.CLng(.RecordCount))

            Else

                arrResultado = Empty

            End If
        End With
    Else
        arrResultado = Empty
    End If
    PegarArray = arrResultado

    Call Publicas.RemoverObjetosMemoria(rs)

    Exit Function
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "ConexaoDB.PegarArray()", Erl)
End Function

'---------------------------------------------------------------------------------------
' Modulo    : ConexaoDB / Módulo de classe
' Rotina    : PegarDicionario() / Function
' Autor     : Jefferson
' Data      : 07/11/2012 - 15:57
' Proposta  : Rotina para criar um dicionario baseado em um recordset
'---------------------------------------------------------------------------------------
Public Function PegarDicionario(ByVal Objeto As Variant, ParamArray Parametros() As Variant) As Object 'Scripting.Dictionary
On Error GoTo TratarErro
Dim dicResultado            As Object ' Scripting.Dictionary
Dim arrAux                  As Variant
Dim lngContador             As Long
Dim intContador             As Integer
Dim strItem                 As String
    
    Set dicResultado = VBA.CreateObject("Scripting.Dictionary")
    
    Parametros = AuxArray.Acertar_Array_Parametros(Parametros)
    arrAux = PegarArray(Objeto, Parametros)
    If VBA.IsArray(arrAux) Then
        For lngContador = 0 To UBound(arrAux, 2) Step 1
        
            arrAux(0, lngContador) = RemoverQuebrasDeLinha(arrAux(0, lngContador))
            
            If Not dicResultado.Exists(arrAux(0, lngContador)) Then
                strItem = Empty
                If UBound(arrAux, 1) > 0 Then
                    For intContador = 1 To UBound(arrAux, 1) Step 1
                        strItem = strItem & arrAux(intContador, lngContador) & "|"
                    Next intContador
                    strItem = VBA.Left(strItem, VBA.Len(strItem) - 1)
                End If
                Call dicResultado.Add(arrAux(0, lngContador), strItem)
            End If
        Next lngContador
        
    End If
    Set PegarDicionario = dicResultado
Exit Function
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "ConexaoDB.PegarDicionario()", Erl)
End Function

'---------------------------------------------------------------------------------------
' Modulo....: ConexaoDB / Módulo de classe
' Rotina....: PegarString() / Function
' Autor.....: Jefferson
' Contato...: jefferson@tecnun.com.br
' Data......: 08/11/2012 - 11:21
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina para criar uma string, separada por um delimitador, tendo como base
'             um Recordset
'---------------------------------------------------------------------------------------
Public Function PegarString(ByVal delimitador As String, ByVal Objeto As Variant, _
                           ParamArray Parametros() As Variant) As String
On Error GoTo TratarErro
Dim arrAux              As Variant
Dim lngContador         As Long
Dim strResultado        As String

    strResultado = VBA.vbNullString

    arrAux = PegarArray(Objeto, Parametros)
    If VBA.IsArray(arrAux) Then
        For lngContador = 0 To UBound(arrAux, 2) Step 1
            strResultado = strResultado & arrAux(0, lngContador) & delimitador
        Next lngContador
        strResultado = VBA.Left(strResultado, VBA.Len(strResultado) - 1)
    End If

    PegarString = strResultado
Exit Function
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "ConexaoDB.PegarString()", Erl)
End Function

'---------------------------------------------------------------------------------------
' Modulo....: ConexaoDB / Módulo de classe
' Rotina....: ObjetoExiste / Function
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 08/08/2013
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina que verifica se um determinado objeto existe no banco de dados informado
'---------------------------------------------------------------------------------------
Public Function ObjetoExiste(ByRef db As Object, ByVal acObj As Access.AcObjectType, _
                             ByVal NomeObjeto As String) As Boolean
On Error Resume Next
Dim strObjEncontrado    As String
Dim blResultado         As Boolean
Dim ctn                 As Object ' dao.Container

    Select Case acObj
        Case Access.AcObjectType.acTable
            strObjEncontrado = db.TableDefs(NomeObjeto).Name
            If Not VBA.Err.Number = 0 Then blResultado = False
        Case Access.AcObjectType.acQuery
            strObjEncontrado = db.QueryDefs(NomeObjeto).Name
            If Not VBA.Err.Number = 0 Then blResultado = False
        Case Else
            Set ctn = db.Containers(acObj)
            strObjEncontrado = ctn.Documents(NomeObjeto).Name
            If Not VBA.Err.Number = 0 Then blResultado = False
    End Select
    If NomeObjeto = strObjEncontrado Then
        blResultado = True
    Else
        blResultado = False
    End If
    ObjetoExiste = blResultado

Exit Function
TrataErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "ConexaoDB.ObjetoExiste()", Erl)
End Function

'Função encapsulada para identifdicar se a tabela existe
Function TabelaExiste(pNomeTabela As String) As Boolean
   TabelaExiste = Me.ObjetoExiste(m_objCodeDb_Instance, acTable, pNomeTabela) And pNomeTabela <> ""
End Function

'Função encapsulada para identifdicar se a consulta
Function ConsultaExiste(pConsulta As String) As Boolean
   ConsultaExiste = Me.ObjetoExiste(m_objCodeDb_Instance, acQuery, pConsulta) And pConsulta <> ""
End Function

'---------------------------------------------------------------------------------------
' Modulo....: ConexaoDB / Módulo de classe
' Rotina....: RecordsetValido() / Function
' Autor.....: Jefferson
' Contato...: jefferson@tecnun.com.br
' Data......: 23/12/2012
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: função para verificar se um rs é valido ou nao
'---------------------------------------------------------------------------------------
Public Function RecordsetValido(ByRef rs As Object) As Boolean 'ADODB.Recordset
On Error GoTo TratarErro
Dim Resultado       As Boolean

    If Not rs Is Nothing Then
        With rs
            If Not .BOF And Not .EOF Then
                Resultado = True
            Else
                Resultado = False
            End If
        End With
    Else
        Resultado = False
    End If
    RecordsetValido = Resultado
Exit Function
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "ConexaoDB.RecordsetValido()", Erl)
End Function

'---------------------------------------------------------------------------------------
' Modulo    : ConexaoDB / Módulo de classe
' Rotina    : ExecutarQuery() / Function
' Autor     : Douglas
' Data      : 14/01/2013 - 18:31
' Proposta  : Rotina que atraves de uma query retorna um recordset que pode ser usado em
'             outras funcoes
'---------------------------------------------------------------------------------------

Public Function ExecutarQuery(ByVal strSQL As String) As Object 'ADODB.Recordset
On Error GoTo TratarErro
Dim arrAux(0 To 0)          As Variant
    Set ExecutarQuery = RodarComando(strSQL, arrAux, TipoComando:=1)   'ADODB.CommandTypeEnum.adCmdText)
Exit Function
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "ConexaoDB.ExecutarQuery()", Erl)
End Function

'---------------------------------------------------------------------------------------
' Modulo    : ConexaoDB / Módulo de classe
' Rotina    : PegarTabelaBanco() / Function
' Autor     : Jefferson
' Data      : 07/11/2012 - 16:22
' Proposta  : Rotina que atraves de uma proc retorna um recordset que pode ser usado em
'             outras funcoes
'---------------------------------------------------------------------------------------
Public Function PegarTabelaBanco(ByVal strNome As String) As Object 'ADODB.Recordset
On Error GoTo TratarErro
Dim tblAux      As Object
    'LockType:=ADODB.LockTypeEnum.adLockOptimistic
    'TipoComando:=ADODB.CommandTypeEnum.adCmdTable
    Set tblAux = VBA.CreateObject("ADODB.Recordset")
    Set tblAux = RodarComando(strNome, Array(), LockType:=3, TipoComando:=2)
    Set PegarTabelaBanco = tblAux
Exit Function
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "ConexaoDB.PegarTabelaBanco()", Erl)
End Function

'---------------------------------------------------------------------------------------
' Modulo....: ConexaoDB / Módulo de classe
' Rotina....: PegarQueryNome / Function
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 15/08/2013
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina para gerar um nome único para a consulta para ser usada em
'             conjunto com o ADOX
'---------------------------------------------------------------------------------------
Public Function PegarQueryNome(ByVal Query As String) As String
On Error GoTo TrataErro
Dim QueryAux            As String
    If Not VBA.Left(Query, 2) = "M_" Then
        If VBA.InStrRev(Query, "Modelo_") > 0 Then
            QueryAux = "M_" & VBA.Replace(Query, "Modelo_", "") & "_" & ChaveUsuario
        Else
            QueryAux = "M_" & Query & "_" & ChaveUsuario
        End If
    End If
    PegarQueryNome = QueryAux
Exit Function
TrataErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "ConexaoDB.PegarQuery()", Erl)
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : ConexaoDB.ExecutarDDL()
' TIPO             : Function
' DATA/HORA        : 29/04/2014 14:32
' CONSULTOR        : (ADELSON)/ TECNUN - Adelson Rosendo Marques da Silva
' CONTATO          : adelson@tecnun.com.br
' DESCRIÇÃO        : Cria um objeto no banco de dados usando uma instrução DDL
'---------------------------------------------------------------------------------------
Function ExecutarDDL(pSQL As String, Optional db As Object) As Boolean
    If db Is Nothing Then Set db = Access.Application.CodeDb
    On Error Resume Next
    Call db.Execute(pSQL)
    ExecutarDDL = VBA.Err.Number = 0
End Function

Public Property Get CodeDb_Instance() As Object
    Set CodeDb_Instance = m_objCodeDb_Instance
End Property

Public Property Set CodeDb_Instance(objCodeDb_Instance As Object)
    Set m_objCodeDb_Instance = objCodeDb_Instance
End Property

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : ConexaoDB.VerificaConsistenciaLink()
' TIPO             : Function
' DATA/HORA        : 11/06/2015 15:37
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Verifica a integridade de um vinculo com uma tabela
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      11/06/2015 15:37    Adelson         Criação/Atualização do procedimento
'---------------------------------------------------------------------------------------
Function VerificaConsistenciaLink(strTabela As String, Optional bForcarRefresh As Boolean = False, Optional contexto As String) As Variant
    Dim sArquivoOrigem As String
    Dim rsInfo As Object, vRetorno As Variant
    Dim bArquivo As Boolean, bLink As Boolean
    Dim sMensagem As String, codRetorno As Boolean

10  On Error GoTo VerificaConsistenciaLink_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
20  Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Function ConexaoDB.VerificaConsistenciaLink()"
    '----------------------------------------------------------------------------------------------------

30  Set rsInfo = Me.PegarRS("Pegar_InfoTabelaVinculada", strTabela)
40  If rsInfo.EOF Then
50      sMensagem = "Tabela '" & strTabela & "' não é uma tabela vinculada"
60      codRetorno = 0
70  Else
80      If Not rsInfo.EOF Then sArquivoOrigem = rsInfo!DBOrigem.value
90      rsInfo.Close
        'Consiste existencia do arquivo
100     bArquivo = FileExists(sArquivoOrigem)
110     codRetorno = bArquivo
        If Not bArquivo Then sMensagem = "Arquivo '" & sArquivoOrigem & "' de origem não existe"
        
120     If bForcarRefresh And bArquivo Then
            'Consiste se o Link esta ativo
130         Set rsInfo = m_objCodeDb_Instance.TableDefs(strTabela)
140         On Error Resume Next
150         rsInfo.RefreshLink
160         If VBA.Err = 0 Then
170             sMensagem = "Arquivo Validado. Vinculo Consiste"
180         Else
190             sMensagem = "Arquivo Validado. Erro na atualização do vinculo (" & VBA.Err.Description & ")"
200         End If
210         codRetorno = bArquivo And VBA.Err = 0
220     End If
230
240 End If
250 Call Publicas.RemoverObjetosMemoria(rsInfo)
260 VerificaConsistenciaLink = VBA.Array(codRetorno, sMensagem)
    Call auxAuditoria.RegistraLog(cstr_ProcedureName & " > " & contexto & " > " & sMensagem)
Fim:
270 On Error GoTo 0
280 Exit Function

VerificaConsistenciaLink_Error:
290 If VBA.Err <> 0 Then
300     lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
310     Call Excecoes.TratarErro(strErrorMessagem, lngErrorNumber, cstr_ProcedureName, VBA.Erl)
320 End If
    GoTo Fim:
    'Debug Mode
330 Resume
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : ConexaoDB.PegarComandoSQLModelo()
' TIPO             : Function
' DATA/HORA        : 18/06/2015 11:38
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Obtem o comando SQL de uma consulta Modelo.
'                    Mesma logica da ModificarConsulta(), porem sem criar a consulta de destino
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      18/06/2015 11:38    Adelson         Criação/Atualização do procedimento
'---------------------------------------------------------------------------------------
Public Function PegarComandoSQLModelo(ByVal strQueryModelo As String, objConexao As Object, ParamArray arrParametros() As Variant) As String
10  On Error GoTo PegarComandoSQLModelo_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
20  Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Function ConexaoDB.PegarComandoSQLModelo()"
    '----------------------------------------------------------------------------------------------------

30  On Error GoTo TratarErro
    Dim mCat As Object           'ADOX.Catalog
    Dim cmdComando As Object     'ADODB.Command
    Dim intContador As Integer
    Dim strComando As String
    
50  If Not ProceduresModelo.Exists(strQueryModelo) Then
'''60      Set mCat = VBA.CreateObject("ADOX.Catalog")
'''70      Set cmdComando = VBA.CreateObject("ADODB.Command")
'''80      Set mCat.ActiveConnection = objConexao  'Conexao ADO com o DB informado
'''90      Set cmdComando = mCat.Procedures(strQueryModelo).Command
'''100     strComando = cmdComando.CommandText
        strComando = CurrentDb.QueryDefs(strQueryModelo).SQL
110     Call VerificarEInserirNoDicionario(strQueryModelo, strComando)
120 Else
130     strComando = ProceduresModelo.item(strQueryModelo)
140 End If

150 For intContador = 0 To UBound(arrParametros) Step 2
160     strComando = VBA.Replace(strComando, arrParametros(intContador), _
                                 arrParametros(intContador + 1))
170 Next intContador

180 PegarComandoSQLModelo = strComando
190 Exit Function
TratarErro:
200 Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "ConexaoDB.PegarComandoSQLModelo()", Erl)

Fim:
210 On Error GoTo 0
220 Exit Function

PegarComandoSQLModelo_Error:
230 If VBA.Err <> 0 Then
240     lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
250     Call Excecoes.TratarErro(strErrorMessagem, lngErrorNumber, cstr_ProcedureName, VBA.Erl)
260 End If
    GoTo Fim:
    'Debug Mode
270 Resume
End Function
'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : ConexaoDB.CriarConsultaPassagem()
' TIPO             : Function
' DATA/HORA        : 18/06/2015 14:13
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Criar uma consulta passagem no banco especificado
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      18/06/2015 14:13    Adelson         Criação/Atualização do procedimento
'---------------------------------------------------------------------------------------
Function CriarConsultaPassagem(SQLComando As String, _
                               ConnectString As String, _
                               Optional db As Object, _
                               Optional nomeConsulta As String) As Boolean
    Dim qrdPT As Object ' QueryDef
10  On Error GoTo CriarConsultaPassagem_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
20  Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Function ConexaoDB.CriarConsultaPassagem()"
    '----------------------------------------------------------------------------------------------------

    'Define a instancia do banco no qual será criado
30  If db Is Nothing Then Set db = Access.Application.CurrentDb

40  If nomeConsulta <> "" Then
50      If ObjetoExiste(db, acQuery, nomeConsulta) Then Call ExecutarDDL("DROP TABLE " & nomeConsulta, db)
60  End If

    'Cria o objeto
70  Set qrdPT = db.CreateQueryDef()
80  With qrdPT
90      .Name = nomeConsulta
        'Define String de conexao (String em branco, a consulta não será Pass Through
100     .Connect = ConnectString
110     .ReturnsRecords = True
120     .ODBCTimeout = 0
130     .SQL = SQLComando
140 End With
150 If nomeConsulta <> "" Then Call db.QueryDefs.Append(qrdPT)
    'Devolve a instancia da consulta
160 CriarConsultaPassagem = ObjetoExiste(db, acQuery, nomeConsulta)

Fim:
170 On Error GoTo 0
180 Exit Function

CriarConsultaPassagem_Error:
190 If VBA.Err <> 0 Then
200     lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
210     Call Excecoes.TratarErro(strErrorMessagem, lngErrorNumber, cstr_ProcedureName, VBA.Erl)
220 End If
    GoTo Fim:
    'Debug Mode
230 Resume
End Function

