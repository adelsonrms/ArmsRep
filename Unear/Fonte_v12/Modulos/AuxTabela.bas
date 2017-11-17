Attribute VB_Name = "AuxTabela"
Attribute VB_Description = "Contem funções relativas as tabelas do banco de dados"
Option Compare Database
Option Explicit

'********************SUB-ROTINAS********************

'---------------------------------------------------------------------------------------
' Modulo....: AuxTabela / Módulo
' Rotina....: Cria_Tabela_Vinculo / Sub
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 14/08/2013
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina para criar a tabela fisicamento no endereço modelo e criar o vinculo
'             no client
'---------------------------------------------------------------------------------------
Private Sub Cria_Tabela_Vinculo(ByRef db As Object, ByVal NomeArquivo As String, _
                                ByVal chave As String, ByVal RelatorioAux As TipoRelatorio)
On Error GoTo TrataErro
Dim tblDef          As Object ' dao.TableDef
Dim contador        As Byte
Dim CaminhoArq      As String
Dim TabelaAux       As String
Dim arrTabelas      As Variant
Dim arrParentID     As Variant

    arrTabelas = Array(Relatorio.item(RelatorioAux).Tabela.Name)
    arrParentID = Array(RelatorioAux)

    CaminhoArq = CriaTabelaDados(NomeArquivo, RelatorioAux)
    For contador = 0 To UBound(arrTabelas) Step 1
        TabelaAux = arrTabelas(contador) & chave
        If Conexao.ObjetoExiste(db, Access.AcObjectType.acTable, TabelaAux) Then
            Call Access.DoCmd.DeleteObject(Access.AcObjectType.acTable, TabelaAux)
        End If
        Set tblDef = db.CreateTableDef(TabelaAux)
        With tblDef
            .SourceTableName = arrTabelas(contador)
            .Connect = ";DATABASE=" & CaminhoArq
        End With
        Call db.TableDefs.Append(tblDef)
        If Not Relatorio.Exists(TabelaAux) Then
            Call Relatorio.Add(TabelaAux, , NomeArquivo, arrTabelas(contador), arrParentID(contador), , , , True, , True)
        End If
        Call Publicas.RemoverObjetosMemoria(tblDef)
    Next contador

Exit Sub
TrataErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxTabela.CriaVinculo()", Erl, True)
End Sub

'---------------------------------------------------------------------------------------
' Modulo....: AuxTabela / Módulo
' Rotina....: FechaRecordSet / Sub
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 14/08/2013
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina desenvolvida para tentar executar o método close do recordset
'             Caso ele venha a falhar não será gerado erro devido ao resume next
'---------------------------------------------------------------------------------------
Public Sub FecharRecordSet(ByRef rs As Object)
On Error Resume Next 'caso o recordset de erro
    If Not rs Is Nothing Then
        rs.Close
    End If
End Sub

'---------------------------------------------------------------------------------------
' Modulo....: AuxTabela / Módulo
' Rotina....: DeletarRelatorio / Sub
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 14/08/2013
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina que, usando ADOX, cria uma consulta para deletar o relatorio passado
'             como parametro.
'---------------------------------------------------------------------------------------
'''Public Sub DeletarRelatorio(ByVal Tabela As String, Optional ByVal ANO As Long = 0, _
'''                           Optional ByVal MES As Integer = 0, Optional ByVal Semana As Integer = 0, _
'''                           Optional ByVal dtRef As Date = 0, Optional ByVal dtDe As Date = 0, _
'''                           Optional ByVal dtAte As Date = 0)

Public Sub DeletarRelatorio(ByVal Tabela As String, ByVal complemento As String)
On Error GoTo TrataErro
    If Relatorio.Exists(Tabela) Then Tabela = Relatorio.item(Tabela).Tabela.Name
    If Not Tabela = VBA.vbNullString Then
        Call Conexao.ModificarConsulta("Modelo_Deleta_Relatorio", "[@Campos]", "Rel.* ", _
                                       "[@Rel]", Tabela, "[@Complemento]", complemento)
        Call Conexao.DeletarRegistros(Conexao.PegarQueryNome("Deleta_Relatorio"))
        Call Conexao.RemoveQuery(Conexao.PegarQueryNome("Modelo_Deleta_Relatorio"))
    End If
Exit Sub
TrataErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxTabela.DeletarRelatorio()", Erl)
End Sub


Public Sub Compactar_BackEnd()
On Error GoTo TratarErro
Dim dbLocal         As Object
Dim tblDef          As Object ' TableDef
Dim dicTabelas      As Object
Dim dicCompactar    As Object
Dim fso             As Object
Dim fsoFile         As Object
Dim KT              As Variant
Dim KC              As Variant
Dim arr             As Variant
Dim Titulo          As String
Dim CaminhoDB       As String
Dim srcNome         As String
Dim sNomeCom        As String
Dim TotArquivos     As Integer
Dim TotArquivosDef  As Integer
Dim QTD             As Integer
Dim SairQuebraVinc  As Boolean
Dim TabelasComVinc  As Integer
Dim Incremento      As Double
Dim strMensagem     As String

Set dicTabelas = VBA.CreateObject("Scripting.Dictionary")
Set dicCompactar = VBA.CreateObject("Scripting.Dictionary")
Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    
    If MessageBoxMaster("S018") = vbNo Then Exit Sub
    
    'Set Banco de dados atual
    Titulo = "Compactando Back End..."
    Call Access.DoCmd.Close(Access.AcObjectType.acForm, "frmPrincipal")
    
    Call Access.DoCmd.openForm("frmProgressoLista", Access.AcFormView.acNormal)
    Call AuxForm.IncrementaBarraProgressoLista(5, Titulo, "Iniciando processo", VBA.vbNullString)

    Do While Not SairQuebraVinc
        Set dbLocal = Access.CodeDb
        TabelasComVinc = 0
        'Quebrar Vinculos dos BackEnds
        For Each tblDef In dbLocal.TableDefs
            If VBA.Len(tblDef.Connect) > 0 Then
                TabelasComVinc = TabelasComVinc + 1
                Call dicTabelas.Add(tblDef.Name & "|" & AuxTabela.PegarSourceTableName(tblDef.Name), tblDef.Connect)
                TotArquivosDef = TotArquivosDef + 1
                If Not dicCompactar.Exists(tblDef.Connect) Then
                    Incremento = Incremento + 0.5
                    If Incremento > 25 Then
                        Call AuxForm.IncrementaBarraProgressoLista(0.1, Titulo, "Removendo Vinculo: " & tblDef.Name, VBA.vbNullString)
                    Else
                        Call AuxForm.IncrementaBarraProgressoLista(0.5, Titulo, "Removendo Vinculo: " & tblDef.Name, VBA.vbNullString)
                    End If
                    Call dicCompactar.Add(tblDef.Connect, tblDef.Name & "|" & AuxTabela.PegarSourceTableName(tblDef.Name))
                    TotArquivos = TotArquivos + 1
                End If
                Call dbLocal.TableDefs.Delete(tblDef.Name)
                dbLocal.TableDefs.Refresh
            End If
        Next tblDef
        dbLocal.TableDefs.Refresh
        If dbLocal.TableDefs.count <= 37 Or TabelasComVinc = 0 Then SairQuebraVinc = True
        Call Publicas.RemoverObjetosMemoria(dbLocal)
    Loop
    Call Publicas.RemoverObjetosMemoria(tblDef)

    ' Renomear BackEnd para (*.bak) / Compactar BackEnds / Excluir Arqs(*.bak)
    Incremento = 40 / dicCompactar.count
    For Each KC In dicCompactar.Keys
        CaminhoDB = VBA.Mid(KC, InStr(1, KC, "DATABASE=") + Len("DATABASE="), Len(KC) - Len("DATABASE="))
        Access.DoCmd.SetWarnings False
        srcNome = VBA.Left(CaminhoDB, VBA.Len(CaminhoDB) - 6) & ".bak"
        Set fsoFile = fso.GetFile(CaminhoDB)
        fsoFile.Name = VBA.Left(fsoFile.Name, VBA.Len(fsoFile.Name) - 6) & ".bak"
        sNomeCom = fsoFile.Name
        QTD = QTD + 1
        Call AuxForm.IncrementaBarraProgressoLista(Incremento, Titulo, "Compactando: " & sNomeCom, VBA.vbNullString)
        Call Access.DBEngine.CompactDatabase(srcNome, CaminhoDB)
        Call VBA.Kill(srcNome)
        Access.DoCmd.SetWarnings True
    Next KC
    QTD = 0
    'Refazer os Vinculos com os BackEnds Compactados
    Set dbLocal = Access.CodeDb
    Incremento = 20 / dicCompactar.count
    For Each KT In dicTabelas.Keys
        arr = VBA.Split(KT, "|")
        If VBA.IsArray(arr) Then
           Set tblDef = dbLocal.CreateTableDef(arr(0))
           With tblDef
                .SourceTableName = arr(1)
                .Connect = dicTabelas.item(KT)
           End With
           QTD = QTD + 1
           Call dbLocal.TableDefs.Append(tblDef)
           If VBA.Err = 0 Then strMensagem = "OK, Link recriado "
           Call AuxForm.IncrementaBarraProgressoLista(Incremento, Titulo, "Link Refeito: " & arr(0), VBA.vbNullString)
        End If
    Next KT

    Call AuxForm.IncrementaBarraProgressoLista(100, Titulo, "Concluído", VBA.vbNullString)
    Call AuxMensagens.MessageBoxMaster("F008")

    Call Access.DoCmd.Close(Access.AcObjectType.acForm, "frmProgressoLista")
    Call Access.DoCmd.openForm("frmPrincipal", Access.AcFormView.acNormal)
    Call Publicas.RemoverObjetosMemoria(tblDef, dbLocal, dicTabelas, dicCompactar, fso, fsoFile)
On Error GoTo 0
Exit Sub
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxTabela.Compactar_BackEnd", Erl, False)
    strMensagem = "Erro : " & VBA.Err.Description
    Resume Next 'O resume next é importante em caso de erro para nao parar a rotina
    Exit Sub
    Resume
End Sub


'---------------------------------------------------------------------------------------
' Modulo....: AuxTabela / Módulo
' Rotina....: Pegar_Tabela / Sub
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 13/01/2014
' Revisão...: Fernando Fernandes (29/09/2015 - inclusão da definição da tabela pelo nome)
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina que tem como objetivo abrir a tabela referente ao tipo de relatorio
'---------------------------------------------------------------------------------------
Public Sub Pegar_Tabela(ByRef Tabela As Object, ByRef NomeTabela As String, _
                         ByVal RelAux As TipoRelatorio, ByVal CriarTabela As Boolean, _
                         ByVal dtRef As Date)
On Error GoTo TratarErro

    If CriarTabela Then
        Set Tabela = AuxTabela.PegarTabela(0, 0, RelAux, False, NomeTabela)
        
    ElseIf NomeTabela <> VBA.vbNullString Then
        Set Tabela = CurrentDb.TableDefs(NomeTabela).OpenRecordset
        
    Else
        Set Tabela = Relatorio.item(RelAux).Tabela.Abre
        NomeTabela = Relatorio.item(RelAux).Tabela.Name
        
    End If
    
    If Tabela Is Nothing Then GoTo Fim

    If VBA.IsDate(dtRef) And dtRef <> 0 Then
        Call AuxTabela.DeletarRelatorio(NomeTabela, AuxTabela.PegarComplemento_Generico(dtRef:=dtRef))
    End If
Fim:

On Error GoTo 0
Exit Sub
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxTabela.Pegar_Tabela", Erl)
End Sub


''''---------------------------------------------------------------------------------------
'''' Modulo....: AuxTabela / Módulo
'''' Rotina....: Pegar_Tabela / Sub
'''' Autor.....: Jefferson Dantas
'''' Contato...: jefferson@tecnun.com.br
'''' Data......: 13/01/2014
'''' Empresa...: Tecnun Tecnologia em Informática
'''' Descrição.: Rotina que tem como objetivo abrir a tabela referente ao tipo de relatorio
''''---------------------------------------------------------------------------------------
'''Public Sub Pegar_Tabela(ByRef tabela As Object, ByRef NomeTabela As String, _
'''                         ByVal RelAux As TipoRelatorio, ByVal CriarTabela As Boolean, _
'''                         ByVal dtRef As Date)
'''On Error GoTo TratarErro
'''
'''    If CriarTabela Then
'''        Set tabela = AuxTabela.PegarTabela(0, 0, RelAux, False, NomeTabela)
'''    Else
'''        Set tabela = Relatorio.item(RelAux).tabela.Abre
'''        NomeTabela = Relatorio.item(RelAux).tabela.name
'''    End If
'''
'''    If tabela Is Nothing Then GoTo Fim
'''
'''    If VBA.IsDate(dtRef) And dtRef <> 0 Then
'''        Call AuxTabela.DeletarRelatorio(NomeTabela, AuxTabela.PegarComplemento_Generico(dtRef:=dtRef))
'''    End If
'''Fim:
'''
'''On Error GoTo 0
'''Exit Sub
'''TratarErro:
'''    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxTabela.Pegar_Tabela", Erl)
'''End Sub


'********************FUNCOES********************

'---------------------------------------------------------------------------------------
' Modulo....: AuxTabela / Módulo
' Rotina....: PegarTabela / Function
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 13/08/2013
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina para checar e manipular as tabelas do sistema de acordo com uma data
'             de referencia
'---------------------------------------------------------------------------------------
Public Function PegarTabela(ByVal Ano As Long, ByVal Mes As Integer, ByVal RelatorioAux As TipoRelatorio, _
                            ByVal ForcaCriacaoTabela As String, ByRef NomeTabela As String) As Object
On Error GoTo TrataErro
Dim tblAux          As Object
Dim chave           As String
Dim NomeArquivo     As String
Dim CaminhoArq      As String

    chave = VBA.IIf(Ano = 0, VBA.vbNullString, Ano) & VBA.IIf(Mes = 0, VBA.vbNullString, VBA.Format(Mes, "00"))
    chave = VBA.IIf(chave = VBA.vbNullString, VBA.vbNullString, "_" & chave)

    NomeArquivo = Relatorio.item(RelatorioAux).NomeRelatorio & chave & ".accdb"
    CaminhoArq = AuxTabela.PegarCaminhoBE & "\" & NomeArquivo
    NomeTabela = Relatorio.item(RelatorioAux).Tabela.Name & chave
    If ForcaCriacaoTabela Then
        Call Cria_Tabela_Vinculo(Access.CodeDb, NomeArquivo, chave, RelatorioAux)
    ElseIf Not Relatorio.Exists(NomeTabela) Then
        Call Cria_Tabela_Vinculo(Access.CodeDb, NomeArquivo, chave, RelatorioAux)
    ElseIf Not Publicas.EnderecoValido(CaminhoArq, tipoEndereco.arquivos) Then
        Call Cria_Tabela_Vinculo(Access.CodeDb, NomeArquivo, chave, RelatorioAux)
    End If

    If Relatorio.Exists(NomeTabela) Then Set tblAux = Relatorio.item(NomeTabela).Tabela.Abre
    Set PegarTabela = tblAux
    Call Publicas.RemoverObjetosMemoria(tblAux)
Exit Function
TrataErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxTabela.PegarTabela()", Erl)
End Function

'---------------------------------------------------------------------------------------
' Modulo....: AuxTabela / Módulo
' Rotina....: PegarCaminhoBE / Function
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 14/08/2013 - Rev: 26/11/2013
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina para definir o caminho dos bancos de dados
'---------------------------------------------------------------------------------------
Public Function PegarCaminhoBE() As String
On Error GoTo TrataErro
Dim fso                 As Object
Dim NomeCaminho         As String
Dim RootPrograma        As String
Dim CaminhoAux          As String
Dim arrAux              As Variant
    
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    
    'RootPrograma = fso.GetParentFolderName(Access.CurrentProject.Path)
    RootPrograma = AuxAplicacao.PegaEnderecoCliente
    If fso.FolderExists(RootPrograma & "\BDs\Dados\") Then
        CaminhoAux = RootPrograma & "\BDs\Dados"
    Else
        arrAux = Conexao.PegarArray(Conexao.PegarRS("Pegar_CaminhoBE", ChaveUsuario))
        If VBA.IsArray(arrAux) Then
            If Not VBA.IsNull(arrAux(0, 0)) Then CaminhoAux = arrAux(0, 0)
        End If
    End If

    If Not CaminhoAux = VBA.vbNullString Then
        If fso.FolderExists(CaminhoAux) Then
            If fso.DriveExists(fso.GetFolder(CaminhoAux).Drive) Then
                If fso.GetDrive(fso.GetFolder(CaminhoAux).Drive).DriveType = Remote Then
                    NomeCaminho = fso.GetDrive(fso.GetFolder(CaminhoAux).Drive).ShareName
                    NomeCaminho = VBA.Replace(CaminhoAux, fso.GetFolder(CaminhoAux).Drive, NomeCaminho)
                Else
                    NomeCaminho = CaminhoAux
                End If
            Else
                NomeCaminho = CaminhoAux
            End If
        End If
    End If
    PegarCaminhoBE = NomeCaminho
Exit Function
TrataErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxTabela.PegarCaminhoBE()", Erl)
End Function

'---------------------------------------------------------------------------------------
' Modulo....: AuxTabela / Módulo
' Rotina....: PegarCaminhoTemplates / Function
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 10/09/2013
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina para definir o caminho dos templates baseados no banco de dados
'---------------------------------------------------------------------------------------
Public Function PegarCaminhoTemplates() As String
On Error GoTo TratarErro
Dim fso             As Object
Dim CaminhoAux      As String
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    CaminhoAux = PegaEndereco_Templates() & "\Modelos"
    PegarCaminhoTemplates = CaminhoAux
On Error GoTo 0
Exit Function
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxTabela.PegarCaminhoTemplates", Erl)
End Function

Public Function AtualizarVinculos(ByVal NomeCaminho As String) As Boolean
1   On Error GoTo TratarErro
    'Declaração de Objetos
    Dim fso As Object
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    
    Dim daoTableDef As Object ' dao.TableDef
    'Declaração de Variáveis
    Dim strCaminhoAtual As String
    Dim strConnect As String
    Dim strBdTarget As String
    Dim Erro As Boolean
    Dim Atualiza As Boolean
    Dim strStatusUpdate As String
    Dim pSubFolder As String
    Dim pFullName_Atual As String
    Dim pFullname_Novo As String

2   Call RegistraLog("")
3   Call RegistraLog(" ### VERIFICANDO A CONSISTENCIA DOS VINCULOS COM O FRONT-ED")
4   Call RegistraLog(VBA.String(150, "*"))

    'Inicializa classe Conexao
5   Call Publicas.Inicializar_Globais(False)

6   Call AuxDataBase.RecriarVinculoTabelasCadastradas(PegaEnderecoBDs)

    'Verifica quais tabelas devem ser atualizadas
7   If Not NomeCaminho = VBA.vbNullString Then
8       For Each daoTableDef In Access.CurrentDb.TableDefs
9           With daoTableDef
10              If Not daoTableDef.Name Like "TEMP_*" Then
11                  If .Attributes = 1073741824 Then  ' dao.dbAttachedTable = 1073741824

12                      strConnect = .Connect
13                      strCaminhoAtual = VBA.Mid(strConnect, VBA.InStr(strConnect, "=") + 1, ((VBA.InStrRev(strConnect, "\") - (VBA.InStr(strConnect, "=") + 1))))

14                      pSubFolder = Nz(DLookup("dbFolder", "Pegar_DBConfig", "tblName = '" & daoTableDef.Name & "'"))
15                      strBdTarget = Nz(DLookup("dbFileName", "Pegar_DBConfig", "tblName = '" & daoTableDef.Name & "'"))

16                      If pSubFolder = "" Then
17                          Call RegistraLog("   Tabela : " & daoTableDef.Name & " > Tabela vinculada porem não está mapeada na relação das tabelas que devem ser atualizadas")
18                          GoTo Proxima
19                      End If

20                      pFullName_Atual = VBA.Split(strConnect, "DATABASE=")(1)
21                      pFullname_Novo = NomeCaminho & "\" & VBA.IIf(Access.Nz(pSubFolder) <> "", pSubFolder & "\", "") & strBdTarget

22                      If Not fso.FolderExists(NomeCaminho) Then
23                          Atualiza = False
24                          strStatusUpdate = "Pasta da aplicação não existe. Tabela não atualizada"

25                      ElseIf pFullName_Atual <> pFullname_Novo Then

26                          strConnect = ";DATABASE=" & pFullname_Novo
27                          If .SourceTableName <> daoTableDef.Name Then
28                              Call Access.CurrentDb.TableDefs.Delete(daoTableDef.Name)
29                              Set daoTableDef = Access.CurrentDb.CreateTableDef(daoTableDef.Name, , daoTableDef.Name, strConnect)
30                          Else
31                              .Connect = strConnect
32                          End If

33                          .RefreshLink

34                          If VBA.Err = 0 Then
35                              strStatusUpdate = "Link atualizado ! Novo Link : " & strConnect
36                          Else
37                              Erro = True
38                              strStatusUpdate = "Erro : " & VBA.Err.Description
39                              Call RegistraLog("   Tabela : " & daoTableDef.Name & " > Ocorreu um erro ao atualizar o vinculo! : ERROR > " & strStatusUpdate)
40                          End If
41                      Else
42                          On Error Resume Next

43                          If .SourceTableName <> daoTableDef.Name Then
44                              Call Access.CurrentDb.TableDefs.Delete(daoTableDef.Name)
45                              Set daoTableDef = Access.CurrentDb.CreateTableDef(daoTableDef.Name, , daoTableDef.Name)
46                              daoTableDef.Connect = strConnect
47                              daoTableDef.RefreshLink
48                          Else
49                              .RefreshLink
50                          End If
51

52                          If VBA.Err = 3024 Then
53                              Atualiza = False
54                              Erro = True
55                              strStatusUpdate = VBA.Err.Number & "-" & VBA.Err.Description
56                              Call RegistraLog("   Tabela : " & daoTableDef.Name & " > Ocorreu um erro ao atualizar o vinculo! : ERROR > " & strStatusUpdate)
57                          Else
58                              strStatusUpdate = ("OK - Nenhuma atualização é necessário. O Link atual não foi modificado")
59                          End If
60                      End If
61                      CurrentDb.Execute "UPDATE tblDBConfig SET lastDtUpdate = Now(), status = '" & strStatusUpdate & "' WHERE tblName = '" & daoTableDef.Name & "'"
62                  End If
63              End If
64          End With
Proxima:
65      Next daoTableDef

66      If Not Erro Then
67          Call RegistraLog(" ### Verificando dos vinculos com o FRONT-ED foi finalizada com sucesso")
68          Atualiza = True
            'Call ReorganizaNavPane
            'Call VerificaFerramenta
69      Else
70          Call RegistraLog(" ### ERROR - Ocorreu um erro durante a verificação dos vinculos")
71      End If
72      Call RegistraLog(String(150, "*"))
73  Else
74      Atualiza = False
75  End If
76  AtualizarVinculos = Atualiza
77  Exit Function
TratarErro:
78  Erro = True
79  Call RegistraLog(" ### ERROR - " & VBA.Err.Description)
80  Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxTabela.AtualizaVinculos", Erl, False, False)
81  Resume Next    'não deletar pois em caso de erro é necessario que ele continue a execução
82  Exit Function
83  Resume
End Function

''''---------------------------------------------------------------------------------------
'''' Modulo....: AuxTabela / Módulo
'''' Rotina....: AtualizarVinculos / Function
'''' Autor.....: Jefferson Dantas
'''' Contato...: jefferson@tecnun.com.br
'''' Data......: 21/09/2013
'''' Empresa...: Tecnun Tecnologia em Informática
'''' Descrição.: Rotina que tem como objetivo verificar e atualizar os vinculos se necessário
''''---------------------------------------------------------------------------------------
''''Revisões
''''1.00 15/06/2015 13:46 ARMS - Incluão do argumento objConexao para identificar a instancia da classe ConexaoDB que será usada.
''''1.00 15/06/2015 13:46 ARMS - Incluido variável Curdb (Database) que vai receber a instancia do CurrentDB a ser manipulado as tabelas
'''
'''Public Function AtualizarVinculos(ByVal tpVerificacao As TipoVerificacao, Optional objConexao As ConexaoDB) As Boolean
'''On Error GoTo TrataErro
''''Declaração de Objetos
'''Dim fso                     As New FileSystemObject
'''Dim daoTableDef             As DAO.TableDef
'''Dim TabelasVinculadas       As Object 'Scripting.Dictionary
''''Declaração de Variáveis
'''Dim QtdErro                 As Integer
'''Dim NomeCaminho             As String
'''Dim strCaminhoAtual         As String
'''Dim strConnect              As String
'''Dim strBdTarget             As String
'''Dim Erro                    As Boolean
'''Dim Atualiza                As Boolean
'''Dim VerificarDicionario     As Boolean
'''Dim arrAux                  As Variant
'''Dim varKey                  As Variant
'''Dim Curdb                   object
'''
'''    'Necessário definir a instancia da classe de conexão que aponta para o banco cliente
'''20  If Not objConexao Is Nothing Then
'''30      Set Curdb = Access.Application.CurrentDb
'''40  Else
'''        Set objConexao = VariaveisEConstantes.Conexao
'''80      Set Curdb = Access.Application.CurrentDb
'''90  End If
'''
'''100 NomeCaminho = AuxTabela.PegarCaminhoBE
'''
'''110 If NomeCaminho = VBA.vbNullString Then
'''120     Atualiza = False
'''130     GoTo Fim
'''140 End If
'''    'Verifica quais tabelas devem ser atualizadas
'''150 Select Case tpVerificacao
'''    Case TipoVerificacao.TabelasLocais
'''160     Set TabelasVinculadas = objConexao.PegarDicionario("Pegar_Tabelas_Vinculadas_Locais")
'''170 Case TipoVerificacao.TabelasCriadas
'''180     Set TabelasVinculadas = objConexao.PegarDicionario("Pegar_Tabelas_Vinculadas")
'''190 End Select
'''200 If Not TabelasVinculadas Is Nothing Then
'''210     VerificarDicionario = True
'''220 End If
'''230 If Not NomeCaminho = VBA.vbNullString Then
'''240     For Each daoTableDef In Curdb.TableDefs
'''250         With daoTableDef
'''260             If .Attributes = dbAttachedTable Then
'''270                 strConnect = .Connect
'''                    'ADELSON 28/05/2015 16:10 - Ignora o tratamento de tabelas vinculadas ao Share Point
'''                    If UCase(strConnect) Like "*WSS*" Then
'''                        'Debug.Print "Tabela Share Point : " & .Name, "URL:" & strConnect
'''                        GoTo ProximaTabela
'''                    End If
'''
'''280                 strCaminhoAtual = VBA.Mid(strConnect, VBA.InStr(strConnect, "=") + 1, ((VBA.InStrRev(strConnect, "\") - (VBA.InStr(strConnect, "=") + 1))))
'''290                 strBdTarget = VBA.Mid(strConnect, VBA.InStrRev(strConnect, "\") + 1, VBA.Len(strConnect))
'''300                 If fso.FileExists(NomeCaminho & "\" & strBdTarget) Then
'''310                     .Connect = ";DATABASE=" & fso.BuildPath(NomeCaminho, strBdTarget)
'''320                     .RefreshLink
'''330                     If VerificarDicionario Then
'''340                         If TabelasVinculadas.Exists(.name) Then
'''350                             Call TabelasVinculadas.Remove(.name)
'''360                         Else
'''                                'Call objConexao.RemoveTabela(.Name)
'''370                             Call Access.DoCmd.DeleteObject(Access.AcObjectType.acTable, .name)
'''380                             Call Publicas.RemoverObjetosMemoria(daoTableDef)
'''390                             If QtdErro > 0 Then QtdErro = QtdErro - 1
'''400                         End If
'''410                     End If
'''420                 ElseIf VerificarDicionario Then
'''430                     If TabelasVinculadas.Exists(.name) Then
'''440                         Call TabelasVinculadas.Remove(.name)
'''450                         GoTo TrataErro
'''460                     Else
'''470                         Call objConexao.RemoveTabela(.name)
'''480                         Call Publicas.RemoverObjetosMemoria(daoTableDef)
'''490                     End If
'''500                 End If
'''510             End If
'''ProximaTabela:
'''520         End With
'''530     Next daoTableDef
'''
'''540     If VerificarDicionario Then
'''550         If TabelasVinculadas.Count > 0 Then
'''560             For Each varKey In TabelasVinculadas.keys
'''570                 arrAux = VBA.Split(TabelasVinculadas.item(varKey), "|")
'''580                 If VBA.IsArray(arrAux) Then
'''590                     Set daoTableDef = Curdb.CreateTableDef(varKey)
'''600                     With daoTableDef
'''610                         .SourceTableName = arrAux(1)
'''620                         .Connect = ";DATABASE=" & NomeCaminho & "\" & arrAux(0)
'''630                     End With
'''640                     Call Curdb.TableDefs.Append(daoTableDef)
'''650                 End If
'''660                 arrAux = Empty
'''670             Next varKey
'''680         End If
'''690     End If
'''700     If Not Erro Then
'''710         Atualiza = True
'''720     ElseIf QtdErro = 0 Then
'''730         Atualiza = True
'''740     End If
'''750 Else
'''760     Atualiza = False
'''770 End If
'''780 Curdb.TableDefs.Refresh
'''Fim:
'''790 AtualizarVinculos = Atualiza
'''800 Exit Function
'''TrataErro:
'''810 Erro = True
'''820 QtdErro = QtdErro + 1
'''830 Resume Next    'não deletar pois em caso de erro é necessario que ele continue a execução
'''End Function

'---------------------------------------------------------------------------------------
' Modulo....: AuxTabela / Módulo
' Rotina....: CriaTabelaDados / Function
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 14/08/2013
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina para a criação da tabela de dados de acordo com a chave baseado no
'             arquivo modelo contido nesse banco de dados
'---------------------------------------------------------------------------------------
Public Function CriaTabelaDados(ByVal NomeArquivo As String, ByVal RelatorioAux As TipoRelatorio) As String
On Error GoTo TrataErro
Dim db              As Object
Dim tblAux          As Object
Dim arquivo         As Object
Dim fldAnexo        As Object 'dao.Field2
Dim caminho         As String

    Set db = Access.CodeDb
    Set tblAux = db.OpenRecordset("tblAnexos")

    caminho = AuxTabela.PegarCaminhoBE
    If Not VBA.Right(caminho, 1) = "\" Then caminho = caminho & "\"
    With tblAux
        .MoveFirst
        Do While Not .EOF
            If .Fields("RelID").value = Relatorio.item(RelatorioAux).ID Then
                Set arquivo = .Fields("ArqAnexo").value
                Exit Do
            End If
            .MoveNext
        Loop
    End With
    If arquivo Is Nothing Then
        Call VBA.Err.Raise(9999, "Não foi possível determinar o arquivo modelo.")
    End If

    caminho = caminho & NomeArquivo
    If Not VBA.Dir(caminho) = "" Then
        Call VBA.SetAttr(caminho, VBA.vbNormal)
        Call VBA.Kill(caminho)
    End If

    Set fldAnexo = arquivo.Fields("FileData")
    fldAnexo.SaveToFile caminho
    arquivo.Close
    CriaTabelaDados = caminho
Fim:
    Call Publicas.RemoverObjetosMemoria(fldAnexo, tblAux, arquivo, db)
Exit Function
TrataErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxTabela.CriaTabelaDados()", Erl)
    GoTo Fim
    Resume
End Function

'---------------------------------------------------------------------------------------
' Modulo....: AuxTabela / Módulo
' Rotina....: PegarComplemento_Generico / Function
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 14/08/2013
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina que cria um complemento de acordo com as informações passadas pelas
'             variaveis que sao as chaves dos relatorios
'---------------------------------------------------------------------------------------
Public Function PegarComplemento_Generico(Optional ByVal Ano As Long, Optional ByVal Mes As Integer, _
                                          Optional ByVal dtRef As Date, Optional ByVal dtDe As Date, _
                                          Optional ByVal dtAte As Date) As String
On Error GoTo TrataErro
Dim Resultado           As String
Dim ResultadoAux        As String

    Resultado = "Rel " & VBA.vbNewLine

    If dtRef > 0 Then
        Resultado = Resultado & "WHERE Rel.DataRef=" & VBA.Format(dtRef, """#""mm/dd/yyyy""#""")
    ElseIf dtDe > 0 Or dtAte > 0 Then
        Resultado = Resultado & "WHERE "
        If dtDe > 0 Then
            ResultadoAux = ResultadoAux & "Rel.DataRef>=" & VBA.Format(dtDe, """#""mm/dd/yyyy""#""")
        End If
        If dtAte > 0 Then
            ResultadoAux = ResultadoAux & VBA.IIf(ResultadoAux <> VBA.vbNullString, " AND ", "") & "Rel.DataRef<=" & VBA.Format(dtAte, """#""mm/dd/yyyy""#""")
        End If
    ElseIf Ano > 0 Or Mes > 0 Then
        Resultado = Resultado & "WHERE "
        If Ano > 0 Then
            ResultadoAux = ResultadoAux & "YEAR(Rel.DataRef)=" & Ano
        End If
        If Mes > 0 Then
            ResultadoAux = ResultadoAux & VBA.IIf(ResultadoAux <> VBA.vbNullString, " AND ", "") & "MONTH(Rel.DataRef)=" & Mes
        End If
    End If
    Resultado = Resultado & ResultadoAux
    PegarComplemento_Generico = Resultado
Exit Function
TrataErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxTabela.PegarComplemento_Generico()", Erl)
End Function

'---------------------------------------------------------------------------------------
' Modulo....: AuxTabela / Módulo
' Rotina....: PegarSourceTableName / Function
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 24/09/2013
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina para pegar o nome original da tabela
'---------------------------------------------------------------------------------------
Public Function PegarSourceTableName(ByVal Tabela As String) As String
On Error GoTo TrataErro
Dim TabelaAux       As String

    If AuxTexto.IsLinhaMatch(Tabela, "[_]{1}\d{4,6}") Then
        TabelaAux = VBA.Left(Tabela, VBA.InStrRev(Tabela, "_") - 1)
    Else
        TabelaAux = Tabela
    End If
    PegarSourceTableName = TabelaAux
Exit Function
TrataErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxTabela.PegarSourceTableName()", Erl)
End Function

Function AbrirCaminhoDatabase()
    Dim fParent As Object
    On Error Resume Next
    Set fParent = Screen.ActiveForm
    Call AbrirFormulario(tag:="frmCaminhoBE", View:=acNormal, WindowMode:=acDialog)
End Function

Function PegarDescricaoDeCodigo(Tabela As String, campo_descricao As String, campo_codigo As String, valor_codigo As Integer)
    PegarDescricaoDeCodigo = Nz(Access.Application.DLookup(campo_descricao, Tabela, campo_codigo & "=" & valor_codigo), "")
End Function

'Exclui uma tabela do banco. Aguarda a liberação caso esteja bloqueada e esperando a finalização de algum processo
'21/08/2015 11:10 - Adelson - Incluido o argumento db object
Function ExcluiTabela(tbName As String, Optional db As Object)
    If db Is Nothing Then Set db = CodeDb
    If tbName = "" Then
        Debug.Print "Não é possivel excluir a tabela. Nome da tabela não informado"
        Exit Function
    End If
    Do
        VBA.DoEvents
        On Error Resume Next
        Call db.Execute("DROP TABLE [" & tbName & "]"):
        If VBA.Err = 0 Or VBA.Err = 3376 Then Exit Function
    Loop
End Function

'Cria uma tabela temporária com a copia das informações de uma origem
'21/08/2015 11:10 - Adelson - Incluido o argumento db object
Function CriarTabelaTemporaria(pSource As String, _
                               Optional db As Object, _
                               Optional pCriarCampoSelecionar As Boolean = True, _
                               Optional prefixoNomeTabela As String) As String
    Dim sTempTable As String
    Dim sqlSource As String
    If db Is Nothing Then Set db = CodeDb
10  sTempTable = "~_" & prefixoNomeTabela & "_" & AuxFileSystem.getNewGUID(10)
20  Publicas.Inicializar_Globais
    'Deleta a tabela caso exista
30  If Conexao.ObjetoExiste(db, acTable, sTempTable) Then Call ExcluiTabela(sTempTable, db)
40  sqlSource = pSource
50  If VBA.Left(UCase(pSource), 6) = "SELECT" Then sqlSource = "(" & pSource & ")"
    'Cria a tabela temporária a partir da Origem
60  Call db.Execute("SELECT " & VBA.IIf(pCriarCampoSelecionar, "0 as selecionar , ", "") & "tb.* INTO [" & sTempTable & "] FROM " & sqlSource & " as tb")
70  CriarTabelaTemporaria = sTempTable
End Function


Function CriarTabelaTemporariaDDL(pSource As String, _
                               Optional db As Object, _
                               Optional pCriarCampoSelecionar As Boolean = True, _
                               Optional prefixoNomeTabela As String) As String
    Dim sTempTable As String
    Dim sqlSource As String
    Dim pSQLFrom As String
    Dim i As Integer
    Dim rsTemp As dao.Recordset, strField As String
    
    If db Is Nothing Then Set db = CodeDb
10  sTempTable = "~_" & prefixoNomeTabela & "_" & AuxFileSystem.getNewGUID(10)
20  Publicas.Inicializar_Globais
    'Deleta a tabela caso exista
30  If Conexao.ObjetoExiste(db, acTable, sTempTable) Then Call ExcluiTabela(sTempTable, db)
40  sqlSource = pSource
50  If VBA.Left(UCase(pSource), 6) = "SELECT" Then sqlSource = "(" & pSource & ")"
    'Cria a tabela temporária a partir da Origem
    pSQLFrom = "SELECT tb.* FROM " & sqlSource & " as tb"
    
    Set rsTemp = CurrentDb.OpenRecordset(pSQLFrom)
    
    strField = vbNewLine & Space(3) & " [rowID] COUNTER"
    If pCriarCampoSelecionar Then
        strField = strField & vbNewLine & Space(3) & ",[selecionar] BIT"
    End If
    
    For i = 0 To rsTemp.Fields.count - 1
        strField = strField & vbNewLine & Space(3) & ",[" & rsTemp.Fields(i).Name & "] " & PegarTipoDado(rsTemp.Fields(i))
    Next
    
    Call CurrentDb.Execute("CREATE TABLE [" & sTempTable & "](" & strField & vbNewLine & ")")
    If TabelaExiste(sTempTable) Then
        pSQLFrom = "SELECT -1 as selecionar, tb.* FROM " & sqlSource & " as tb"
        Call CurrentDb.Execute("INSERT INTO [" & sTempTable & "] " & pSQLFrom)
    End If

70  CriarTabelaTemporariaDDL = sTempTable
End Function

Function PegarTipoDado(fd As Field) As String
    With fd
        Select Case .Type
        Case dao.DataTypeEnum.dbText
            PegarTipoDado = "TEXT(" & .size & ") "

        Case dao.DataTypeEnum.dbDate
            PegarTipoDado = "DATETIME "

        Case dao.DataTypeEnum.dbNumeric
            PegarTipoDado = "NUMERIC "

        Case dao.DataTypeEnum.dbLong
            PegarTipoDado = "LONG "

        Case dao.DataTypeEnum.dbCurrency
            PegarTipoDado = "CURRENCY "

        Case dao.DataTypeEnum.dbDouble
            PegarTipoDado = "DOUBLE "
        
        Case dao.DataTypeEnum.dbInteger
            PegarTipoDado = "INT "
        
        Case dao.DataTypeEnum.dbSingle
            PegarTipoDado = "SINGLE "
        
        Case dao.DataTypeEnum.dbMemo
            PegarTipoDado = "MEMO "
        
        End Select
    End With
End Function


Function TabelaExiste(pNomeTabela As String, Optional db As Object) As Boolean
   Call Inicializar_Globais(False)
   If db Is Nothing Then Set db = CodeDb
   TabelaExiste = Conexao.ObjetoExiste(db, acTable, pNomeTabela) And pNomeTabela <> ""
End Function

Function ConsultaExiste(pNomeConsulta As String, Optional db As Object) As Boolean
   Call Inicializar_Globais(False)
   If db Is Nothing Then Set db = CodeDb
   ConsultaExiste = Conexao.ObjetoExiste(db, acQuery, pNomeConsulta) And pNomeConsulta <> ""
End Function
'---------------------------------------------------------------------------------------
' Modulo....: AuxExport/ Módulo
' Rotina....: PegaCaminhoArquivoAnexo / Function
' Autor.....: Victor Félix
' Contato...: victor.santos@mondial.com.br
' Data......: 24/04/2013
' Empresa...: Mondial Tecnologia em Informática LTDA.
' Descrição.: Rotina que utiliza as tabelas de anexo do BackEnd de Apoio da aplicação
' para encontrar templates e salvar um arquivo real na máquina para ser trabalhado
'---------------------------------------------------------------------------------------
Public Function PegaCaminhoArquivoAnexo(ByVal strNomeTemplate As String, _
                                        ByVal strNomeTabela As String, _
                                        Optional ByVal strToDir As String = VBA.vbNullString, _
                                        Optional ByVal strFileName As String = VBA.vbNullString, _
                                        Optional bShowMsg As Boolean = True, _
                                        Optional db As Object) As String
10  On Error GoTo TratarErro

    Dim rsTabela As Object
    Dim rsArquivo As Object
    Dim fldAnexo As Object 'dao.Field2
    Dim strCaminhoArq As String
    Dim strTempDir As String
    Dim btContaRegitros As Byte

20  If db Is Nothing Then Set db = Access.CodeDb
30  Set rsTabela = db.OpenRecordset(strNomeTabela)

40  If strToDir = VBA.vbNullString Then strToDir = VBA.Environ("Temp")
50  If strFileName = VBA.vbNullString Then strFileName = "~_" & AuxFileSystem.getNewGUID(15, True) & ".mytmp"

60  strTempDir = strToDir
    
    Set rsTabela = db.OpenRecordset("SELECT * FROM " & strNomeTabela & " WHERE NomeTemplate  = '" & strNomeTemplate & "'")
'70  rsTabela.FindFirst "NomeTemplate='" & strNomeTemplate & "'"

80  If rsTabela.EOF Then
90      If bShowMsg Then MessageBoxMaster "Template para relatório '" & strNomeTemplate & "' não foi localizado na tabela '" & strNomeTabela & "'", VBA.vbExclamation
100 Else
110     Set rsArquivo = rsTabela.Fields("AnexoTemplate").value

120     If strFileName = "OriginalName" Then strFileName = rsArquivo!FileName.value

130     strCaminhoArq = strTempDir & "\" & strFileName

140     If Not VBA.Dir(strCaminhoArq) = "" Then
150         VBA.SetAttr strCaminhoArq, VBA.vbNormal
160         On Error Resume Next
170         VBA.Kill strCaminhoArq
180     End If

190     Set fldAnexo = rsArquivo.Fields("FileData")

200     fldAnexo.SaveToFile strCaminhoArq
210     rsArquivo.Close
220     PegaCaminhoArquivoAnexo = strCaminhoArq
230 End If

240 Call Publicas.RemoverObjetosMemoria(fldAnexo, rsTabela, rsArquivo, db)

250 Exit Function
TratarErro:
260 Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxExport.PegaCaminhoArquivoAnexo()", Erl, , False)
270 Exit Function
280 Resume
End Function

