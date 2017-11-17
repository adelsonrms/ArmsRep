Attribute VB_Name = "AuxDataBase"
Option Compare Database
Dim db         As Object

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : auxBackEnd.RecriarVinculoTabelasCadastradas()
' TIPO             : Sub
' DATA/HORA        : 19/04/2016 18:39
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Força a recriação dos vinculos com as tabelas cadastras
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      19/04/2016 18:39
'---------------------------------------------------------------------------------------
Function RecriarVinculoTabelasCadastradas(strCaminho As String) As Boolean
    Dim rsTables As Object
    Dim tbDef As Object
    Dim pSubFolder As String
    Dim strBdTarget As String

    '---------------------------------------------------------------------------------------
1   On Error GoTo RecriarVinculoTabelasCadastradas_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
2   Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Sub auxBackEnd.RecriarVinculoTabelasCadastradas()"
    '---------------------------------------------------------------------------------------

3   Set rsTables = CodeDb.OpenRecordset("Pegar_DBConfig")

4   Do While Not rsTables.EOF
5       Set tbDef = Nothing
6       On Error Resume Next
7       Set tbDef = CodeDb.TableDefs(rsTables!tblname.value)

8       If VBA.Err <> 0 Then
9           Debug.Print rsTables!tblname.value, "Tabela nao existe, será vinculada"
14          pSubFolder = rsTables!dbFolder.value
15          strBdTarget = rsTables!dbfilename.value
16          If pSubFolder = "" Then
17              Call RegistraLog("   Tabela : " & rsTables.Name & " > Tabela vinculada porem não está mapeada na relação das tabelas que devem ser atualizadas")
18              GoTo Proxima
19          End If
            Call CriarTabelaVinculada(strCaminho & "\" & VBA.IIf(Nz(pSubFolder) <> "", pSubFolder & "\", "") & strBdTarget, rsTables!tblname.value)
10      Else
11          If CodeDb.TableDefs(rsTables!tblname.value).Attributes = dao.dbAttachedTable Then
12           '  Debug.Print rsTables!tblName.Value, "OK, Tabela ja vinculada..."
21          End If
        End If
Proxima:
22          rsTables.MoveNext
23  Loop

    RecriarVinculoTabelasCadastradas = lngErrorNumber = 0

Fim:
24  On Error GoTo 0
25  Exit Function

RecriarVinculoTabelasCadastradas_Error:
26  If VBA.Err <> 0 Then
27      lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
        'Debug.Print cstr_ProcedureName, "Linha : " & VBA.Erl() & " - " & strErrorMessagem
28      Call Excecoes.TratarErro(strErrorMessagem, lngErrorNumber, cstr_ProcedureName, VBA.Erl())
29  End If
    GoTo Fim:
    'Debug Mode
30  Resume
End Function


'--------------------------------------------------------------------------------------------------------------------------------
'Autor      : Adelson Silva - Mondial Informatica (adelsons@gmail.com)
'Data/Hora  : 31/10/2013 18:18
'Descrição  : Cria ou atualiza uma tabela vinculada a um arquivo de origem
'---------------------------------------------------------------------------------------------------------------------------------
'# Paramentros de Entrada
'---------------------------------------------------------------------------------------------------------------------------------
' pDbDestination            : Instancia Database do banco de dados de destino
' pFileSource            : Endereço completo do arquivo de banco de dados de origem
' tableSource       : Nome da tabela de origem
' tableDestination  : Nome da tabela de destino. Pode ser ignorada (optional). Se nao passar, assume o nome da tabela de origem
'---------------------------------------------------------------------------------------------------------------------------------
'# Revisões
'---------------------------------------------------------------------------------------------------------------------------------
'31/10/2013 - Criado a função
'31/10/2013 - Organização e documentação do módulo
'03/02/2014 - Adicionado argumento PWD para conectar ao um banco protegido por senha
'30/09/2015 - Implementado no projeto Bradesco BBI
'28/06/2017 - Alterado nomenclatura da assinatura
'---------------------------------------------------------------------------------------------------------------------------------

'Function CriarTabelaVinculada(pFileSource As String, _
'                              pTableNameSource As String, _
'                              Optional pDbDestination As Object, _
'                              Optional pTableNameDestination As String, _
'                              Optional pDBPassword As String) As Variant
                              
                              
Function CriarTabelaVinculada(pArquivoOrigem As String, _
                              pTabelaOrigem As String, _
                              Optional pTabelaDestino As String, _
                              Optional pSenha As String) As Variant

10  On Error GoTo Err_CriarTabelaVinculada

    Dim strConnect As String
    Dim tdfLkd As Object 'TableDef
    Dim rstLinked As Object
    Dim addNew As Boolean
    Dim sErro As String
    Dim bUpdate As Boolean
    Dim bUpdateLink As Boolean
    Dim sOLD As String

20  On Error Resume Next

    strConnect = Build_CONNECT_LINKED_TABLE(pArquivoOrigem, , pSenha)

    If pDbDestination Is Nothing Then Set pDbDestination = CurrentDb()

30  If pTabelaDestino = "" Then pTabelaDestino = pTabelaOrigem

40  Set tdfLkd = pDbDestination.TableDefs(pTabelaDestino)

50  If VBA.Err.Number = 0 Then
60      With tdfLkd
70          If .Connect = strConnect Then
80              GoTo Finaliza
90          Else
100             bUpdateLink = True
110             sOLD = .Connect
120             pDbDestination.TableDefs.Delete pTabelaDestino: pDbDestination.TableDefs.Refresh
130         End If
140     End With
150 End If

160 On Error GoTo Err_CriarTabelaVinculada
170 Set tdfLkd = pDbDestination.CreateTableDef(pTabelaDestino): addNew = True

180 On Error GoTo Err_CriarTabelaVinculada

190 With tdfLkd
200     If .Connect <> strConnect Then
210         .Connect = strConnect: bUpdate = True
220         .SourceTableName = pTabelaOrigem
230     End If
240 End With

Finaliza:
250 If bUpdate Then
260     If addNew Then Call pDbDestination.TableDefs.Append(tdfLkd)
270     pDbDestination.TableDefs.Refresh

280     If bUpdateLink Then

290         LogLinkedTable sOperação:=pDbDestination.Name & " > New db path, Change Link Tables....", _
                           bPrint:=False, _
                           oHost:=CurrentProject


300         LogLinkedTable sOperação:=pDbDestination.Name & " > OLD     : Linked Table SUCESS > TABLE : [" & tdfLkd.Name & "].CONNECT = " & sOLD, _
                           bPrint:=False, _
                           oHost:=CurrentProject

310         LogLinkedTable sOperação:=pDbDestination.Name & " > UPDATED : Linked Table SUCESS > TABLE : [" & tdfLkd.Name & "].CONNECT = " & tdfLkd.Connect, _
                           bPrint:=False, _
                           oHost:=CurrentProject
320     Else
330         LogLinkedTable sOperação:=pDbDestination.Name & " > CREATE NEW : Linked Table SUCESS > TABLE : [" & tdfLkd.Name & "].CONNECT = " & tdfLkd.Connect, _
                           bPrint:=False, _
                           oHost:=CurrentProject
340     End If

350     If tdfLkd.Connect <> "" Then
360         On Error Resume Next
370         tdfLkd.RefreshLink
380         If VBA.Err = 0 Then
390             CriarTabelaVinculada = Array(0, "A table was linked successfully", tdfLkd.SourceTableName, tdfLkd.Name, tdfLkd.LastUpdated)
400         Else
410             CriarTabelaVinculada = Array(1, "ERROR - " & VBA.Err.Description, "", "", VBA.Now)
420         End If
430     Else
440         CriarTabelaVinculada = Array(2, "ERROR - TABLE NOT LINKED", "", "", VBA.Now)
450     End If
460 Else
470     If sErro <> "" Then
480         LogLinkedTable sOperação:=sErro, _
                           bPrint:=True, _
                           oHost:=CurrentProject

490     Else
500         LogLinkedTable sOperação:=CurrentProject.Name & " > Nada alterado. Table '" & pTabelaOrigem & "' ja vinculada á origem : " & pArquivoOrigem, _
                           bPrint:=False, _
                           oHost:=CurrentProject
510     End If
520 End If

530 sErro = ""

540 Exit Function

Err_CriarTabelaVinculada:
550 If VBA.Err <> 0 Then
560     sErro = VBA.Dir(pDbDestination.Name) & " > Linked Table ERROR > Erro configurar Linked Table [" & pTabelaOrigem & "] no BD Origem : '" & pArquivoOrigem & "." & VBA.vbNewLine & " Internal VBA.Error : " & VBA.Error
570     bUpdate = False
580     CriarTabelaVinculada = Array(1, "ERROR - " & VBA.Err.Description, VBA.Now)
        Debug.Print sErro
590     Exit Function
600 End If
610 Exit Function
620 Resume
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : auxLinkTable.AnalisarTabelaVinculada()
' TIPO             : Function
' DATA/HORA        : 30/09/2015 17:33
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Analisa uma tabela vinculada e retorna informações sobre a mesma
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      30/09/2015 17:33
' * 1.01      30/09/2015 17:33                      Alterado o nome da função de : 'AnalyzeLinkedTable' para 'AnalisarTabelaVinculada'
'---------------------------------------------------------------------------------------
Public Function AnalisarTabelaVinculada(pTableToAnalyze As String, Optional dbHost As Object, Optional bShowMsg As Boolean = True) As Variant
    Dim dbLkd As Object
    Dim tdf As Object 'TableDef
    Dim sConnect As String
    Dim colConnectString As Collection
    Dim cs As Variant

    '0 - OK, a tabela esta linkada e atualizada
    '1 - A tabela esta linkada, porem não foi atualizada por algum motivo
    '2 - A tabela configurada ainda não existe no banco de dados selecionado
    '----------------------------------------------------------------------------------------------------
10  On Error GoTo AnalisarTabelaVinculada_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
20  Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Function auxLinkTable.AnalisarTabelaVinculada()"
    '----------------------------------------------------------------------------------------------------

40  On Error Resume Next

50  If dbHost Is Nothing Then Set dbHost = CurrentDb()

60  Set tdf = dbHost.TableDefs(pTableToAnalyze)
    'Se a tabela foi encontrada, significa que ja esta vinculada.
70  If Not tdf Is Nothing Then
        'Bom, se ja esta vinculada, tenta atualizar o link
80      On Error Resume Next
90      tdf.RefreshLink

100     VBA.Err.Clear
110     Set colConnectString = ParseConnect(tdf.Connect)

120     sConnect = ""
130     For Each cs In VBA.Split(tdf.Connect, ";")
140         If InStr(CStr(cs), "=") > 0 Then
150             If UCase(VBA.Split(CStr(cs), "=")(0)) = "PWD" Then
160                 sConnect = sConnect & VBA.Split(CStr(cs), "=")(0) & "=*****;"
170             Else
180                 sConnect = sConnect & VBA.Split(CStr(cs), "=")(0) & "=" & VBA.Split(CStr(cs), "=")(1) & ";"
190             End If
200         Else
210             sConnect = sConnect & CStr(cs) & ";"
220         End If
230     Next

240     If VBA.Err = 0 Then
250         AnalisarTabelaVinculada = Array(0, "OK - Linked Connected", tdf.SourceTableName, tdf.Name, tdf.LastUpdated, sConnect, colConnectString("DATABASE"))
260     Else
270         AnalisarTabelaVinculada = Array(1, "Linked with error ! (" & VBA.Err.Description & ")", "", "", VBA.Now, sConnect, colConnectString("DATABASE"))
280     End If
290 Else
300     AnalisarTabelaVinculada = Array(2, "Table Not found in Database!", "", "Table Not found in Database", "Table Not found in Database")
310 End If

Fim:
330 On Error GoTo 0
340 Exit Function

AnalisarTabelaVinculada_Error:
350 If VBA.Err <> 0 Then
360     lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
370     Call Excecoes.TratarErro(strErrorMessagem, lngErrorNumber, cstr_ProcedureName, VBA.Erl)
380 End If
    GoTo Fim:
    'Debug Mode
390 Resume
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : auxLinkTable.ParseConnect()
' TIPO             : Function
' DATA/HORA        : 30/09/2015 17:34
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Analisa a String de conexao da uma tabela vinculada
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      30/09/2015 17:34
'---------------------------------------------------------------------------------------
Function ParseConnect(pConnectionString As String) As Collection
    Dim vConnect As Variant
    Dim vPart As Variant
    Dim cs As New Collection
    Dim pKey As String, pValue As String
    Dim pExist As Boolean
    
    '----------------------------------------------------------------------------------------------------
10  On Error GoTo ParseConnect_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
20  Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Function auxLinkTable.ParseConnect()"
    '----------------------------------------------------------------------------------------------------

30  If pConnectionString <> "" Then
40      vConnect = VBA.Split(pConnectionString, ";")
50      For Each vPart In vConnect
60          If vPart <> "" Then
70              If InStr(vPart, "=") > 0 Then
                    pKey = VBA.Split(CStr(vPart), "=")(0)
                    pValue = VBA.Split(CStr(vPart), "=")(1)
                    On Error Resume Next
                    pValue = cs(pKey)
                    pExist = VBA.Err = 0
                    On Error GoTo ParseConnect_Error
                    If Not pExist Then Call cs.Add(pValue, pKey)
90              Else
100                 cs.Add "", CStr(vPart)
110             End If
120         End If
130     Next
140 End If
150 Set ParseConnect = cs
Fim:
160 On Error GoTo 0
170 Exit Function

ParseConnect_Error:
180 If VBA.Err <> 0 Then
190     lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
200     Call Excecoes.TratarErro(strErrorMessagem, lngErrorNumber, cstr_ProcedureName, VBA.Erl)
210 End If
    GoTo Fim:
    'Debug Mode
220 Resume
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : auxLinkTable.LogLinkedTable()
' TIPO             : Sub
' DATA/HORA        : 30/09/2015 17:30
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Registra um log ao vincular uma tabela
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      30/09/2015 17:30
'---------------------------------------------------------------------------------------
Public Sub LogLinkedTable(sOperação As String, _
                          Optional bStartLog As Boolean = False, _
                          Optional bSendToLoading As Boolean = False, _
                          Optional bPrint As Boolean, _
                          Optional bCloseHifen As Boolean = False, _
                          Optional pCaracter As String = "-", _
                          Optional pFileNameLog As String = "", _
                          Optional PrintDatetime As Boolean = True, _
                          Optional bShowFileInNotePad As Boolean = False, Optional oHost As Object)

    Dim NumFileLog As Integer
    Dim msgLog As String
    Dim FileNameLog As String

    '----------------------------------------------------------------------------------------------------
10  On Error GoTo LogLinkedTable_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
20  Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Sub auxLinkTable.LogLinkedTable()"
    '----------------------------------------------------------------------------------------------------

30  If Not oHost Is Nothing Then If VBA.Dir(oHost.Path & "\Log\", VBA.vbDirectory) = "" Then VBA.MkDir oHost.Path & "\Log\"

40  If pFileNameLog <> "" Then
50      FileNameLog = pFileNameLog
60  Else
70      FileNameLog = oHost.Path & "\Log\LinkedTables.log"
80  End If
90  msgLog = VBA.IIf(PrintDatetime, VBA.Now & VBA.vbTab & Environ("ComputerName") & "\" & Environ("UserName") & VBA.vbTab & "|" & VBA.vbTab, "") & sOperação
100 NumFileLog = openFileLog(FileNameLog, bStartLog, PrintDatetime)

110 If bCloseHifen Then Print #NumFileLog, String(300, pCaracter)
120 Print #NumFileLog, msgLog
130 If bCloseHifen Then Print #NumFileLog, String(300, pCaracter)
140 Close #NumFileLog

150 If bPrint Then Debug.Print msgLog

160 dtLastTime = VBA.Now

Fim:
170 On Error GoTo 0
180 Exit Sub

LogLinkedTable_Error:
190 If VBA.Err <> 0 Then
200     lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
210     Call Excecoes.TratarErro(strErrorMessagem, lngErrorNumber, cstr_ProcedureName, VBA.Erl)
220 End If
    GoTo Fim:
    'Debug Mode
230 Resume

End Sub

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : auxLinkTable.openFileLog()
' TIPO             : Function
' DATA/HORA        : 30/09/2015 17:30
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Cria um novo arquivo de log
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      30/09/2015 17:30
'---------------------------------------------------------------------------------------
Public Function openFileLog(pathLog As String, Optional bClear As Boolean = False, Optional PrintDatetime As Boolean = True) As Integer
'----------------------------------------------------------------------------------------------------
10  On Error GoTo openFileLog_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
20  Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Function auxLinkTable.openFileLog()"
    '----------------------------------------------------------------------------------------------------

30  NumFile = VBA.FreeFile()
    'Reinicia o log caso esteja cheio
40  If VBA.Dir(pathLog) <> "" Then
50      If VBA.FileSystem.FileLen(pathLog) > 300000 Then
60          Name pathLog As pathLog & "_" & VBA.Format(VBA.Now, "yyyymmddhhnnss") & ".txt"
70          bClear = True
80      End If
90  End If
100 If VBA.Dir(pathLog) = "" Or bClear Then
110     Open pathLog For Output As #NumFile
120     If PrintDatetime Then
130         Print #NumFile, VBA.Left("DATA / HORA" & VBA.String(VBA.Len(VBA.Now), " "), VBA.Len(VBA.Now)) & VBA.vbTab & "MAQUINA \ USUARIO" & VBA.vbTab & "LOG GERADO PELA FERRAMENTA em : " & VBA.Now
140         Print #NumFile, VBA.String(120, "-")
150     End If
160 Else
170     Open pathLog For Append As #NumFile
180 End If
190 openFileLog = NumFile
Fim:
200 On Error GoTo 0
210 Exit Function

openFileLog_Error:
220 If VBA.Err <> 0 Then
230     lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
240     Call Excecoes.TratarErro(strErrorMessagem, lngErrorNumber, cstr_ProcedureName, VBA.Erl)
250 End If
    GoTo Fim:
    'Debug Mode
260 Resume
End Function
'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : AuxImportText.importDataToAccess()
' TIPO             : Function
' DATA/HORA        : 05/02/2015 19:17
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Essa função efetua a carga de dados de um arquivo texto diretamente
'                    para um arquivo Access
'                    Utiliza tecnicas ADO para descarregar os dados diretamente sem a
'                    necessidade de leitura linha a linha
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      05/02/2015 19:17    Adelson         Criação/Atualização do procedimento
'---------------------------------------------------------------------------------------
Function importDataToAccess(pToTable As String, _
                            pFromFile As String, _
                            Optional pToDBFile As String, _
                            Optional pTableName As String = "Sheet1$") As Long
    Dim lngErrorNumber As Long, strErrorMessagem As String
20  Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Function AuxImportText.importDataToAccess()"
    '----------------------------------------------------------------------------------------------------

30  On Error GoTo Err_Handle

    Dim cDB    As Object
    Dim ADODrive As cTFW_DBDrive
    Dim strSQL As String
    Dim qtdRows As Long
    Dim strConnect As String
    Dim qtdRowsToImport As Long

    'String de Conexão para carga de arquivo texto para access
    '------------------------------------------------------------------------------------------
40  strConnect = "Text;"
    strConnect = strConnect & "DSN=IMPORT_TXT_CASH_FLOW;"
50  strConnect = strConnect & "FMT=Delimited;"
60  strConnect = strConnect & "HDR=NO;"
70  strConnect = strConnect & "IMEX=2;"
80  strConnect = strConnect & "CharacterSet=850;"
90  strConnect = strConnect & "ACCDB=YES;"
100 strConnect = strConnect & "DATABASE=" & mdlShell.GetDirectoryPath(pFromFile) & ""
    '------------------------------------------------------------------------------------------
110 pTableName = VBA.Replace(VBA.Dir(pFromFile), ".", "#")

120 If pToDBFile = "" Then pToDBFile = CurrentDb.Name

130 Set ADODrive = New cTFW_DBDrive

    'Configura o arquivo schema.ini para a definição de carga dos dados
'140 Call ADODrive.createSchemaFileForImport(pFromFile, "Format", "Delimited(|)", "ColNameHeader", "False", "MaxScanRows", 0)

    'String com o Comando SQL apontando o caminho do arquivo e a planilha
150 strSQL = "SELECT * INTO [" & pToTable & "] FROM [" & strConnect & "].[" & pTableName & "]"
    
    qtdRowsToImport = getRowsInTextFile(pFromFile)
    RegistraLog "Quantidade de Registros a serem importado..." & qtdRowsToImport
    
    'Abre a conexão
160 Set cDB = ADODrive.OpenConnection(pToDBFile)

    'Inicia uma transaction
170 With cDB
        .BeginTrans
        'Não Assincrono (qtdRows retornará a quantidade de registros)
180     On Error Resume Next
190     .Execute "DROP TABLE [" & pToTable & "]"
200     RegistraLog "   ADO /  ACCESS > IMPORTING FROM XL FILE : " & pFromFile
210     On Error Resume Next
220     Call .Execute(strSQL, qtdRows, ExecuteOptionEnum.adExecuteNoRecords)
230     RegistraLog "   ADO /  ACCESS > IMPORTING FINISHED....Rows affected (" & qtdRows & ")"

240     If VBA.Err = 0 Then
            .Execute "DELETE FROM tblCFDATA AS tb WHERE tb.TRADE_REF_ID='*'"
            'Em uma execução Assincrono, Connection.Statet estará com os valores [adStateOpen + adStateExecuting] indicando que o comando ainda está sendo executando
            'Em uma execução Não Assincrona essa linha não será executada pois o state estará somente com o valor adStateOpen
250         Do While .State = ADODB.ObjectStateEnum.adStateOpen + ADODB.ObjectStateEnum.adStateExecuting
260             Debug.Print VBA.Now, "Executando....."
270         Loop
            'Salva a alteração
280         .CommitTrans
290         RegistraLog "   ADO Commited > Importação finalizada com sucesso na base (" & qtdRows & ")"
300     Else
            'Reverte as alterações
310         .RollbackTrans
320         RegistraLog "   ADO Rollback > Ocorreu um erro no comando. Revertendo alterações"
330         RegistraLog "   ADO Rollback > " & VBA.Error
340     End If
    End With

Finaliza:
350 cDB.Close
360 Set cDB = Nothing
    
    If qtdRowsToImport = qtdRows Then
        Call RegistraLog("Sucesso na importação..." & qtdRows & " foram importadas !")
    End If

370 importDataToAccess = qtdRows
    
    

470 On Error GoTo 0
480 Exit Function

Err_Handle:
390 If VBA.Err <> 0 Then
400     RegistraLog "   ERRO AO IMPORTAR OS DADOS DO EXCEL P/ O ACCESS > " & VBA.Error
410     RegistraLog "   mdlImport.importDataToAccess() INTERNAL ERROR > " & VBA.Error
420     VBA.MsgBox "Ocorreu um erro inesperado na importação do Excel p/ o Access !" & VBA.vbNewLine & VBA.vbNewLine & "Internal VBA.Error : " & VBA.Error, VBA.vbCritical, "Error"
        qtdRows = -1
430     GoTo Finaliza
440 End If
450 Exit Function
460 Resume
End Function

Function getRowsInTextFile(pTextFile As String, Optional bHasHeader As Boolean) As Long
    On Error GoTo Err_Handle
    Dim objTextFile As Object
    Set objTextFile = mdlShell.getFSOFile(pTextFile).OpenAsTextStream(ForReading)
    Call objTextFile.ReadAll
    getRowsInTextFile = VBA.IIf(bHasHeader, objTextFile.Line - 1, objTextFile.Line)
    objTextFile.Close
    Set objTextFile = Nothing
    Exit Function
Err_Handle:
390 If VBA.Err <> 0 Then
        RegistraLog "Erro ao tentar determinar a quantidade de linhas no arquivo '" & pTextFile & "'"
        RegistraLog "## " & VBA.Error
    End If
End Function
'----------------------------------------------------------------
'Campos para configurar a especificação
'----------------------------------------------------------------
'DateDelim
'DateFourDigitYear
'DateLeadingZeros
'DateOrder
'DecimalPoint
'FieldSeparator
'FileType
'SpecID
'SpecName
'SpecType
'StartRow
'TextDelim
'TimeDelim

Sub configuracaoEspecificacaoImportacao(db As Object, pSpecName As String, ParamArray especificacoes() As Variant)
    Dim rsSpec As Object
    Dim ispec As Integer
    
    Call CreateTableSpec(db)
    db.TableDefs.Refresh
    Set rsSpec = db.OpenRecordset("SELECT * FROM MSysIMEXSpecs WHERE SpecName = '" & pSpecName & "'")
    
    If rsSpec.EOF Then rsSpec.addNew Else rsSpec.edit
    
    With rsSpec
        '------------------------------------------------------------
        'Configurações Padrões
        '------------------------------------------------------------
        !DateDelim.value = "/"
        !DateFourDigitYear.value = True
        !DateLeadingZeros.value = False
        !DateOrder.value = 2
        !DecimalPoint.value = "."
        !FieldSeparator.value = ";"
        !FileType.value = 437
        !SpecID.value = 1
        !SpecName.value = pSpecName
        !SpecType.value = 1
        !StartRow.value = 1
        !TextDelim.value = ""
        !TimeDelim.value = ":"
        '------------------------------------------------------------
        'Altera as configurações desejadas
        '------------------------------------------------------------
        For ispec = 0 To UBound(especificacoes) Step 2
            .Fields(especificacoes(ispec)).value = especificacoes(ispec + 1)
        Next
        
        .Update
        
    End With
    
End Sub

Function CreateTableSpec(Optional db As Object)
    Dim pDDL As String
    Dim tb As Object 'TableDef
    
    If db Is Nothing Then Set db = CurrentDb
    
Verifica:
    On Error Resume Next
    Set tb = db.TableDefs("MSysIMEXSpecs")
    
    If VBA.Err.Number = 3265 Then
        pDDL = pDDL & VBA.vbNewLine & "CREATE TABLE MSysIMEXSpecs ("
        pDDL = pDDL & VBA.vbNewLine & "      DateDelim Text(2)"
        pDDL = pDDL & VBA.vbNewLine & "     ,DateFourDigitYear  YESNO"
        pDDL = pDDL & VBA.vbNewLine & "     ,DateLeadingZeros  YESNO"
        pDDL = pDDL & VBA.vbNewLine & "     ,DateOrder  INTEGER"
        pDDL = pDDL & VBA.vbNewLine & "     ,DecimalPoint  TEXT (2)"
        pDDL = pDDL & VBA.vbNewLine & "     ,FieldSeparator  TEXT (2)"
        pDDL = pDDL & VBA.vbNewLine & "     ,FileType  INTEGER"
        pDDL = pDDL & VBA.vbNewLine & "     ,SpecID  COUNTER"
        pDDL = pDDL & VBA.vbNewLine & "     ,SpecName  TEXT (64)"
        pDDL = pDDL & VBA.vbNewLine & "     ,SpecType  BYTE"
        pDDL = pDDL & VBA.vbNewLine & "     ,StartRow  LONG"
        pDDL = pDDL & VBA.vbNewLine & "     ,TextDelim  TEXT (2)"
        pDDL = pDDL & VBA.vbNewLine & "     ,TimeDelim  TEXT (2)"
        pDDL = pDDL & VBA.vbNewLine & ")"
        Call db.Execute(pDDL)
        db.TableDefs.Refresh
        Access.Application.SetHiddenAttribute acTable, "MSysIMEXSpecs", True
        GoTo Verifica
    End If
    

End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : AuxImportText.CriarConsultaVinculadaText()
' TIPO             : Function
' DATA/HORA        : 01/09/2015 18:45
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Cria uma consulta vinculada à um arquivo TXT
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      01/09/2015 18:45    Adelson         Criação/Atualização do procedimento
'---------------------------------------------------------------------------------------
Function CriarConsultaVinculadaText(pOrigem As String, _
                                    strEspecificacao As String, _
                                    Optional strNomeConsulta As String, _
                                    Optional bRetornaRecorset As Boolean = False, _
                                    Optional definicao_campos As String = "*", _
                                    Optional FiltroWhere As String = "", _
                                    Optional Demilitador As String = "Delimited") As Variant
                                    
10  On Error GoTo CriarConsultaVinculadaText_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
20  Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Function AuxImportText.CriarConsultaVinculadaText()"
    '----------------------------------------------------------------------------------------------------
    Dim strTempFile As String
    Dim qdf As Object
    
30  If strNomeConsulta = "" Then strNomeConsulta = "~QRYTMP_" & VBA.Format(VBA.Now, "YYYYMMDD_HHNNSS")

    'Copia o arquivo para uma pasta temporária para efetuar a importação
40  strTempFile = Environ("Temp") & "\" & strNomeConsulta & ".txt"
50  Call FileSystem.FileCopy(pOrigem, strTempFile)

    'String de Conexão para carga de arquivo texto para access
    '------------------------------------------------------------------------------------------
60  strConnect = "Text;"
70  strConnect = strConnect & "DSN=" & strEspecificacao & ";"
80  strConnect = strConnect & "FMT=" & Demilitador & ";" 'Fixed
90  strConnect = strConnect & "HDR=NO;"
100 strConnect = strConnect & "IMEX=2;"
110 strConnect = strConnect & "CharacterSet=850;"
120 strConnect = strConnect & "ACCDB=YES;"
130 strConnect = strConnect & "DATABASE=" & PegarPasta(strTempFile)
    '------------------------------------------------------------------------------------------
140 pTableName = VBA.Replace(VBA.Dir(strTempFile), ".", "#")

150 If pToDBFile = "" Then pToDBFile = CurrentDb.Name

160 On Error Resume Next
170 Set qdf = CurrentDb.QueryDefs(strNomeConsulta)
180 If Not qdf Is Nothing Then Call CurrentDb.Execute("DROP TABLE " & strNomeConsulta)
190 On Error GoTo CriarConsultaVinculadaText_Error

    'String com o Comando SQL apontando o caminho do arquivo e a planilha
200 strSQL = "SELECT " & definicao_campos & " FROM [" & strConnect & "].[" & pTableName & "] as T " & FiltroWhere

210 Set qdf = CurrentDb.CreateQueryDef(strNomeConsulta)
220 qdf.SQL = strSQL
230 CriarConsultaVinculadaText = Array(strNomeConsulta, VBA.IIf(bRetornaRecorset, qdf.OpenRecordset, Nothing))
Fim:
240 On Error GoTo 0
250 Exit Function

CriarConsultaVinculadaText_Error:
260 If VBA.Err <> 0 Then
270     lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
280     Call Excecoes.TratarErro(strErrorMessagem, lngErrorNumber, cstr_ProcedureName, VBA.Erl)
290 End If
    GoTo Fim:
    'Debug Mode
300 Resume

End Function



'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : auxLinkTable.Build_CONNECT_LINKED_TABLE()
' TIPO             : Function
' DATA/HORA        : 18/11/2015 10:54
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Monta uma String de Conexão para criar uma tabela vinculada
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      18/11/2015 10:54
'---------------------------------------------------------------------------------------
Function Build_CONNECT_LINKED_TABLE(pDataBase As String, Optional SourceType As String, Optional pPWD As String, Optional Extendend As String) As String
    Dim sConnect As String
    Dim lngPonto As Long

    '---------------------------------------------------------------------------------------
10  On Error GoTo Build_CONNECT_LINKED_TABLE_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
20  Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Function auxLinkTable.Build_CONNECT_LINKED_TABLE()"
    '---------------------------------------------------------------------------------------
30  If SourceType = "" Then
40      lngPonto = VBA.InStrRev(pDataBase, ".")
50      SourceType = UCase(Right(pDataBase, Len(pDataBase) - lngPonto))
60  End If

70  Select Case SourceType
    Case "XLSX", "XLSM", "XLSB"
80      sConnect = "Excel 12.0 Xml;HDR=YES;IMEX=2;ACCDB=YES;DATABASE=" & pDataBase & ";" & Extendend
    Case "XLS", "XLB"
        sConnect = "Excel 8.0;HDR=YES;IMEX=2;ACCDB=YES;DATABASE=" & pDataBase & ";" & Extendend
90  Case "ACCDB"
100     sConnect = ";DATABASE=" & pDataBase & VBA.IIf(pPWD <> "", ";PWD=" & pPWD, "")
110 Case "TXT", "CSV"

120 Case "TEXT"
130 End Select
140 Build_CONNECT_LINKED_TABLE = sConnect

Fim:
150 On Error GoTo 0
160 Exit Function

Build_CONNECT_LINKED_TABLE_Error:
170 If VBA.Err <> 0 Then
180     lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
        'Debug.Print cstr_ProcedureName, "Linha : " & VBA.Erl() & " - " & strErrorMessagem
190     Call Excecoes.TratarErro(strErrorMessagem, lngErrorNumber, cstr_ProcedureName, VBA.Erl())
200 End If
    GoTo Fim:
    'Debug Mode
210 Resume
End Function
'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : mdlDBConfig.getCodeDbFolderName()
' TIPO             : Function
' DATA/HORA        : 11/09/2014 15:05
' CONSULTOR        : (ADELSON)/ TECNUN - Adelson Rosendo Marques da Silva
' CONTATO          : adelson@tecnun.com.br
' DESCRIÇÃO        : Recupera o nome da pasta padrão configurada na ferramenta
'---------------------------------------------------------------------------------------
Function getCodeDbFolderName()
    Dim DBFolderName As String

    '## MZTools Custom - Variáveis de Ambiente
    '----------------------------------------------------------------------------------------------------
10  On Error GoTo getCodeDbFolderName_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
20  Dim dtSartRunProc As Date: dtSartRunProc = Time
    Const cstr_ProcedureName As String = "Function prjDRE.mdlDBConfig.getCodeDbFolderName()"
30  Call pegaValor("proceName", cstr_ProcedureName)
    'Habilitar Log de inicio/fim da rotina (-1 (ou True) / 0 ou False) ??
    Const cEnableLog As Boolean = 0

    '----------------------------------------------------------------------------------------------------

50  DBFolderName = getDefDBFolderName

60  If DBFolderName = "" Then
70      DBFolderName = VBA.InputBox("Necessário informar o nome de uma pasta padrão onde ficarão os bancos de dados" & VBA.vbNewLine & VBA.vbNewLine & _
                                "Caso não seja informado, o nome padrão 'BDs' será adotado. Caso não exista, a pasta será criada na raiz da aplicação.", "Pasta Padrão do Para o Back-End", "BDs")
80      Call pegaValor("DBConfig_DBFolderName", DBFolderName)
90  End If

100 getCodeDbFolderName = DBFolderName

110 On Error GoTo 0

130 Exit Function

getCodeDbFolderName_Error:
140 If VBA.Err <> 0 Then
150     lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
160     Call Excecoes.TratarErro(strErrorMessagem, lngErrorNumber, cstr_ProcedureName, VBA.Erl)
170 End If
180 Exit Function
    'Debug Mode
190 Resume
End Function

Function getDBPath() As String
    getDBPath = pegaValor("DB_PATH")
End Function

Function getDefDBFolderName()
    On Error Resume Next
    getDefDBFolderName = pegaValor("DBConfig_DBFolderName")
End Function

'Remove as tabelas da tabela 'tblDBConfig' que estao com link quebrado
Sub RemoveTabelasLinkQuebrado()
    Dim rsTables As Object
    Dim tbDef  As Object
10  Set rsTables = CodeDb.OpenRecordset("Pegar_DBConfig")

20  Do While Not rsTables.EOF
30      Set tbDef = Nothing
40      On Error Resume Next
50      Set tbDef = CodeDb.TableDefs(rsTables!tblname.value)
60      If VBA.Err <> 0 Then
70          Debug.Print rsTables!tblname.value, "Tabela nao existe, pode ser removida do de-para"
80          CodeDb.Execute "DELETE FROM  tblDBConfig  WHERE tblName = '" & rsTables!tblname.value & "'"
90      Else
100         If CodeDb.TableDefs(rsTables!tblname.value).Attributes = dbAttachedTable Then
110             On Error Resume Next
120             CodeDb.TableDefs(rsTables!tblname.value).RefreshLink
130             If VBA.Err <> 0 Then
140                 Debug.Print rsTables!tblname.value, "Erro na atualização...", VBA.Err.Description
150                 CodeDb.Execute "DELETE FROM  tblDBConfig  WHERE tblName = '" & rsTables!tblname.value & "'"
                Else
                    Debug.Print rsTables!tblname.value, "OK, Aualizado..."
160             End If
170         Else
                Debug.Print "#### Não ta Tabela vinculada", rsTables!tblname.value
180         End If
190     End If
200     rsTables.MoveNext
210 Loop
End Sub

Sub ExcluiTabelasVinculadasNaoMapeadas()
    Dim tb As Object 'TableDef
    For Each tb In CodeDb.TableDefs
        If tb.Attributes = dbAttachedTable Then
            If VBA.IsNull(DLookup("tblName", "Pegar_DBConfig", "tblName = '" & tb.Name & "'")) Then
                Debug.Print tb.Name, tb.Connect
                CodeDb.Execute "DROP TABLE [" & tb.Name & "]"
            End If
        End If
    Next
End Sub


Sub AdicionaTabelasNaoMapeadas()
    Dim tb As Object 'TableDef
    For Each tb In CodeDb.TableDefs
        If tb.Attributes = dbAttachedTable Then
            If VBA.IsNull(DLookup("tblName", "Pegar_DBConfig", "tblName = '" & tb.Name & "'")) Then
            CodeDb.Execute "INSERT INTO tblDBConfig (dbFileName, tblName, dbFolder, lastDtUpdate, Status)" & _
            "VALUES ('" & VBA.Dir(VBA.Split(tb.Connect, "=")(1)) & "', '" & tb.Name & "', '" & AuxFileSystem.getFSOFile(CStr(VBA.Split(tb.Connect, "=")(1))).ParentFolder.Name & "', #" & VBA.Now & "#, 'adicionado automaticamente')"
            End If
        End If
    Next
End Sub

Sub IdentificarQueriesComErros()
    Dim qdf As Object 'QueryDef
    Dim lngCount As Long
    Dim db     As Object
10  Debug.Print "Total", CodeDb.QueryDefs.count
    ' DBEngine.BeginTrans
20  Set db = DBEngine.OpenDatabase(CodeDb.Name, , True)

30  For Each qdf In db.QueryDefs
40      On Error Resume Next
50      If Not qdf.Name Like "*_REMOVER" Then

            'Debug.Print "Ultima", qdf.Name
60          qdf.OpenRecordset
70          If VBA.Err <> 0 Then
80              If VBA.Err.Number = 3078 Then
                    'Debug.Print qdf.Name, VBA.Err.Number, VBA.Err.description
90                  lngCount = lngCount + 1
100                 qdf.Name = VBA.Replace(qdf.Name, "_REMOVER", "") & "_REMOVER"
'110                 Stop
120             End If
130         Else
140             qdf.Close
150         End If

160     End If
170 Next
180 Debug.Print "Total de Queries com Erro : ", lngCount
End Sub

Function MostrarAvisoInconsistenciaLinks(Optional pMostraTelaGerenciadorBackEnd As Boolean, Optional pComplementoMensagem As String = "")
    Call MessageBoxMaster("F020", pComplementoMensagem)
    If pMostraTelaGerenciadorBackEnd Then Call AbrirCaminhoDatabase
End Function

Sub OcultarTodosObjetos()
Call ExibiOcultaObjetos(CurrentData.AllTables, True)
Call ExibiOcultaObjetos(CurrentData.AllQueries, True)
Call ExibiOcultaObjetos(CurrentProject.AllModules, True)
Call ExibiOcultaObjetos(CurrentProject.AllForms, True)
Call ExibiOcultaObjetos(CurrentProject.AllMacros, True)
Call ExibiOcultaObjetos(CurrentProject.AllReports, True)
End Sub

Sub ExibiOcultaObjetos(colObjetos As Object, Optional bOcultar As Boolean = False)
    Dim objAcc As AccessObject
    For Each objAcc In colObjetos
        If (Not objAcc.Name Like "BL_*") And (Not objAcc.Name Like "MSys*") Then
        
            Call Access.Application.SetHiddenAttribute(objAcc.Type, objAcc.Name, bOcultar)
        End If
    Next
End Sub

Function getSQLFromQuery(pQueryName As String, ParamArray arrParametros() As Variant) As String
    On Error GoTo ErroQueryDef
    Dim intContador As Integer
    Dim strComando As String
    strComando = CurrentDb.QueryDefs(pQueryName).SQL
    For intContador = 0 To UBound(arrParametros) Step 2
        strComando = VBA.Replace(strComando, arrParametros(intContador), arrParametros(intContador + 1))
    Next intContador
    getSQLFromQuery = strComando
ErroQueryDef:
    If VBA.Err = 3265 Then
        getSQLFromQuery = "QueryDef.SQL(Query não existe)"
    ElseIf VBA.Err <> 0 Then
        getSQLFromQuery = "QueryDef.SQL(Erro desconhecido) > " & VBA.Err.Description
    End If
End Function


'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : AuxForm.DefinirPropriedadeObjeto()
' TIPO             : Sub
' DATA/HORA        : 03/04/2014 16:02
' CONSULTOR        : (ADELSON)/ TECNUN - Adelson Rosendo Marques da Silva
' CONTATO          : adelson@tecnun.com.br
' DESCRIÇÃO        : Configura uma propriedade generica em um objeto (tabela, formulário queries, campos, etc)
'---------------------------------------------------------------------------------------
'Exemplo de Propriedades comuns
'-------------------------------------------------------------------------------------
'Format         : Define a formatação de um campo.
'                 Por ex 'Standard' no argumento strValue formata o valor numerico para o formato interno 'Padrão'
'Description    : Define a descrição de um objeto. Isso aparece na caixa descrição quando clicando com o botão direito do mouse no item 'Propriedades'
'TotalsRow      : Tabela/Consulta (Boolean) - Permite visualiza a linha de totais em queries e tabelas no modo de visualização de folha de dados
'AggregateType  : Campo - (Inteiro) Define o tipo de função que será usada na linha de totais
'                  -1 - Não mostra nehhuma função
'                   0 - Função de Soma
'                   1 - Media
'                  As demais funções seguem a mesma ordem conforme a visualização do access
'Há outras propriedades possivelmente existentes porem nao documentadas
'Adelson : 18/07/2016 - Alterado o objeto de manipulação das propriedade.
'

Function DefinirPropriedadeObjeto(objObject, _
                                  strNomePropriedade As String, _
                                  strValue As String, _
                                  Optional tipoValue As Integer = 1, _
                                  Optional bPrint As Boolean = False) As Variant
    Dim p As Object    ' DAO.Property
    Dim arrResult
1   On Error GoTo DefinirPropriedadeObjeto_Error


    Dim objType As String
    Dim db As Object
    Dim doc As Object ' DAO.Document
2   Set db = CurrentDb
    Dim prp As Object 'Property
    Dim ctrLoop As Object 'Container
    Dim sNomeObjeto As String
    
    'Se for um objeto do Acces, tenta localizar o objeto pelo nome
3   If Not VBA.IsObject(objObject) Then
4       sNomeObjeto = CStr(objObject)
5       objType = CLng(Nz(DLookup("Type", "MSysObjects", "Name='" & CStr(objObject) & "'")))
6       Select Case objType
        Case -32768: strContainerName = "Forms": GoSub PegarObjeto
7       Case 1, 5, 6: strContainerName = "Tables": GoSub PegarObjeto
8       Case -32764: strContainerName = "Reports": GoSub PegarObjeto
9       Case -32761: strContainerName = "Modules": GoSub PegarObjeto
10      End Select
12  Else
        'Se nao, doc apontará para o objeto informado
        If VBA.TypeName(objObject) = "Database" Then
            Set doc = CurrentDb
        Else
            Set doc = objObject
13          sNomeObjeto = objObject.Name
        End If
14  End If

15  On Error Resume Next
16  Set p = doc.Properties(strNomePropriedade)
17  On Error Resume Next

18  If p Is Nothing Then
22      Set p = doc.CreateProperty(strNomePropriedade, tipoValue)
23      p.value = strValue
24      doc.Properties.Append p
25      doc.Properties.Refresh
27      If bPrint Then Debug.Print "Nova propriedade '" & strNomePropriedade & "' adicionada em '" & doc.Name & "' (" & VBA.TypeName(doc) & ")"
28      arrResult = Array(-1, "Nove propriedade foi adiconada")
29  Else
30      p.value = strValue
31      doc.Properties.Refresh
32      If bPrint Then Debug.Print "A propriedade '" & strNomePropriedade & "'  em '" & doc.Name & "' (" & VBA.TypeName(doc) & ")'  ja existe..."
33      arrResult = Array(-1, "Propriedade existente foi atualizada")
34  End If

Fim:
35  DefinirPropriedadeObjeto = arrResult
    Call RemoverObjetosMemoria(p, doc, db, objObject)

36  On Error GoTo 0
37  Exit Function

PegarObjeto:
38  For Each ctrLoop In db.Containers
39      If ctrLoop.Name = strContainerName Then
40          For Each doc In ctrLoop.Documents
41              If doc.Name = sNomeObjeto Then
42                  Return
43              End If
44          Next doc
45      End If
46  Next ctrLoop

DefinirPropriedadeObjeto_Error:
48  If VBA.Err <> 0 Then
49      Call auxAuditoria.RegistraLog("AuxForm.DefinirPropriedadeObjeto() - Linha : " & VBA.Erl & " / ERROR : " & VBA.Err.Number & "-" & VBA.Err.Description, CurrentProject.Path & "\Log\VBAErros.txt", , , bPrint)
50      arrResult = Array(0, "Erro na configuração da propriedade. Motivo : " & VBA.Err.Description)
51      GoTo Fim
52  End If

53  Exit Function
54  Resume
End Function

Function PegarContainerPorNome(sNomeObjeto As String) As Object
    Dim objType As String
    Dim obj As Object ' dao.Document 'DAO.Container
    Dim db As Object
    Dim doc As Object ' dao.Document
    Set db = CurrentDb
    Dim prp As Object 'Property
    Dim ctrLoop As Object 'Container
    
1   objType = CLng(Nz(DLookup("Type", "MSysObjects", "Name='" & sNomeObjeto & "'")))
2   Select Case objType
    Case -32768: strContainerName = "Forms": GoSub PegarObjeto
4   Case 5: strContainerName = "Tables": GoSub PegarObjeto
5       'Set obj = CurrentDb.QueryDefs(sNomeObjeto)
'        Set obj = CurrentDb.Containers("Tables").Documents(sNomeObjeto)
6   Case 1, 6: strContainerName = "Tables": GoSub PegarObjeto
'7       Set obj = CurrentDb.Containers("Tables").Documents(sNomeObjeto)
8   Case -32764: strContainerName = "Reports": GoSub PegarObjeto
'9       Set obj = CurrentDb.Containers("Reports").Documents(sNomeObjeto)
10  Case -32761: strContainerName = "Modules": GoSub PegarObjeto
'11      Set obj = CurrentDb.Containers("Modules").Documents(sNomeObjeto)
12  End Select

    Set PegarContainerPorNome = doc
    
    Exit Function

PegarObjeto:



    For Each ctrLoop In db.Containers
            If ctrLoop.Name = strContainerName Then
            For Each doc In ctrLoop.Documents
                If doc.Name = sNomeObjeto Then
                    'Set obj = doc
                    Return
                    'For Each prp In doc.Properties
                        'If prp.name = "Description" Then
                         '  Debug.Print " - " & doc.name & " : " & prp.Value
                        'End If
                    'Next prp
                End If
            Next doc
            End If
    Next ctrLoop
    
    For iobj = 1 To CurrentDb.Containers(strContainerName).Documents.count
        If CurrentDb.Containers(strContainerName).Documents(iobj).Name = sNomeObjeto Then Set obj = CurrentDb.Containers(strContainerName).Documents(iobj): Stop: Exit For
    Next iobj
    
    Return

End Function

Function PegarPropriedadeObjeto(objObject, Optional strNomePropriedade As String = "Description")
    Dim p As Object
    Dim strContainerName As String
    Dim objType

3   If Not VBA.IsObject(objObject) Then
4       sNomeObjeto = CStr(objObject)
5       objType = CLng(Nz(DLookup("Type", "MSysObjects", "Name='" & CStr(objObject) & "'")))
6       Select Case objType
        Case -32768: strContainerName = "Forms": GoSub PegarObjeto
7       Case 1, 5, 6: strContainerName = "Tables": GoSub PegarObjeto
8       Case -32764: strContainerName = "Reports": GoSub PegarObjeto
9       Case -32761: strContainerName = "Modules": GoSub PegarObjeto
10      End Select
12  Else
        Set doc = objObject
13      sNomeObjeto = objObject.Name
14  End If

15  On Error Resume Next
16  Set p = doc.Properties(strNomePropriedade)
17  On Error Resume Next

18  If Not p Is Nothing Then PegarPropriedadeObjeto = p.value

    Exit Function

PegarObjeto:
38  For Each ctrLoop In CurrentDb.Containers
39      If ctrLoop.Name = strContainerName Then
40          For Each doc In ctrLoop.Documents
41              If doc.Name = sNomeObjeto Then
42                  Return
43              End If
44          Next doc
45      End If
46  Next ctrLoop

End Function


Function DeletarPorData()
    Dim Data As Variant
    Dim dbArq As Object
    Dim arquivos As Object
'    data = VBA.InputBox("Informar uma DataRef")
    If VBA.MsgBox("Limpar bases agora ?" & VBA.vbNewLine & VBA.vbNewLine & _
              "Obs. A ultima data será mantida", VBA.vbYesNo + VBA.vbQuestion) = VBA.vbYes Then
        With VBA.CreateObject("Scripting.FileSystemObject")
            Set arquivos = .GetFolder("D:\Data\Dev\Projetos\Tecnun\Bradesco\PRIVATE\Ferramenta_Private\Homologacao v3.22.0\BDs - Cópia")
            For Each dbArq In arquivos.Files
                Set db = DBEngine.OpenDatabase(dbArq.Path)
                Call DeleteAllTables(db, ata)
                db.Close
                Set db = Nothing
            Next dbArq
        End With
        VBA.MsgBox "Concluído ! Compactar o BD", VBA.vbInformation

   End If
End Function

Function DeleteAllTables(Optional db As Object, Optional pData)
    Dim Data As Date
    If Not VBA.IsMissing(pData) Then Data = CDate(pData)
    
    With db.OpenRecordset("SELECT Name AS tblName FROM (" & SQLTabelas & ") WHERE categoria='Usuario'")
        Do While Not .EOF
            If TemCampo(!tblname.value, "DataRef", db) And Not VBA.IsMissing(pData) Then
                Data = PegarUltimaDataRef(!tblname.value, db)
                strWhereDataRef = " WHERE NOT DataRef = #" & VBA.Format(Data, "mm/dd/yyyy") & "#"
                Debug.Print VBA.Now, "Limpando tabela : [" & !tblname.value & "] em : '" & VBA.Dir(db.Name) & "'...Mantem DataRef : " & Data
                Call db.Execute("DELETE FROM [" & !tblname.value & "]" & strWhereDataRef)
            Else
                Debug.Print VBA.Now, "NÃO TEM DATAREF : Tabela : [" & !tblname.value & "] em : '" & VBA.Dir(db.Name) & "'"
                strWhereDataRef = ""
            End If
            .MoveNext
        Loop
    End With
End Function

Function PegarUltimaDataRef(strTabela As String, Optional db As Object)
'    On Error Resume Next
    If db Is Nothing Then Set db = CurrentDb
    PegarUltimaDataRef = Nz(db.OpenRecordset("SELECT MAX(DataRef) FROM [" & strTabela & "]")(0).value)
End Function

Function SQLTabelas()
    SQLTabelas = "SELECT o.Name,VBA.vba.iif(o.Type=1,'Tabela Local','Tabela Vinculada') AS tipo,VBA.vba.iif(o.Type=1,vba.iif(o.name Like 'MSys*','Sistema',vba.iif(o.name Like 'f_*','Interna',vba.iif(o.name Like '*{*' And o.name Like '*}*' Or o.name Like '*~TMPCLP*','Temporária','Usuario'))),vba.iif(o.name Like 'f_*','Interna','Vinculada')) AS categoria,VBA.vba.iif(o.Type=6,nz('CONNECT : ' & o.Connect & '| LOCATION : ','') & o.Database,'(Database Local)') AS conexaoExterna" & _
                  " FROM MSysObjects AS o" & _
                  " WHERE o.Type In (1,6)"
End Function

Function TemCampo(pTabela As String, pNome As String, Optional db As Object) As Boolean
    On Error Resume Next
    If db Is Nothing Then Set db = CurrentDb
    TemCampo = db.OpenRecordset(pTabela).Fields(pNome).value <> ""
End Function

Public Sub SetStartupOptions(PropName As String, propType As Variant, propValue As Variant)
On Error GoTo Err_SetStartupOptions
    Dim ErrMsg As String, ErrChoice, ErrAns
  'Set passed startup property.

  'some of the startup properties you can use...
  ' "StartupShowDBWindow", DB_BOOLEAN, False
  ' "StartupShowStatusBar", DB_BOOLEAN, False
  ' "AllowBuiltinToolbars", DB_BOOLEAN, False
  ' "AllowFullMenus", DB_BOOLEAN, False
  ' "AllowBreakIntoCode", DB_BOOLEAN, False
  ' "AllowSpecialKeys", DB_BOOLEAN, False
  ' "AllowBypassKey", DB_BOOLEAN, False

  Dim dbs As Database

  Dim prp As Object

  Set dbs = CurrentDb

    If PropName = "ModoProducao" Then
        dbs.Properties("ShowDocumentTabs") = propValue
        dbs.Properties("AllowBreakIntoCode") = propValue
        dbs.Properties("AllowSpecialKeys") = propValue
        dbs.Properties("AllowBypassKey") = propValue
        dbs.Properties("AllowFullMenus") = propValue
        dbs.Properties("StartUpShowDBWindow") = propValue
    Else
        dbs.Properties(PropName) = propValue
    End If

  Set dbs = Nothing
  Set prp = Nothing

Exit_SetStartupOptions:
    Exit Sub

Err_SetStartupOptions:
    Select Case VBA.Err.Number
        Case 3270
           Set prp = dbs.CreateProperty(PropName, propType, propValue)
            Resume Next
        Case Else
          ErrAns = MsgBox(ErrMsg, _
          vbCritical + vbQuestion + ErrChoice, "SetStartupOptions")
          If ErrAns = vbYes Then
              Resume Next
          ElseIf ErrAns = vbCancel Then
              On Error GoTo 0
              Resume
          Else
              Resume Exit_SetStartupOptions
          End If
    End Select
End Sub



Function ConsultaExiste(pNomeConsulta As String, Optional db As Object) As Boolean
   Call Inicializar_Globais(False)
   If db Is Nothing Then Set db = CurrentDb
   ConsultaExiste = Conexao.ObjetoExiste(db, acQuery, pNomeConsulta) And pNomeConsulta <> ""
End Function

'Salva valores em uma tabela qualquer passando um array de valores
'ColunasValores deve conter os pares de Coluna = Valor
Function GravarDadosNaTabela(tabelaDestino As String, _
                             strArquivoDados As String, _
                             ParamArray ColunasValores()) As Variant
    On Error GoTo Erro:
    Dim lngErrorNumber As Long, strErrorMessagem As String: Const cstr_ProcedureName As String = "salvarDados()"
    
    Dim iColuna As Integer, coluna
    Dim tdTabela As TableDef, fdColuna As Field
    Dim rs As Recordset
    Dim retID
    Dim valor
    Dim TIPO As dao.DataTypeEnum
    Dim dbDestino As Database
    
    If strArquivoDados <> VBA.vbNullString Then
        Set dbDestino = DBEngine.OpenDatabase(strArquivoDados)
    Else
        Set dbDestino = Application.CurrentDb
    End If
    
    arrColunas = ColunasValores(0)
    arrValor = ColunasValores(1)
    
    '------------------------------------------------------------------------------------
    ' TABELA/CAMPOS JA EXISTEM !
    '------------------------------------------------------------------------------------
    If Not TabelaExiste(tabelaDestino, dbDestino) Then
        Call AuxDataBase.CriarTabela(tabelaDestino, arrColunas, arrValor, dbDestino)
    Else
        Call AuxDataBase.RedefinirCampos(tabelaDestino, arrColunas, arrValor, dbDestino)
    End If

    If UBound(ColunasValores) = 2 Then arrChave = ColunasValores(2)
    
    retID = 0
    
    If VBA.IsArray(arrChave) Then
        Set rs = dbDestino.OpenRecordset("SELECT * FROM [" & tabelaDestino & "] WHERE Format(" & arrChave(0) & ",'@')='" & arrChave(1) & "'")
    Else
        Set rs = dbDestino.OpenRecordset("SELECT * FROM [" & tabelaDestino & "] WHERE 1<>1") 'Força a criação de uma nova linha
    End If
    
    With rs
        If Not .EOF Then .edit Else .addNew
        For iColuna = 0 To UBound(arrColunas)
            coluna = arrColunas(iColuna)
            If VBA.InStr(coluna, ";") > 0 Then coluna = VBA.Split(coluna, ";")(0)
            valor = arrValor(iColuna)
            Call salvarValor(.Fields(coluna), valor)
        Next iColuna
        On Error Resume Next
        Call salvarValor(.Fields("usuario"), ChaveUsuario())
        Call salvarValor(.Fields("data_hora"), VBA.Now)
        On Error GoTo Erro:
        'Retorna o ID
        retID = .Fields(0).value
        .Update
        .Close
    End With
Fim:
    GravarDadosNaTabela = retID
    Exit Function
Erro:
    If VBA.Err() <> 0 Then
        lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
        Debug.Print cstr_ProcedureName, VBA.Err.Number & "-" & VBA.Err.Description
        Call VBA.Err.Raise(VBA.Err.Number, _
                           cstr_ProcedureName, _
                           "Ocorreu um erro ao Salvar os dados na tabela '" & VBA.CStr(tabelaDestino) & "'" & VBA.vbNewLine & VBA.vbNewLine & _
                           "Erro interno > Linha (" & VBA.Erl() & ") - Mensagem : " & strErrorMessagem)
        GoTo Fim
    End If
    Resume
End Function

Function salvarValor(campo, valor)
On Error GoTo ErrorHandler
Dim TIPO As dao.DataTypeEnum
    valor = VBA.CVar(VBA.Trim(valor))
    TIPO = AuxDataBase.DeterminaTipoDadoPorValor(valor)
    Select Case TIPO
    Case dao.DataTypeEnum.dbNumeric, dao.DataTypeEnum.dbDouble, dao.DataTypeEnum.dbSingle, dao.DataTypeEnum.dbDecimal
        campo.value = VBA.CDbl(VBA.Replace(valor, ".", ","))
    Case dbDate
        campo.value = (valor)
    Case Else
        If valor = "" Then
            campo.value = Null
        Else
            campo.value = valor
        End If
    End Select
Fim:
Exit Function
ErrorHandler:
    If VBA.Err() <> 0 Then
        If VBA.Err = 9292 Then
            Debug.Print VBA.Err.Description
        Else
            Call TratarErro(VBA.Err.Description, VBA.Err.Number, cstr_ProcedureName)
        End If
    End If
GoTo Fim
Resume
End Function

Function DeterminaTipoDadoPorValor(valor) As dao.DataTypeEnum
    Dim TIPO As dao.DataTypeEnum
    'Padrão é texto
    TIPO = dao.DataTypeEnum.dbText
    If VBA.IsNumeric(valor) Then
        TIPO = dao.DataTypeEnum.dbDouble
    Else
        If VBA.Len(valor) = 10 Then
            If VBA.IsDate(valor) Then
                TIPO = dao.DataTypeEnum.dbDate
            End If
        End If
    End If
    DeterminaTipoDadoPorValor = TIPO
End Function

Function RedefinirCampos(tabelaDestino As String, arrCols, arrVals, Optional db As Database)
    Dim tdTabela As TableDef, fdColuna As Field, dxPK As index
    Dim coluna As String
    Dim colunaExiste As Boolean
    Dim colunaCheck As String
    
    If db Is Nothing Then Set db = CurrentDb
    
    For iColuna = 0 To UBound(arrCols)
        coluna = arrCols(iColuna)
        valor = arrVals(iColuna)
        If VBA.InStr(coluna, ";") > 0 Then
            coluna = VBA.Split(coluna, ";")(0)
            pk = VBA.UCase(VBA.Split(arrCols(iColuna), ";")(1)) = "PK"
        End If
        On Error Resume Next
        colunaCheck = ""
        colunaCheck = db.TableDefs(tabelaDestino).Fields(coluna).Name
        If colunaCheck = "" Then
            Set fdColuna = db.TableDefs(tabelaDestino).CreateField(VBA.CStr(coluna), AuxDataBase.DeterminaTipoDadoPorValor(valor))
            fdColuna.OrdinalPosition = iColuna
            Call db.TableDefs(tabelaDestino).Fields.Append(fdColuna)
            db.TableDefs(tabelaDestino).Fields.Refresh
        End If
    Next iColuna
End Function

Function CriarTabela(tabelaDestino As String, arrCols, arrVals, Optional db As Database)
    On Error GoTo Erro:
    Dim lngErrorNumber As Long, strErrorMessagem As String: Const cstr_ProcedureName As String = "CriarTabela()"
    
    Dim tdTabela As TableDef, fdColuna As Field, dxPK As index
    Dim TIPO As dao.DataTypeEnum
    Dim iColuna As Integer
    Dim pk As Boolean
    
    If db Is Nothing Then Set db = CurrentDb
    
    Set tdTabela = db.CreateTableDef(tabelaDestino)
    
    Set fdColuna = tdTabela.CreateField("RowID", dao.DataTypeEnum.dbLong)
    fdColuna.Attributes = dbAutoIncrField
'    Call AdicionarField(tdTabela, fdColuna)
    Call tdTabela.Fields.Append(fdColuna)
    tdTabela.Fields.Refresh
    
    
    For iColuna = 0 To UBound(arrCols)
        pk = False
        'Valor/Coluna
        coluna = arrCols(iColuna)
        
'        If coluna = "LndrRate" Then Stop
        
        If VBA.InStr(coluna, ";") > 0 Then
            coluna = VBA.Split(coluna, ";")(0)
            pk = VBA.UCase(VBA.Split(arrCols(iColuna), ";")(1)) = "PK"
        End If
        
        valor = arrVals(iColuna)
        'Determina o tipo de acordo com o valor
        TIPO = AuxDataBase.DeterminaTipoDadoPorValor(valor)
        
        'Cria a coluna
        Set fdColuna = tdTabela.CreateField(VBA.CStr(coluna), TIPO)
        'Adiciona na tabela
        Call tdTabela.Fields.Append(fdColuna)
        tdTabela.Fields.Refresh
        
        If pk Then
            'Para criar a chave primaria,
            '1 - Cria um indice na tabela
            '2 - Define o indice como Chave Primaria (indice.Primary=True)
            '3 - Adiciona uma coluna em um
           Set dxPK = tdTabela.CreateIndex("PK")
           dxPK.Primary = True
           Call dxPK.Fields.Append(dxPK.CreateField(fdColuna.Name, fdColuna.Type))
            'Adiciona na tabela
            Call tdTabela.Indexes.Append(dxPK)
            tdTabela.Indexes.Refresh
        End If
        
    Next iColuna
    
    Set fdColuna = tdTabela.CreateField("usuario", dbText, 50)
    'Call AdicionarField(tdTabela, fdColuna)
    Call tdTabela.Fields.Append(fdColuna)
    tdTabela.Fields.Refresh
    
    
    Set fdColuna = tdTabela.CreateField("data_hora", dbDate)
    fdColuna.DefaultValue = "=Now()"
    'Call AdicionarField(tdTabela, fdColuna)
    Call tdTabela.Fields.Append(fdColuna)
    tdTabela.Fields.Refresh
    
    
    Call db.TableDefs.Append(tdTabela)
    
Fim:
    CriarTabela = retID
    Exit Function
Erro:
    If VBA.Err() <> 0 Then
        lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
        Debug.Print cstr_ProcedureName, VBA.Err.Number & "-" & VBA.Err.Description
        Call VBA.Err.Raise(VBA.Err.Number, _
                            cstr_ProcedureName, _
                           cstr_ProcedureName & " > Ocorreu um erro na criação da tabela'" & VBA.CStr(tabelaDestino) & "'" & VBA.vbNewLine & VBA.vbNewLine & _
                           "Erro interno > Linha (" & VBA.Erl() & ") - Mensagem : " & strErrorMessagem)
        GoTo Fim
    End If
    Resume
    
End Function

Private Sub AdicionarField(Tabela As TableDef, fd As Field)
    Call Tabela.Fields.Append(fd)
    Tabela.Fields.Refresh
End Sub

Sub ConfigurarBotoesEdicao()
    Dim cb As Object
    Dim ct As Object
    Set cb = Application.CommandBars("Navigation Pane List Pop-up")
    
    Set ct = cb.FindControl(tag:="btnAbrirTabelaBE")
    If ct Is Nothing Then Set ct = cb.Controls.Add()
    ct.Caption = "Abrir no Back-End"
    ct.OnAction = "AbrirTabelaNoBackEnd"
    ct.tag = "btnAbrirTabelaBE"
    ct.FaceId = 10


    Set ct = cb.FindControl(tag:="btnEditarTabelaBE")
    If ct Is Nothing Then Set ct = cb.Controls.Add()
    ct.Caption = "Edição Rápida"
    ct.OnAction = "EditarTabelaNoBackEnd"
    ct.tag = "btnEditarTabelaBE"
    ct.FaceId = 10

End Sub

Sub AbrirTabelaNoBackEnd()
    Dim accApp As Access.Application
    Dim strArquivo As String
    Dim vDados
    
    vDados = AnalisarTabelaVinculada(Application.CurrentObjectName)
    If vDados(0) = 0 Then
        strArquivo = vDados(6)
        VBA.Shell "MSACCESS """ & strArquivo & """", vbNormalFocus
    End If
End Sub

Sub EditarTabelaNoBackEnd()
    Dim accApp As Access.Application
    Dim strArquivo As String
    Dim vDados
    Dim strNomeObjeto As String
    strNomeObjeto = Application.CurrentObjectName
    vDados = AnalisarTabelaVinculada(strNomeObjeto)
    If vDados(0) = 0 Then
        Call DoCmd.openForm("frmEdicaoRapida", acNormal, , , acFormEdit, acDialog, strNomeObjeto)
    End If
End Sub

Sub CadastrarTabelaBackEnd(pTabela As String, pArquivo As String, Optional pSubPasta As String = "Dados")
    Call Inicializar_Globais
    Call Conexao.InserirRegistros("Deleta_Tabela_BackEnd", pTabela, pArquivo, pSubPasta)
    Call Conexao.InserirRegistros("Insere_Tabela_BackEnd", pTabela, pArquivo, pSubPasta)
End Sub

Function AtualizarIcone(strTabela As String, filtro As String, icone)
    Dim rs As Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM [" & strTabela & "] WHERE " & filtro)
    Call AnexaArquivo("", "", CurrentProject.Path & "\Imagens\" & icone & ".ico", True, CurrentDb, rs!icone)
End Function

Function pegarIconePorCodigo(codigo As String)

End Function


'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : AuxExport.AnexaArquivo()
' TIPO             : Function
' DATA/HORA        : 11/02/2015 12:26
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Anexa um arquivo qualquer na tabela de anexos identificados pela template
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      11/02/2015 12:26    Adelson         Criação/Atualização do procedimento
'---------------------------------------------------------------------------------------
Function AnexaArquivo(strNomeTabela As String, _
                      strNomeTemplate As String, _
                      pFileName As String, _
                      Optional pClearBefore As Boolean = True, _
                      Optional db As dao.Database, _
                      Optional pFieldToAddAttachment As Object) As Boolean

    Dim rsTabela As Object    ' DAO.Recordset2
    Dim rsArquivo As dao.Recordset2
    Dim fldAnexo As dao.Field2
    Dim strCaminhoArq As String
    Dim strTempDir As String
    Dim btContaRegitros As Byte

    On Error GoTo AnexaArquivo_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
    Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Function AuxExport.AnexaArquivo()"
    '----------------------------------------------------------------------------------------------------

10  If db Is Nothing Then Set db = Access.CurrentDb

    If Not pFieldToAddAttachment Is Nothing Then
50      Set rsArquivo = pFieldToAddAttachment.value
    Else
        If Not TabelaExiste(strNomeTabela) Then
            MsgBox "Tabela '" & strNomeTabela & "' não localizada !", vbExclamation, "Erro"
            Exit Function
        End If
20      Set rsTabela = db.OpenRecordset("SELECT * FROM " & strNomeTabela & " WHERE NomeTemplate  = '" & strNomeTemplate & "'")
40      If rsTabela.EOF Then rsTabela.addNew Else rsTabela.edit
        Set rsArquivo = rsTabela.Fields("ArqAnexo").value
        rsTabela!NomeTemplate = strNomeTemplate
    End If

60  If pClearBefore Then
70      Do While Not rsArquivo.EOF
80          rsArquivo.Delete
90          rsArquivo.MoveNext
100     Loop
110 End If

    If FileExists(pFileName) Then
120 With rsArquivo
130     If .EOF Then .addNew Else .edit
160     Call .Fields("FileData").LoadFromFile(pFileName)
180     .Update
190     .Close
200 End With
    rsTabela.Update
    End If

230 AnexaArquivo = VBA.Err.Number = 0

Fim:
    On Error GoTo 0
    Exit Function

AnexaArquivo_Error:
    If Err <> 0 Then
        lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
        Call MsgBox(lngErrorNumber & " - " & strErrorMessagem & vbNewLine & vbNewLine & "Em : " & cstr_ProcedureName, vbCritical, "Erro")
    End If
    GoTo Fim:
    'Debug Mode
    Resume
End Function

'Extrai um anexo de uma tabela especifica para um local
Public Function ExtrairAnexo(ByVal strNomeTabela As String, _
                             ByVal strNomeTemplate As String, _
                             Optional ByVal strFileName As String = VBA.vbNullString, _
                             Optional ByVal strToDir As String = VBA.vbNullString, _
                             Optional bShowMsg As Boolean = True) As String
10  On Error GoTo TratarErro

    Dim rsTabela As Object
    Dim rsArquivo As Object
    Dim fldAnexo As Object 'dao.Field2
    Dim strCaminhoArq As String
    Dim strTempDir As String
    Dim btContaRegitros As Byte
    Dim db As Object

20  Set db = Access.CurrentDb
30  Set rsTabela = db.OpenRecordset(strNomeTabela)

40  If strToDir = VBA.vbNullString Then strToDir = VBA.Environ("Temp")
50  If strFileName = VBA.vbNullString Then strFileName = "~_" & AuxOSFileSystem.getNewGUID(15, True) & ".mytmp"

60  strTempDir = strToDir

    Set rsTabela = db.OpenRecordset("SELECT * FROM " & strNomeTabela & " WHERE NomeTemplate  = '" & strNomeTemplate & "'")

80  If rsTabela.EOF Then
90      If bShowMsg Then MessageBoxMaster "Template para anexo '" & strNomeTemplate & "' não foi localizado na tabela '" & strNomeTabela & "'", VBA.vbExclamation
100 Else
110     Set rsArquivo = rsTabela.Fields("ArqAnexo").value
        If rsArquivo.EOF Then
            'Não há arquivo anexo
            ExtrairAnexo = "(Não há anexo)"
        Else
120         If strFileName = "OriginalName" Then strFileName = rsArquivo!FileName.value
130         strCaminhoArq = strTempDir & "\" & strFileName
            'Exclui o arquivo, caso exista
140         If Not VBA.Dir(strCaminhoArq) = "" Then
150             VBA.SetAttr strCaminhoArq, VBA.vbNormal
160             On Error Resume Next
170             VBA.Kill strCaminhoArq
180         End If
            Set fldAnexo = rsArquivo.Fields("FileData")
            'Salva o arquivo no disco. No endereço informado
200         fldAnexo.SaveToFile strCaminhoArq
210         rsArquivo.Close
220         ExtrairAnexo = strCaminhoArq
        End If
230 End If

240 Call Publicas.RemoverObjetosMemoria(fldAnexo, rsTabela, rsArquivo, db)

250 Exit Function
TratarErro:
260 Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxExport.PegaCaminhoArquivoAnexo()", Erl, , False)
270 Exit Function
280 Resume
End Function


'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : AuxExport.AnexaArquivo()
' TIPO             : Function
' DATA/HORA        : 11/02/2015 12:26
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Anexa um arquivo qualquer na tabela de anexos identificados pela template
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      11/02/2015 12:26    Adelson         Criação/Atualização do procedimento
'---------------------------------------------------------------------------------------
Public Function AnexaArquivoApp(strNomeTabela As String, pFileName As String, Optional pFieldToAddAttachment As Object) As Boolean

    Dim rsTabela As Object    ' DAO.Recordset2
    Dim rsArquivo As dao.Recordset2
    Dim fldAnexo As dao.Field2
    Dim strCaminhoArq As String
    Dim strTempDir As String
    Dim btContaRegitros As Byte
    Dim db As Database

    On Error GoTo AnexaArquivo_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
    Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "AuxExport.AnexaArquivo()"
    '----------------------------------------------------------------------------------------------------

10  Set db = Access.CurrentDb

    If Not pFieldToAddAttachment Is Nothing Then
50      Set rsArquivo = pFieldToAddAttachment.value
    Else
        If Not TabelaExiste(strNomeTabela) Then
            MsgBox "Tabela '" & strNomeTabela & "' não localizada !", vbExclamation, "Erro"
            Exit Function
        End If
20      Set rsTabela = db.OpenRecordset("SELECT * FROM " & strNomeTabela)
40      If rsTabela.EOF Then rsTabela.addNew Else rsTabela.edit
        Set rsArquivo = rsTabela.Fields("ArqAnexo").value
    End If

70      Do While Not rsArquivo.EOF
            If VBA.LCase(pFileName) Like "*" & VBA.LCase(rsArquivo!FileName) Then
                rsArquivo.Delete
                Exit Do
            End If
90          rsArquivo.MoveNext
100     Loop

120 With rsArquivo
130     .addNew
160      Call .Fields("FileData").LoadFromFile(VBA.LCase(pFileName))
180     .Update
190     .Close
200 End With
    rsTabela.Update

230 AnexaArquivoApp = VBA.Err.Number = 0

Fim:
    On Error GoTo 0
    Exit Function

AnexaArquivo_Error:
    If Err <> 0 Then
        lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
        Call MsgBox(lngErrorNumber & " - " & strErrorMessagem & vbNewLine & vbNewLine & "Em : " & cstr_ProcedureName, vbCritical, "Erro")
    End If
    GoTo Fim:
    'Debug Mode
    Resume
End Function

'Extrai um anexo de uma tabela especifica para um local
Private Function ExtrairAnexoApp(ByVal strNomeTabela As String, _
                             ByVal strNomeTemplate As String, _
                             Optional ByVal strFileName As String = VBA.vbNullString, _
                             Optional ByVal strToDir As String = VBA.vbNullString, _
                             Optional bShowMsg As Boolean = True) As String
10  On Error GoTo TratarErro

    Dim rsTabela As Object
    Dim rsArquivo As Object
    Dim fldAnexo As Object 'dao.Field2
    Dim strCaminhoArq As String
    Dim strTempDir As String
    Dim btContaRegitros As Byte
    Dim db As Object

20  Set db = Access.CurrentDb
30  Set rsTabela = db.OpenRecordset(strNomeTabela)

40  If strToDir = VBA.vbNullString Then strToDir = VBA.Environ("Temp")
50  If strFileName = VBA.vbNullString Then strFileName = "~_" & VBA.Format(Now, "yyyymmddhhnnss") & ".mytmp"

60  strTempDir = strToDir

    Set rsTabela = db.OpenRecordset("SELECT * FROM " & strNomeTabela)

80  If rsTabela.EOF Then
90      If bShowMsg Then MessageBoxMaster "Template para anexo '" & strNomeTemplate & "' não foi localizado na tabela '" & strNomeTabela & "'", VBA.vbExclamation
100 Else
110     Set rsArquivo = rsTabela.Fields("ArqAnexo").value
        If rsArquivo.EOF Then
            'Não há arquivo anexo
            ExtrairAnexoApp = "(Não há anexo)"
        Else
120         If strFileName = "OriginalName" Then strFileName = rsArquivo!FileName.value
            If VBA.Dir(strTempDir, vbDirectory) = "" Then Call MkFullDirectory(strTempDir)

130         strCaminhoArq = strTempDir & "\" & strFileName
            'Exclui o arquivo, caso exista
140         If Not VBA.Dir(strCaminhoArq) = "" Then
150             VBA.SetAttr strCaminhoArq, VBA.vbNormal
160             On Error Resume Next
170             VBA.Kill strCaminhoArq
180         End If
            Set fldAnexo = rsArquivo.Fields("FileData")
            Do While Not rsArquivo.EOF
                   If VBA.LCase(strCaminhoArq) Like "*" & VBA.LCase(rsArquivo!FileName) Then
                       fldAnexo.SaveToFile strCaminhoArq
                       rsArquivo.Close
                       ExtrairAnexoApp = strCaminhoArq
                       Exit Do
                   End If
                 rsArquivo.MoveNext
            Loop
        End If
230 End If

240 Call RemoverObjetosMemoria(fldAnexo, rsTabela, rsArquivo, db)

250 Exit Function
TratarErro:
260 Call mTFW_Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "mTFW_Database.ExtrairAnexo()", Erl, , False)
270 Exit Function
280 Resume
End Function

