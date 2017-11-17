VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTFW_DBDrive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------------------------------------------
' Module    : DBDrive
' DateTime  : 19/10/2008 01:18
' Author    : Adelson
' Purpose   :
'---------------------------------------------------------------------------------------
Public Enum eTypeSource
    FileDB = 1
    ODBC = 2
End Enum

'Drivers de conexão ODBC
Private Const adoDriver_ODBCAccess As String = "Driver={Microsoft Access Driver (*.mdb)};"  'Access
Private Const adoDriver_ODBCdBase As String = "Driver={Microsoft dBase Driver (*.dbf)};"   'dBase
Private Const adoDriver_ODBcTFW_Excel As String = "Driver={Microsoft Excel Driver (*.xls)};"   'Excel
Private Const adoDriver_ODBCOracle As String = "Driver={Microsoft ODBC for Oracle};"        'Oracle
Private Const adoDriver_ODBCSQLServer As String = "Driver={SQL Serve};"                     'SQL Server
Private Const adoDriver_ODBCArqTexto As String = "Driver={Microsoft Text Driver (*.txt;*csv)};"    'Arquivo de Texto

'Banco de Dados, Username e Senha para drivers comuns
Private Const adoODBC_DataBase As String = "BDQ"    ' > Access, dBase, Excel
Private Const adoODBC_UserID As String = "UID"    '> Access, Oracle, SQL Server
Private Const adoODBC_Password As String = "PWD"    '> Access, Oracle, SQL Server

'Banco de Dados SQL
Private Const adoODBC_dbSQL As String = "Database"    '> SQL Server

'Driver ID
Private Const adoODBC_DriverIDExcel As String = "DriverID"    'Excel (790)
Private Const adoODBC_DriverIDOracle As String = "DriverID"    'Oracle (277)

'Diretório padrão
Private Const adoODBC_DefualtDir As String = "DefualtDir"    'Excel

'Servidor
Private Const adoODBC_Servidor As String = "Server"    'Oracle, SQL Server

'Propriedades para Arquivos de Texto
Private Const adoODBC_Texto_TipoArquivosSuportados As String = "Extensions"     'Tipos de arquivos de texto suportados (asc,csv,tab,txt)
Private Const adoODBC_Texto_PersisteSecurit As String = "Persist Security Info"    '(valores possiveis: (True ou False)

'OLEDB
Private Const adoDriver_OLEDB_Jet As String = "Microsoft.Jet.OLEDB.4.0"    'Access, Excel
Private Const adoDriver_OLEDB_Oracle As String = "msdaora"    'Oracle
Private Const adoDriver_OLEDB_SQLServer As String = "SQLOLEDB"    'Oracle

'ACE (Office 2007+)
Private Const adoDriver_OLEDB_ACE As String = "Microsoft.ACE.OLEDB.12.0"    'Access, Excel 2007 +

'Banco de Dados, Username e Senha para drivers comuns
Private Const adoOLEDB_DataBase As String = "Data Source"    '> Access, Oracle ,SQL Server, Excel
Private Const adoOLEDB_Password As String = "Password"    '> Access, SQL Server
Private Const adoOLEDB_UserID As String = "User ID"    '> Access, Oracle, SQL Server

'Propriedades especiais.
'Excel
Private Const adoOLEDB_Extenção As String = "Extended Properties"    '> Excel (Excel 8.0; HDR=Yes ou HDR=No)

Private Const adoExtPropertyExcel_XLS As String = "Excel 8.0"
Private Const adoExtPropertyExcel_XLSX As String = "Excel 12.0 XML"
Private Const adoExtPropertyExcel_XLSM As String = "Excel 12.0 Macro"
Private Const adoExtPropertyExcel_XLSB As String = "Excel 12.0"
Private Const adoExtPropertyTEXT As String = "text"

'SQL Server
Private Const adoOLEDB_NetWorkLib As String = "Network Library"    '> SQL Server (DBMSSOCN)
Private Const adoOLEDB_dbSQL As String = "Initial Catalog"    '> SQL Server
Private Const adoOLEDB_ConexãoSegura As String = "Trusted_Connection"    '> SQL Server (True ou False)

'Private WithEvents m_objDBConnection As ADODB.Connection
'Private WithEvents m_objRecordSetEx As ADODB.Recordset

Private m_objDBConnection As Object
Private m_objRecordSetEx As Object
'
Public Enum ADODB_CursorTypeEnum
    adOpenForwardOnly = 0
    adOpenKeyset = 1
    adOpenDynamic = 2
    adOpenStatic = 3
    adOpenUnspecified = -1
End Enum

Public Enum ADODB_CommandTypeEnum
    adCmdFile = 256
    adCmdStoredProc = 4
    adCmdTable = 2
    adCmdTableDirect = 512
    adCmdText = 1
    adCmdUnknown = 8
    adCmdUnspecified = -1
End Enum

Public Enum ADODB_LockTypeEnum
    adLockBatchOptimistic = 4
    adLockOptimistic = 3
    adLockPessimistic = 2
    adLockReadOnly = 1
    adLockUnspecified = -1
End Enum

Public Enum ADODB_ExecuteOptionEnum
    adOptionUnspecified = -1
    adAsyncExecute = 16
    adAsyncFetch = 32
    adAsyncFetchNonBlocking = 64
    adExecuteNoRecords = 128
    adExecuteStream = 1024
    adExecuteRecord = 2048
End Enum


Public Enum ADODB_ObjectStateEnum
    adStateClosed = 0
    adStateOpen = 1
    adStateConnecting = 2
    adStateExecuting = 4
    adStateFetching = 8
End Enum

Public Enum ADODB_PersistFormatEnum
    adPersistXML = 1
    adPersistADTG = 0
End Enum

Function getFSO() As Object
    Set getFSO = VBA.CreateObject("Scripting.FileSystemObject")
End Function

Private Sub Class_Terminate()
    Set m_objDBConnection = Nothing
End Sub
'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : ADODBDrive.OpenConnection()
' TIPO             : Function
' DATA/HORA        : 05/02/2015 17:32
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Rotina generica para a conexão com um banco de dados via ADO
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      05/02/2015 17:32    Adelson         Criação/Atualização do procedimento
'---------------------------------------------------------------------------------------
Public Function OpenConnection(strSource As String, Optional UserName As String = "", Optional password As String = "", Optional CurrentConn As Object) As Object
    Dim strConnecString As String

    On Error GoTo OpenConnection_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
    Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Function ADODBDrive.OpenConnection()"
    '----------------------------------------------------------------------------------------------------

    strConnecString = getConnectionStringJet(strSource)
    'Checa se a conexão ja esta aberta
    If ConnectionIsOpen(CurrentConn) Then Set OpenConnection = CurrentConn: Exit Function
10  Set m_objDBConnection = VBA.CreateObject("ADODB.Connection")

20  With m_objDBConnection
        If getTypeSource(strSource) = ODBC Then
            .Open ("DSN=" & strSource & ";")
        Else
            'Se o banco de dados estiver protegido e for informado a senha e o UserName
            If password <> "" Then
                'Senha
                .Properties("Jet OLEDB:Database Password") = password
                'Abre a conexão
            End If
            Call .Open(strConnecString)
        End If
40  End With

    Set OpenConnection = m_objDBConnection

Fim:
    On Error GoTo 0
    Exit Function

OpenConnection_Error:
    If VBA.Err <> 0 Then
        lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
        VBA.MsgBox "Error : " & lngErrorNumber & "-" & strErrorMessagem, VBA.vbCritical
    End If
    GoTo Fim:
    'Debug Mode
    Resume
End Function


'---------------------------------------------------------------------------------------
' PROCEDIMENTO : ConnectToSqlServer
' TIPO         : Function
' DATA - HORA  : 18/10/2008 01:10
' AUTOR        : Adelson Rosendo Marques da Silva
'                  Email:   adelson.silva@mondial.com.br
' COMENTÁRIOS  : Cria uma conexão com o SQL SERVER. Retorna um objeto Connection Aberto
'---------------------------------------------------------------------------------------
Public Function ConnectToSqlServer(strServidorSQL As String, strBancoDeDados As String, Optional UserName As String = "", Optional Senha As String, Optional UsarLogonWindows As Boolean = True, Optional CurrentConn As Object) As Object
    On Error GoTo Erro_SQL

    'Checa se a conexão ja esta aberta
    If ConnectionIsOpen(CurrentConn) Then Set ConnectToSqlServer = CurrentConn: Exit Function

    'Instacia o Objeto Connection
    'Set ConnectToSqlServer = New Object
10  Set ConnectToSqlServer = VBA.CreateObject("ADODB.Connection")
    Set m_objDBConnection = ConnectToSqlServer

20  With m_objDBConnection
        'Driver de conexão usado pelo SQL SERVER
        .Provider = adoDriver_OLEDB_SQLServer
        'Informa o servidor
        .Properties(adoOLEDB_DataBase) = strServidorSQL
        'Informa o banco de dados
        .Properties(adoOLEDB_dbSQL) = strBancoDeDados
        'Permite o login do windows
        If UsarLogonWindows Then
            'Indica que será utilizado o logon do windows
            .Properties("Integrated Security") = "SSPI"
        Else
            'UserName
            .Properties(adoOLEDB_UserID) = UserName
            'Senha
            .Properties(adoOLEDB_Password) = Senha
        End If
        'Desconsidera o tempo limite padrão de execução do comando
        .CommandTimeout = 0
        'Desconsidera o tempo limite de conexão
        .ConnectionTimeout = 0
        'Realiza a conexão
30      .Open
40  End With

    Exit Function
Erro_SQL:

    If VBA.MsgBox(Err & ":" & VBA.Error & VBA.Chr(13) & VBA.Err.source & VBA.Chr(13) & VBA.Chr(13) & "Depurar ?", VBA.vbCritical + VBA.vbYesNo) = VBA.vbYes Then
        Stop
        Resume
    End If

End Function
'---------------------------------------------------------------------------------------
' PROCEDIMENTO : ConnectToFirebird
' TIPO         : Function
' DATA - HORA  : 18/10/2008 01:17
' AUTOR        : Adelson Rosendo Marques da Silva
'                  Email:   adelson.silva@mondial.com.br
' COMENTÁRIOS  : Cria uma conexão com o Firebird
'---------------------------------------------------------------------------------------
'
Public Function ConnectToFirebird(strSource As String, Optional UserName As String = "SYSDBA", Optional Senha As String = "masterkey", Optional CurrentConn As Object) As Object
    On Error GoTo Erro_SQL

    'Checa se a conexão ja esta aberta
    If ConnectionIsOpen(CurrentConn) Then Set ConnectToFirebird = CurrentConn: Exit Function

    'Set ConnectToFirebird = New Object
10  Set ConnectToFirebird = VBA.CreateObject("ADODB.Connection")
    Set m_objDBConnection = ConnectToFirebird

20  With m_objDBConnection
        If getTypeSource(strSource) = ODBC Then
            .Open ("DSN=" & strSource & ";")
        ElseIf getTypeSource(strSource) = FileDB Then
            .ConnectionString = "Driver=Firebird/InterBase(r) driver;Dbname=" & strSource & ";MODE=READ;CHARSET=NONE;;UID=" & UserName & ";PWD=" & Senha & ""
            .Open
        Else
            VBA.MsgBox "Formato de conexão inválida." & VBA.vbNewLine & VBA.vbNewLine & "DB NAME : " & strSource, VBA.vbCritical, "Conexão Invalida"
        End If

        If .State = 1 Then
            ' Debug.Print "Conexão Aberta..." & ConnectToFirebird.ConnectionString
        Else
            Debug.Print "Conexão Não Abriu.." & VBA.Error
        End If

40  End With

    Exit Function
Erro_SQL:
    Call VBA.MsgBox(Err & ":" & VBA.Error & VBA.Chr(13) & VBA.Err.source & VBA.Chr(13) & VBA.Chr(13) & "Um erro ocorreu durante a conexão com o banco de dados", VBA.vbCritical)

    Exit Function
    Resume

End Function

Public Function ConnectExcel(pExcelFileDataBase As String, Optional CurrentConn As Object) As Object
    On Error GoTo Erro_SQL
    Dim cn As Object    ' ADODB.Connection
    'Checa se a conexão ja esta aberta
    If ConnectionIsOpen(CurrentConn) Then Set ConnectExcel = CurrentConn: Exit Function

    'Instacia o Objeto Connection
    'Set ConnectExcel = New Object
    If m_objDBConnection Is Nothing Then
10      Set m_objDBConnection = VBA.CreateObject("ADODB.Connection")
    End If

20  With m_objDBConnection
        Select Case UCase(getFSO.getExtensionFile(pExcelFileDataBase))
        Case "XLS"
            .Provider = adoDriver_OLEDB_Jet
            .Properties("Extended Properties").value = adoExtPropertyExcel_XLS & ";HDR=YES"

        Case "XLSX"
            .Provider = adoDriver_OLEDB_ACE
            .Properties("Extended Properties").value = adoExtPropertyExcel_XLSX & ";HDR=YES"

        Case "XLSM"
            .Provider = adoDriver_OLEDB_ACE
            .Properties("Extended Properties").value = adoExtPropertyExcel_XLSM & ";HDR=YES"

        Case "XLSB"
            .Provider = adoDriver_OLEDB_ACE
            .Properties("Extended Properties").value = adoExtPropertyExcel_XLSB & ";HDR=YES"

        End Select
        .Properties("Data Source") = pExcelFileDataBase
        .CommandTimeout = 0
30      .Open
40  End With

    Set ConnectExcel = m_objDBConnection

    Exit Function
Erro_SQL:

    If VBA.MsgBox(Err & ":" & VBA.Err.Description & VBA.Chr(13) & VBA.Err.source & VBA.Chr(13) & VBA.Chr(13) & "Depurar ?", VBA.vbCritical + VBA.vbYesNo) = VBA.vbYes Then
        Stop
        Resume
    End If

End Function


Public Function getConnectionStringJet(server_or_file As String, _
                                       Optional Database As String, _
                                       Optional ReadOnly As Boolean)
    Dim bServerIsFile As Boolean
    Dim sProvider As String
    Dim sExtProperties As String
    Dim pExtFile As String
    Dim pDataSource As String

    'Verifica se o servidor é um arquivo ou nome de uma instancia de um servidor (padrao sql)
    bServerIsFile = VBA.Dir(server_or_file) <> ""

    If bServerIsFile Then
        pExtFile = UCase(getExtensionFile(server_or_file))
        'Determina o provider (Driver)
        Select Case pExtFile
        Case "XLS", "MDB"    'Arquivos de dados Excel e Access 2007+
            sProvider = adoDriver_OLEDB_Jet
            pDataSource = VBA.Chr(34) & server_or_file & VBA.Chr(34)
            If Not pExtFile = "MDB" Then sExtProperties = ";" & adoExtPropertyExcel_XLS & ";HDR=YES"
        Case "XLSX", "XLSM", "XLSB", "ACCDB"    'Arquivos de dados Excel e Access 2007+
            sProvider = adoDriver_OLEDB_ACE
            pDataSource = VBA.Chr(34) & server_or_file & VBA.Chr(34)
            Select Case Right(pExtFile, 1)
            Case "XLSX": sExtProperties = ";" & adoExtPropertyExcel_XLSX & ";HDR=YES"
            Case "XLSM": sExtProperties = ";" & adoExtPropertyExcel_XLSM & ";HDR=YES"
            Case "XLSB": sExtProperties = ";" & adoExtPropertyExcel_XLSB & ";HDR=YES"
            End Select
        Case "TXT"
            sProvider = adoDriver_OLEDB_Jet
            pDataSource = VBA.Chr(34) & GetDirectoryPath(server_or_file) & VBA.Chr(34)
            sExtProperties = adoExtPropertyTEXT & ";HDR=YES"
        End Select
    Else

    End If

    If sExtProperties <> "" Then sExtProperties = ";Extended Properties=" & VBA.Chr(34) & sExtProperties & VBA.Chr(34)

    getConnectionStringJet = "Provider=" & sProvider & _
                             ";Data source=" & pDataSource & _
                             VBA.IIf(ReadOnly, ";Mode=Read", ";Mode=ReadWrite") & _
                             sExtProperties

End Function

Function getTypeSource(sSource As String) As eTypeSource
    If InStr(1, sSource, "\") > 0 Then
        getTypeSource = FileDB
    Else
        getTypeSource = ODBC
    End If
End Function


Function ConnectionIsOpen(CurrentConn As Object) As Boolean
'O Proximo bloco abaixo verifica se essa conexão ja ta aberta
    If Not VBA.IsMissing(CurrentConn) Then
        'Verifica se o objeto ta instanciado na memoria
        If Not CurrentConn Is Nothing Then
            'Se sim, verifica se a conexãoja ta aberta
            If CurrentConn.State = 1 Then
                ConnectionIsOpen = True
                Exit Function
            End If
        End If
    End If
    ConnectionIsOpen = False
End Function
'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : DBDrive.GetRecordSet()
' TIPO             : Function
' DATA/HORA        : 14/01/2015 00:15
' CONSULTOR        : (Adelson)/ TECNUN - Adelson Rosendo Marques da Silva
' CONTATO          : adelson@tecnun.com.br
' DESCRIÇÃO        : Recupera os dados em um recordset
'---------------------------------------------------------------------------------------
Function GetRecordSet(pSource As String, _
                      Optional pServerConnection As Object, _
                      Optional pDataSource)
    Dim cn As Object
    '----------------------------------------------------------------------------------------------------
10  On Error GoTo GetRecordSet_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
20  Dim dtSartRunProc As Date: dtSartRunProc = Time
    Dim varivaeis As Variant
    Const cstr_ProcedureName As String = "Function prjArmsFinanceManager.DBDrive.GetRecordSet()"
    '----------------------------------------------------------------------------------------------------

30  If Not pServerConnection Is Nothing Then
40      Set cn = pServerConnection
50  Else
60      'If CStr(pDataSource) = "" Then pDataSource = ThisWorkbook.FullName
70      Set cn = ConnectExcel(CStr(pDataSource))
80  End If

    Set m_objRecordSetEx = VBA.CreateObject("ADODB.Recordset")
100 Call m_objRecordSetEx.Open(pSource, cn, ADODB_CursorTypeEnum.adOpenStatic)
110 Set GetRecordSet = m_objRecordSetEx

120 On Error GoTo 0
130 Exit Function

GetRecordSet_Error:
140 If VBA.Err <> 0 Then
150     lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
160     varivaeis = Array(vbNullString)
        'Quando não há o modulo de exceções
170     Call VBA.MsgBox("Ocorreu um erro inesperado no processamento" & VBA.vbNewLine & VBA.vbNewLine & _
                        "Detalhes : " & cstr_ProcedureName & "." & VBA.Erl & " - " & lngErrorNumber & "-" & strErrorMessagem, VBA.vbCritical)
180 End If
190 Exit Function
    'Debug Mode
200 Resume
End Function

Public Function SaveRSToXML(ConnectionString As String, _
                            SQLString As String, FullPath As String) As Boolean

    Dim oCn As Object
    Dim oCmd As Object

    Set oCn = VBA.CreateObject("ADODB.Connection")
    Set oCmd = VBA.CreateObject("ADODB.Command")

    On Error GoTo ErrorHandler:

    oCn.ConnectionString = ConnectionString
    oCn.Open
    Set oCmd.ActiveConnection = oCn
    oCmd.CommandText = SQLString
    oCmd.CommandType = ADODB_CommandTypeEnum.adCmdText
    Set m_objRecordSetEx = oCmd.Execute
    m_objRecordSetEx.Save FullPath, ADODB_PersistFormatEnum.adPersistXML
    SaveRSToXML = True

ErrorHandler:
    On Error Resume Next
    Set m_objRecordSetEx = Nothing
    Set oCmd = Nothing
    If oCn.State <> 0 Then oCn.Close
    Set oCn = Nothing

End Function

Public Function LoadRsFromXML(FullPath As String) As Object    'ADODB.Recordset
    Set m_objRecordSetEx = VBA.CreateObject("ADODB.Recordset")
    On Error Resume Next
    If VBA.Dir(FullPath) = "" Then Exit Function
    'm_objRecordSetEx.Open FullPath, "Provider=MSPersist;", adOpenForwardOnly, adLockReadOnly, adCmdFile
    If VBA.Err.Number = 0 Then
        Set LoadRsFromXML = m_objRecordSetEx
    End If
End Function


Public Function RecordsetFromXMLString(sXML As String) As Object
    Dim oStream As Object    ' ADODB.stream
    Set oStream = VBA.CreateObject("ADODB.stream")

    With oStream
        Call .Open
        Call .WriteText(sXML)   'Give the XML string to the ADO Stream
        .position = 0    'Set the stream position to the start
    End With

    Set m_objRecordSetEx = VBA.CreateObject("ADODB.Recordset")

    Call m_objRecordSetEx.Open(oStream)     'Abre um recordset de um Stream (criado de um XML)
    oStream.Close
    Set oStream = Nothing
    Set RecordsetFromXMLString = m_objRecordSetEx
    Set m_objRecordSetEx = Nothing
End Function

Public Function RecordsetFromXMLDocument(pSource As Variant) As Object
    Dim xDoc As Object
    Set xDoc = VBA.CreateObject("MSXML.DOMDocument")
    If VBA.TypeName(pSource) Like "DOMDocument" Then
        Set xDoc = pSource
    ElseIf InStr(pSource, "<") > 0 And InStr(pSource, "/>") > 0 Then
        Call xDoc.LoadXML(pSource)
    ElseIf VBA.Dir(pSource) <> "" Then
        Call xDoc.Load(pSource)
    End If

    Set m_objRecordSetEx = VBA.CreateObject("ADODB.Recordset")

    m_objRecordSetEx.Open pSource    'pass the DOM Document instance as the Source argument

    Set RecordsetFromXMLDocument = m_objRecordSetEx  'return the recordset
    Set m_objRecordSetEx = Nothing

End Function

Public Property Get DBConnection() As Object    ' ADODB.Connection
    Set DBConnection = m_objDBConnection
End Property

Public Property Set DBConnection(objDBConnection As Object)
    Set m_objDBConnection = objDBConnection
End Property

'###### EVENTOS ####################################################################################
'
'Private Sub m_objDBConnection_ConnectComplete(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
'    Debug.Print VBA.Now, "m_objDBConnection_ConnectComplete() - " & pConnection.Properties("Data Source") & " is connected !"
'End Sub
'
'Private Sub m_objDBConnection_InfoMessage(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
'    Debug.Print VBA.Now, "m_objDBConnection_InfoMessage() - Connection info..." & adStatus
'End Sub
'
'Private Sub m_objRecordSetEx_EndOfRecordset(fMoreData As Boolean, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'    Debug.Print VBA.Now, "m_objRecordSetEx_EndOfRecordset()", "fMoreData = " & fMoreData
'End Sub
'
'Private Sub m_objRecordSetEx_FetchComplete(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'    Debug.Print VBA.Now, "m_objRecordSetEx_FetchComplete"
'End Sub
'
'Private Sub m_objRecordSetEx_FetchProgress(ByVal Progress As Long, ByVal MaxProgress As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'    Debug.Print VBA.Now, "m_objRecordSetEx_FetchProgress", "Progress = " & Progress, "MaxProgress = " & MaxProgress
'End Sub
'
'Private Sub m_objRecordSetEx_WillChangeField(ByVal cFields As Long, ByVal Fields As Variant, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'    Debug.Print VBA.Now, "m_objRecordSetEx_WillChangeField"
'End Sub


Function getExtensionFile(FileName As String) As String
    Dim fd As Object
    'Tenta primeiramente pelo Fso
    If FileExists(FileName) Then
        getExtensionFile = getFSO.GetExtensionName(FileName)
    End If
End Function

Function GetDirectoryPath(sFileName As String) As String
    On Error Resume Next
    Dim sTempFileName As String, fso As Object
    'Tenta primeiramente pelo Fso
    Set fso = getFSO
    'Verifica se o FSO foi instanciado
    If Not fso Is Nothing Then
        GetDirectoryPath = fso.GetParentFolderName(sFileName)
    Else
        'Se o Fso não estiver instalado, tenta a função Dir
        sTempFileName = VBA.Dir(sFileName)
        GetDirectoryPath = ""
    End If
    'Limpa a variável
    Set fso = Nothing
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : Form_frmLoginServer.TestConnection()
' TIPO             : Function
' DATA/HORA        : 29/09/2015 23:13
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Testa a conexão com um Servidor/BD SQL
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      29/09/2015 23:13
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function TestConnection(Optional pServerName As String, _
                                Optional pDBName As String, _
                                Optional pUserName As String, _
                                Optional pPassword As String, _
                                Optional pObjectStatus)
    '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
10  On Error GoTo Comando68_Click_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
20  Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Sub Form_frmLoginServer.TestConnection()"
    '----------------------------------------------------------------------------------------------------
    Dim strConn As String
    Dim objConnection As Object
    Dim modoAcesso As Integer
    
30  Set objConnection = VBA.CreateObject("ADODB.Connection")
    'Options.adAsyncConnect = 16
    'State.adStateConnecting = 2
40  Call objConnection.Open(ConnectionString:=BuildConnectionString(modoAcesso, , pServerName, pDBName, pUserName, pPassword), Options:=16)

    'Aguarda a finalização da conexão
50  On Error Resume Next
60  Access.Application.Echo 0
    Do Until objConnection.State <> 2:
70      VBA.DoEvents
80      If Not VBA.IsMissing(pObjectStatus) Then pObjectStatus = VBA.Now & " - Aguarde, Conectando...."
90  Loop
100 Access.Application.Echo 1

110 If VBA.Err.Number = 0 Then
120     TestConnection = Array(0, "Conexão efetuada com sucesso", "", VbMsgBoxStyle.vbInformation)
130 Else
140     TestConnection = Array(VBA.Err.Description, "Erro ao conectar ao banco de dados desejado", VBA.Err.Description, VbMsgBoxStyle.vbCritical)
150 End If

    'Finaliza a conexão
160 If Not objConnection Is Nothing Then
170     If objConnection.State <> 0 Then objConnection.Close
180     Set objConnection = Nothing
190 End If
Fim:
200 On Error GoTo 0
210 Exit Function

Comando68_Click_Error:
220 If VBA.Err <> 0 Then
230     lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
240     Call Excecoes.TratarErro(strErrorMessagem, lngErrorNumber, cstr_ProcedureName, VBA.Erl)
250 End If
    GoTo Fim:
    'Debug Mode
260 Resume

End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : BLSGPB.ConnectionString()
' TIPO             : Property
' DATA/HORA        : 29/09/2015 18:39
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Função generica que monta uma String de Conexão ADO
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      29/09/2015 18:39
'---------------------------------------------------------------------------------------
Private Property Get BuildConnectionString(Optional loginType As Integer = 1, _
                                                 Optional strAppSource As String, _
                                                 Optional strServerName As String, _
                                                 Optional strBDName As String, _
                                                 Optional m_sUserName As String, _
                                                 Optional m_sPassword As String) As String
    Dim strCString As String

    '----------------------------------------------------------------------------------------------------
10  On Error GoTo ConnectionString_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
20  Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Property BLSGPB.ConnectionString()"
    '----------------------------------------------------------------------------------------------------

30  If strAppSource = "" Then strAppSource = Environ("ComputerName") & "." & Environ("UserName") & "." & CurrentDb.Name & VBA.IIf(cstr_ProcedureName <> "", "." & cstr_ProcedureName, "")

40  strCString = "Provider=SQLOLEDB;"
50  strCString = strCString & VBA.vbNewLine & "Server=" & strServerName & ";"
60  strCString = strCString & VBA.vbNewLine & "Database=" & strBDName & ";"
70  strCString = strCString & VBA.vbNewLine & "Workstation ID=" & Environ("ComputerName") & ";"    'Nome da Maquina
80  strCString = strCString & VBA.vbNewLine & "Application Name=" & strAppSource & ";"

90  If loginType = 1 Then
        'Conexão Segura ou Integrada com o Windows
100     strCString = strCString & VBA.vbNewLine & "Integrated Security=SSPI;"
110 ElseIf loginType = 2 Then
        'Conexão nativa SQL Server
120     strCString = strCString & VBA.vbNewLine & "User Id=" & m_sUserName & ";"
130     strCString = strCString & VBA.vbNewLine & "Password=" & m_sPassword & ";"
140 End If

150 strCString = VBA.Replace(strCString, VBA.vbNewLine, "")
160 BuildConnectionString = strCString
Fim:
180 On Error GoTo 0
190 Exit Property

ConnectionString_Error:
200 If VBA.Err <> 0 Then
210     lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
220     Call VBA.MsgBox("Erro na construção da ConnectionString " & VBA.vbNewLine & VBA.vbNewLine & cstr_ProcedureName & ">" & strErrorMessagem & "(" & VBA.Erl & ")", VBA.vbCritical, "Erro")
230 End If
    GoTo Fim:
    'Debug Mode
240 Resume
End Property


