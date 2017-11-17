VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTFW_WebUtil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Enum eDownloadMethod
    API = 1
    HTTP_Request = 2
End Enum

Public Enum eHTTP_ReadyState
    HTTP_ReadyState_00_NaoInicializado = 0    '   Uninitialized - open() has not been called yet.
    HTTP_ReadyState_01_LOADING = 1    '  Loading - send() has not been called yet.
    HTTP_ReadyState_02_LOADED = 2    '  Loaded - send() has been called, headers and status are available.
    HTTP_ReadyState_03_DOWNLOAD = 3    '  Interactive - Downloading, responseText holds the partial data.
    HTTP_ReadyState_04_COMPLETE = 4    '  Completed - Finished with all operations.
End Enum


'Todos os Status HHTP possiveis
'Os demais serão 'Desconhecidos'
'Os mais comuns são : 200 - OK, 404 - Pagina não encontrada, 403 - Negado
Public Enum eHTTP_Status
    [HTTP_Status_100_Continue] = 100
    [HTTP_Status_101_Switching Protocols] = 101
    [HTTP_Status_102_Processing] = 102
    [HTTP_Status_200_OK] = 200
    [HTTP_Status_201_Created] = 201
    [HTTP_Status_202_Accepted] = 202
    [HTTP_Status_203_Non-Authoritative Information] = 203
    [HTTP_Status_204_No Content] = 204
    [HTTP_Status_205_Reset Content] = 205
    [HTTP_Status_206_Partial Content] = 206
    [HTTP_Status_207_Multi-Status] = 207
    [HTTP_Status_208_Already Reported] = 208
    [HTTP_Status_226_IM Used] = 226
    [HTTP_Status_300_Multiple Choices] = 300
    [HTTP_Status_301_Moved Permanently] = 301
    [HTTP_Status_302_Found] = 302
    [HTTP_Status_303_See Other] = 303
    [HTTP_Status_304_Not Modified] = 304
    [HTTP_Status_305_Use Proxy] = 305
    [HTTP_Status_306_(Unused)] = 306
    [HTTP_Status_307_Temporary Redirect] = 307
    [HTTP_Status_308_Permanent Redirect] = 308
    [HTTP_Status_400_Bad Request] = 400
    [HTTP_Status_401_Unauthorized] = 401
    [HTTP_Status_402_Payment Required] = 402
    [HTTP_Status_403_Forbidden] = 403
    [HTTP_Status_404_Not Found] = 404
    [HTTP_Status_405_Method Not Allowed] = 405
    [HTTP_Status_406_Not Acceptable] = 406
    [HTTP_Status_407_Proxy Authentication Required] = 407
    [HTTP_Status_408_Request Timeout] = 408
    [HTTP_Status_409_Conflict] = 409
    [HTTP_Status_410_Gone] = 410
    [HTTP_Status_411_Length Required] = 411
    [HTTP_Status_412_Precondition Failed] = 412
    [HTTP_Status_413_Payload Too Large] = 413
    [HTTP_Status_414_URI Too Long] = 414
    [HTTP_Status_415_Unsupported Media Type] = 415
    [HTTP_Status_416_Range Not Satisfiable] = 416
    [HTTP_Status_417_Expectation Failed] = 417
    [HTTP_Status_421_Misdirected Request] = 421
    [HTTP_Status_422_Unprocessable Entity] = 422
    [HTTP_Status_423_Locked] = 423
    [HTTP_Status_424_Failed Dependency] = 424
    [HTTP_Status_426_Upgrade Required] = 426
    [HTTP_Status_428_Precondition Required] = 428
    [HTTP_Status_429_Too Many Requests] = 429
    [HTTP_Status_431_Request Header Fields Too Large] = 431
    [HTTP_Status_451_Unavailable For Legal Reasons] = 451
    [HTTP_Status_500_Internal Server VBA.Error] = 500
    [HTTP_Status_501_Not Implemented] = 501
    [HTTP_Status_502_Bad Gateway] = 502
    [HTTP_Status_503_Service Unavailable] = 503
    [HTTP_Status_504_Gateway Timeout] = 504
    [HTTP_Status_505_HTTP Version Not Supported] = 505
    [HTTP_Status_506_Variant Also Negotiates] = 506
    [HTTP_Status_507_Insufficient Storage] = 507
    [HTTP_Status_508_Loop Detected] = 508
    [HTTP_Status_510_Not Extended] = 510
    [HTTP_Status_511_Network Authentication Required] = 511
End Enum


'Evento que registra as mensagems para serem recebida por clientes
Public Event OnRunning(pMessage As String)

Public Enum eHostType
    ieBrowser = 1
    HTTPRequest = 2
End Enum

Public Enum eResponseType
    ResponseResultBinary = 0
    ResponseResultXML = 1
    ResponseResultTEXT = 2
    ResponseResultXMLDOM = 3
End Enum

'Tipo de Cotação do Dolar
Public Enum eCotacaoDolarTipo
    Compra = 10813
    venda = 1
End Enum

'Codigo de Moedas mais utilizadas
Public Enum Moeda
    DOLAR_DOS_ESTADOS_UNIDOS = 61
    LIBRA_ESTERLINA = 115
    PESO_ARGENTINO = 156
    EURO = 222
    DOLAR_OURO = 236
End Enum

Private Type tpURL
    Scheme As String
    Host As String
    Port As Long
    URI As String
    Query As String
End Type

'---------------------------------------------------------------------------------------
' MÓDULO           : prjArmsFinanceManager.cBacen
' TIPO             : Módulo de classe
' DATA/HORA        : 13/11/2014 23:55
' CONSULTOR        : (ADELSON)/ TECNUN  - Adelson Rosendo Marques da Silva
' DESCRIÇÃO        : Contem funções que efetua consultas e baixa informações de cotação no site do Banco Central
'---------------------------------------------------------------------------------------
' + HISTÓRIO DE REVISÃO
'---------------------------------------------------------------------------------------
' DATA / DESCRIÇÃO
'---------------------------------------------------------------------------------------
' 13/11/2014 23:55 -
'---------------------------------------------------------------------------------------
Option Explicit


'OFFICE 64BITS
#If VBA7 Then
    Private m_sURL_WEBSERVICE_BACEN As String
    'API's que auxiliam algumas ações
    Private Declare PtrSafe Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
    Private Declare PtrSafe Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumberOfBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
    Private Declare PtrSafe Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hInternet&, ByVal lpszUrl As String, ByVal lpszHeaders As String, ByVal dwHeadersLength&, ByVal dwFlags&, ByVal dwContext&) As Long
    Private Declare PtrSafe Function InternetCloseHandle Lib "wininet" (ByRef hInet As Long) As Long
    Private Declare PtrSafe Function InternetGetCookie Lib "wininet.dll" Alias "InternetGetCookieA" (ByVal lpszUrlName As String, ByVal lpszCookieName As String, ByVal lpszCookieData As String, lpdwSize As Long) As Boolean
    Private Declare PtrSafe Function InternetGetConnectedState Lib "wininet" (ByRef dwFlags As Long, ByVal dwReserved As Long) As Long
#Else
    'OFFICE 32BITS
    Private Declare Function InternetGetCookie Lib "wininet.dll" Alias "InternetGetCookieA" (ByVal lpszUrlName As String, ByVal lpszCookieName As String, ByVal lpszCookieData As String, lpdwSize As Long) As Boolean
    Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
    Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
    Private Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumberOfBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
    Private Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hInternet&, ByVal lpszUrl As String, ByVal lpszHeaders As String, ByVal dwHeadersLength&, ByVal dwFlags&, ByVal dwContext&) As Long
    Private Declare Function InternetGetConnectedState Lib "wininet" (ByRef dwFlags As Long, ByVal dwReserved As Long) As Long
#End If

'Tipo de Conexão
Private Const CONNECT_LAN As Long = &H2
Private Const CONNECT_MODEM As Long = &H1
Private Const CONNECT_PROXY As Long = &H4
Private Const CONNECT_OFFLINE As Long = &H20
Private Const CONNECT_CONFIGURED As Long = &H40

Private Const MAX_BUFFER_LENGTH = 8162
Private Const API_AGENT_NAME As String = "Web Drive Library"
Private Const INTERNET_OPEN_TYPE_DIRECT = 1
Private Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000

Public Event DebugPrint(pMensagem As String)


Function OpenInternetExplorer(modeInstance As eGetInstanceType) As cIE_Application
    Dim cIE As cIE_Application
    Set cIE = New cIE_Application
    Call cIE.startNewIE("")
    Set OpenInternetExplorer = cIE
End Function
'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : mdlBacen.openDocument()
' TIPO             : Function
' DATA/HORA        : 13/11/2014 23:13
' CONSULTOR        : (ADELSON)/ TECNUN - Adelson Rosendo Marques da Silva
' CONTATO          : adelson@tecnun.com.br
' DESCRIÇÃO        : Instancia um objeto HTMLDocument a partir de uma URL usando HTTPRequest
'---------------------------------------------------------------------------------------
Public Function openDocument(pUrl As String) As Object
    Dim htmlDoc As Object
    Dim tr     As Object
    Dim xmlhttp As Object

    'Obtemos uma instância do objeto XMLHTTP
    'Com ele se acessa a página
    '----------------------------------------------------------------------------------------------------
    On Error GoTo openDocument_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
    Dim dtSartRunProc As Date: dtSartRunProc = Time
    Const cstr_ProcedureName As String = "Function prjArmsFinanceManager.mdlBacen.openDocument()"
    '----------------------------------------------------------------------------------------------------

    Set xmlhttp = VBA.CreateObject("MSXML2.ServerXMLHTTP")
    
    'Abrimos uma conexão com o servidor da URL
    xmlhttp.Open "GET", pUrl, False
    
    'Enviamos o pedido
    xmlhttp.send

    Set htmlDoc = VBA.CreateObject("HtmlFile")
    htmlDoc.Write xmlhttp.responseText
    'Obtemos a resposta com a página HTML da URL
    Set openDocument = htmlDoc
    'Finalizamos o objeto
    Set xmlhttp = Nothing

    On Error GoTo 0
    Exit Function

openDocument_Error:
    If VBA.Err.Number = -2147012889 Then
        lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.source & "-" & VBA.Err.Description

    ElseIf VBA.Err <> 0 Then
        lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
        'Quando não há o modulo de exceções
        Call VBA.MsgBox("Ocorreu um erro ao criar a nova tarefa" & VBA.vbNewLine & VBA.vbNewLine & _
                        "Detalhes : " & cstr_ProcedureName & "." & VBA.Erl & " - " & lngErrorNumber & "-" & strErrorMessagem)
    End If
    Exit Function
    'Debug Mode
    Resume
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cTFW_WebUtil.download()
' TIPO             : Function
' DATA/HORA        : 08/11/2014 15:39
' CONSULTOR        : (ADELSON)/ TECNUN - Adelson Rosendo Marques da Silva
' CONTATO          : adelson@tecnun.com.br
' DESCRIÇÃO        : Efetua o download de um arquivo na Internet
'---------------------------------------------------------------------------------------
Private Function downloadViaApi(ByVal sURL As String, Optional ByVal DEBUG_ON As Boolean, Optional destFilePath As String) As String
    Dim hOpen As Long, hFile As Long, ret As Long
    Dim sBuff As String * MAX_BUFFER_LENGTH, sData As String
    Dim sFileName As String, iFile As Integer, dData As Double

    '----------------------------------------------------------------------------------------------------
10  On Error GoTo Download_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
20  Dim dtSartRunProc As Date: dtSartRunProc = Time
    Const cstr_ProcedureName As String = "Function prjArmsFinanceManager.cTFW_WebUtil.download()"
    '----------------------------------------------------------------------------------------------------

30  If destFilePath = "" Then
40      If InStr(1, sURL, "/") Then
50          sFileName = Environ("Temp") & "\" & Right$(sURL, Len(sURL) - VBA.InStrRev(sURL, "/"))
60      Else
70          sFileName = Environ("Temp") & "\Download" & VBA.Format(VBA.Now, "hhmmssmmddyy") & ".tmp"
80      End If
90  Else
100     sFileName = destFilePath
110 End If

120 If DEBUG_ON Then Debug.Print "Conectando a ... " & sURL
130 hOpen = InternetOpen(API_AGENT_NAME, INTERNET_OPEN_TYPE_DIRECT, VBA.vbNullString, VBA.vbNullString, 0)
140 If hOpen = 0 Then
150     VBA.MsgBox "Error opening Internet connection"
160     downloadViaApi = False
170     Exit Function
180 End If
190 hFile = InternetOpenUrl(hOpen, sURL, VBA.vbNullString, 0, INTERNET_FLAG_NO_CACHE_WRITE, 0)
200 If hFile = 0 Then
210     dData = 0
220 Else
230     If DEBUG_ON Then Debug.Print "Downloading ... " & VBA.Format(Len(sData) / 1024, 0) & " KB"
240     InternetReadFile hFile, sBuff, MAX_BUFFER_LENGTH, ret
250     sData = sBuff
260     Do While ret <> 0
270         If DEBUG_ON Then Debug.Print "Downloading ... " & VBA.Format(Len(sData) / 1024, 0) & " KB"
280         InternetReadFile hFile, sBuff, MAX_BUFFER_LENGTH, ret
290         sData = sData + Mid(sBuff, 1, ret)
300     Loop
310     dData = Len(sData): iFile = VBA.FreeFile
320     Open sFileName For Binary Access Write Lock Write As #iFile
330     Put #iFile, , sData: Close #iFile
340 End If
350 InternetCloseHandle hFile
360 InternetCloseHandle hOpen
370 sData = ""


410 If VBA.Dir(sFileName) <> "" Then If DEBUG_ON Then Debug.Print "Download concluído : " & VBA.Format(VBA.FileLen(sFileName) / 1024, "0.0 KB") & " / Destino : " & sFileName

380 If dData Then downloadViaApi = sFileName

420 On Error GoTo 0
430 Exit Function

Download_Error:
440 If VBA.Err <> 0 Then
450     lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
        'Quando não há o modulo de exceções
460     Call VBA.MsgBox("Ocorreu um erro ao criar a nova tarefa" & VBA.vbNewLine & VBA.vbNewLine & _
                        "Detalhes : " & cstr_ProcedureName & "." & VBA.Erl & " - " & lngErrorNumber & "-" & strErrorMessagem)
470 End If
480 Exit Function
    'Debug Mode
490 Resume
End Function
'---------------------------------------------------------------------------------------

'Função escapsulada somente para Download via HTTP
Public Function HTTP_Download(pUrl As String, _
                              Optional destFilePath As String, _
                              Optional ByVal bDebug As Boolean) As Variant
    HTTP_Download = Download(pUrl, HTTP_Request, destFilePath, bDebug)
End Function

'Função escapsulada somente para Upload
'Retorna o response para tratamento posterior
'Usa o metodo POST
Public Function HTTP_Upload(pUrl As String, _
                                  Optional vHeaders, _
                                  Optional vBodySend, _
                                  Optional pUserName As String, _
                                  Optional pPassword As String, _
                                  Optional pResponseType As eResponseType, _
                                  Optional bAssync As Boolean = False, _
                                  Optional ByVal bDebugPrint As Boolean) As Variant
    Dim vResponse As Variant
    
     vResponse = CreateHTTPSession("POST", pUrl, vHeaders, , vBodySend, pUserName, pPassword, pResponseType, bAssync, bDebugPrint)
     
     If VBA.IsObject(vResponse(1)) Then
        Set HTTP_Upload = vResponse(1)
     Else
        HTTP_Upload = vResponse(1)
     End If
     
End Function
'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cTFW_WebUtil.Download()
' TIPO             : Function
' DATA/HORA        : 30/09/2016 13:44
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Rotina otimizada para efetuar o download por meio de dois metodos :
'                    1 - Por meio de API : Nesse modo é possivel verificar byte a byte
'                        que esta dendo baixado.
'                    2 - Por meio de solicitação HTTP : Nesse modo é possivel tratar o
'                        retorno de varias formas
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      30/09/2016 13:44
'---------------------------------------------------------------------------------------
Function Download(ByVal Url As String, _
                  Optional metodo As eDownloadMethod = eDownloadMethod.API, _
                  Optional destFilePath As String, _
                  Optional ByVal bDebug As Boolean)

    Dim sFileName As String
    Dim oStream As Object
    Dim NomeArquivo As String
    Dim headers As Collection
    Dim RetDownload As Variant

    '---------------------------------------------------------------------------------------
1   On Error GoTo Download_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
2   Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Function cTFW_WebUtil.Download()"
    '---------------------------------------------------------------------------------------
3   Select Case metodo
    Case eDownloadMethod.API
4       Download = downloadViaApi(Url, bDebug, destFilePath)

5   Case eDownloadMethod.HTTP_Request
        '--------------------------------------------------------------------------------------------------------------------------------
        'Executa a chamada do HTTP que realiza o download do arquivo solicitado
        '--------------------------------------------------------------------------------------------------------------------------------
6       RetDownload = CreateHTTPSession(pMethod:="GET", pAction:=Url, pResponseType:=ResponseResultBinary, bAssync:=True, bDebugPrint:=bDebug)

        '--------------------------------------------------------------------------------------------------------------------------------
        'Trata o retorno do Download
        '--------------------------------------------------------------------------------------------------------------------------------
        
        If RetDownload(0) <> 0 Then
        
8       If destFilePath = "" Then
9           Set headers = RetDownload(2)
10          Call RaiseEventPrint("Tamanho : " & formatFileSize(getItemColletion(headers, "Content-Length")), bDebug)
11          Call RaiseEventPrint("Tipo de Arquivo : " & getItemColletion(headers, "Content-Type"), bDebug)
12          NomeArquivo = getItemColletion(headers, "Content-Disposition")

            'Configura o nome do arquivo de destino
13          If NomeArquivo = "" Then
14              NomeArquivo = Right$(Url, Len(Url) - VBA.InStrRev(Url, "/"))
15          Else
16              NomeArquivo = VBA.Replace(VBA.CStr(VBA.Split(VBA.CStr(VBA.Split(VBA.LCase(NomeArquivo), ";")(1)), "filename=")(1)), VBA.Chr(34), "")
17          End If

18          If NomeArquivo <> "" Then
19              sFileName = Environ("Temp") & "\" & NomeArquivo
20          Else
21              sFileName = Environ("Temp") & "\Download" & VBA.Format(VBA.Now, "hhmmssmmddyy") & ".tmp"
22          End If
23      Else
24          sFileName = destFilePath
25      End If
        'Descarrega o conteudo do Stream (bytes) que foi baixado e salva no arquivo
        Call salvarStreamBytes(RetDownload(1), sFileName)
        End If
        Download = sFileName
35  End Select


Fim:
36  On Error GoTo 0
37  Exit Function

Download_Error:
38  If VBA.Err <> 0 Then
39      lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
        Debug.Print cstr_ProcedureName, "Linha : " & VBA.Erl() & " - " & strErrorMessagem
41  End If
    GoTo Fim:
    'Debug Mode
42  Resume
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cTFW_WebUtil.salvarStreamBytes()
' TIPO             : Sub
' DATA/HORA        : 30/09/2016 13:52
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Cria um arquivo a partir de uma origem de buffer de bytes binario (possivel criar qualquer arquivo)
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      30/09/2016 13:52
'---------------------------------------------------------------------------------------
Private Sub salvarStreamBytes(oStreamBytesBuffer, arquivoDestino As String)
    Dim oStream As Object ' ADODB.stream
    '---------------------------------------------------------------------------------------
1   On Error GoTo salvarStreamBytes_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
2   Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Sub cTFW_WebUtil.salvarStreamBytes()"
    '---------------------------------------------------------------------------------------
    
    Debug.Print "Salvando dados..."
    
3   Set oStream = VBA.CreateObject("ADODB.Stream")
4   With oStream
5       .Open
6       .Type = 1    '1 - adTypeBinary ou 2 - adTypeText
7       Call oStream.Write(oStreamBytesBuffer)
8       If VBA.Dir(arquivoDestino) <> "" Then Call VBA.Kill(arquivoDestino)
9       Call .SaveToFile(arquivoDestino, 2)  'adSaveCreateOverWrite
10      Call .Close
11  End With
12  Set oStream = Nothing


Fim:
13  On Error GoTo 0
14  Exit Sub

salvarStreamBytes_Error:
15  If VBA.Err <> 0 Then
16      lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
        Debug.Print cstr_ProcedureName, "Linha : " & VBA.Erl() & " - " & strErrorMessagem
18  End If
    GoTo Fim:
    'Debug Mode
19  Resume
End Sub

Private Function getItemColletion(C As Collection, pKey As String)
    On Error Resume Next
    getItemColletion = C.item(pKey)
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : Download()
' TIPO             : Function
' DATA/HORA        : 30/09/2016 13:52
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Efetua o download de um arquivo qualquer especificado pela URL
'---------------------------------------------------------------------------------------
Function HTTP_Download_Basico(strURL As String, arquivoDestino As String)
    '---------------------------------------
    'Variáveis necessárias
    '---------------------------------------
    Dim http As Object
    Dim rs As Integer
    Dim status As Integer
    Dim response As Variant
    Dim oStream As Object
    Dim strTamanho_Arquivo As String
    
1   Set http = VBA.CreateObject("MSXML2.ServerXMLHTTP")
    
    Debug.Print "Baixando dados de ..." & strURL & vbNewLine & VBA.String(100, "-")
    
2   With http
        '---------------------------------------
        '1 - Envia os dados da solicitação à URL de download
        '---------------------------------------
3       Call .Open("GET", strURL, True)
4       Call .send
5
        '---------------------------------------
        '2 - Aguarda a transferencia
        '---------------------------------------
6       On Error Resume Next
7       Do While (Not rs = 4)
8           VBA.DoEvents
9           rs = http.ReadyState
10      Loop
11      status = http.status
12      Do While (Not status = 200)
13          VBA.DoEvents
14          status = http.status
15      Loop
        '---------------------------------------
        '3 - Salva os dados transferidos na memoria
        '---------------------------------------
16      response = http.ResponseBody
17  End With

    '---------------------------------------
    '4 - Descarrega os dados da memoria em arquivo no disco
    '---------------------------------------
18  On Error GoTo 0
19  Set oStream = VBA.CreateObject("ADODB.Stream")
20  With oStream
21      .Open
22      .Type = 1
23      Call oStream.Write(response)
24      If VBA.Dir(arquivoDestino) <> "" Then Call VBA.Kill(arquivoDestino)
25      Call .SaveToFile(arquivoDestino, 2)
26      Call .Close
        If VBA.Dir(arquivoDestino) <> "" Then
            Debug.Print "Arquivo baixado com sucesso. Tamanho : " & Format(VBA.FileLen(arquivoDestino) / 1000, "KB 0.00") & vbNewLine & VBA.String(100, "-")
        End If
27  End With
28  Set oStream = Nothing
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cBacen.CreateHTTPSession()
' TIPO             : Function
' DATA/HORA        : 14/11/2014 10:32
' CONSULTOR        : (ADELSON)/ TECNUN - Adelson Rosendo Marques da Silva
' CONTATO          : adelson@tecnun.com.br
' DESCRIÇÃO        : Função Generica para criar uma instancia de uma sessão HTTP
'---------------------------------------------------------------------------------------
' Informações uteis sobre as referencias que manipulam as solicitações HTTP
' ------------------------------------------------------------------------
' Descrição     : Microsoft WinHTTP Services, version 5.1
' Arquivo DLL   : C:\Windows\system32\winhttp.dll
' Nome          : WinHttp
' Arquivo DLL   : {662901FC-6951-4854-9EB2-D9A2570F2B2E}
' Classe        : WinHttp.WinHttpRequest
' MSXML 6.0     : 88d96a0a-f192-11d4-a65f-0040963251e5
' ------------------------------------------------------------------------

Private Function CreateHTTPSession(pMethod As String, _
                                   pAction As String, _
                                   Optional vHeaders, _
                                   Optional vParam As String, _
                                   Optional vBodySend, _
                                   Optional pUserName As String, _
                                   Optional pPassword As String, _
                                   Optional pResponseType As eResponseType, _
                                   Optional bAssync As Boolean = False, _
                                   Optional ByVal bDebugPrint As Boolean) As Variant

    Dim http As Object ' MSXML2.xmlhttp
    'Dim HTTP2 As WinHttp.WinHttpRequest
    Dim vHeader
    Dim rhdr
    Dim FileName As String
    Dim tURL As tpURL
    Dim oStream As Object

    Dim HTTP_Response
    Dim HTTP_status As eHTTP_Status
    Dim HTTP_ReadState As eHTTP_ReadyState
    Dim HTTP_Headers As New Collection

    '----------------------------------------------------------------------------------------------------
1   On Error GoTo CreateHTTPSession_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
2   Dim dtSartRunProc As Date: dtSartRunProc = Time
    Const cstr_ProcedureName As String = "Function cBacen.CreateHTTPSession()"
    '----------------------------------------------------------------------------------------------------

3   If Not InternetIsConnected() Then
4       lngErrorNumber = VBA.Rnd(): strErrorMessagem = "Não há conexão ativa com a Internet"
5       GoTo RetResponse
6   End If

7   Set http = VBA.CreateObject("MSXML2.ServerXMLHTTP")


8   tURL = ExtractUrl(pAction)

9   With http
        'Sintaxe para : WinHttp.WinHttpRequest
        'Call .Open(pMethod, pAction, True)
        'Sintaxe para : MSXML2.ServerXMLHTTP
10      Call .Open(pMethod, pAction, bAssync, pUserName, pPassword)

        'Configura os Headers
11      If Not VBA.IsMissing(vHeaders) Then
12          If VBA.IsArray(vHeaders) Then
13              For Each vHeader In vHeaders
14                  Call .setRequestHeader(VBA.Split(vHeader, "|")(0), VBA.Split(vHeader, "|")(1))
15              Next
16          Else
17              vHeader = vHeaders
18              Call .setRequestHeader(VBA.Split(vHeader, "|")(0), VBA.Split(vHeader, "|")(1))
19          End If
20      End If

        'Envia os Request (Download ou Upload)
21      Call .send(vBodySend)
        Debug.Print .getAllResponseHeaders()

        'Trata status de processamento atraves do ReadyState
22      HTTP_ReadState = getHTTP_ReadyState(http)

23      Call RaiseEventPrint("Aguarde, HTTP_ReadState - " & getHTTP_ReadyState_Desc(HTTP_ReadState), bDebugPrint)

24      Do While Not HTTP_ReadState = 4
25          VBA.DoEvents
26          HTTP_ReadState = getHTTP_ReadyState(http)
27          Call RaiseEventPrint("Aguarde, HTTP_ReadState - " & getHTTP_ReadyState_Desc(HTTP_ReadState), bDebugPrint)
28          If HTTP_ReadState = HTTP_ReadyState_03_DOWNLOAD And HTTP_Headers.count = 0 Then
                'Obtem a informação de todos os cabeçalhos
                'Call RaiseEventPrint(.GetAllResponseHeaders(), bDebugPrint)
29              For Each rhdr In VBA.Split(.getAllResponseHeaders(), VBA.vbNewLine)
30                  If Trim(rhdr) <> "" Then
31                      Call addHeader(HTTP_Headers, VBA.Trim(VBA.CStr(VBA.Split(VBA.LCase(rhdr), ":")(0))), VBA.Trim(VBA.CStr(VBA.Split(VBA.LCase(rhdr), ":")(1))))
32                  End If
33              Next
34              Debug.Print "Tamanho do Arquivo : " & formatFileSize(CLng(HTTP_Headers("Content-Length")), True)
35          End If
            '36          Call getHTTPResponse(HTTP)
37      Loop

38      HTTP_status = -1

        'Uma vez que o processamento esteja pronto, trata o codigo de status HTTP
39      If http.ReadyState = 4 Then
40          Do While Not HTTP_status = HTTP_Status_200_OK
41              VBA.DoEvents
42              HTTP_status = getHTTPStatus(http)
43              Call RaiseEventPrint("HTTP STATUS : " & HTTP_status & " - " & getHTTPStatus_Desc(http), bDebugPrint)

44              Select Case HTTP_status
                Case eHTTP_Status.[HTTP_Status_400_Bad Request], _
                     eHTTP_Status.HTTP_Status_403_Forbidden
45                  Call RaiseEventPrint("Bad Request ou Negado  ! " & http.StatusText, bDebugPrint)
46                  Exit Do
47              Case eHTTP_Status.[HTTP_Status_500_Internal Server VBA.Error], _
                     eHTTP_Status.[HTTP_Status_502_Bad Gateway]
48                  Call RaiseEventPrint("Erro interno do Servidor  ! " & http.StatusText, bDebugPrint)
49                  Exit Do
50              End Select
51          Loop
52      End If

RetResponse:
53      If lngErrorNumber <> 0 Then
54          HTTP_Response = strErrorMessagem
55      Else
56          Select Case pResponseType
            Case eResponseType.ResponseResultXMLDOM
57              Set HTTP_Response = .responseXML
58          Case eResponseType.ResponseResultXML
59              HTTP_Response = .responseXML.XML
60          Case eResponseType.ResponseResultTEXT
61              HTTP_Response = .responseText
62          Case eResponseType.ResponseResultBinary
63              HTTP_Response = .ResponseBody
64              If HTTP_status = HTTP_Status_200_OK Then
65                  strErrorMessagem = "HTTP Request : Dados Obtidos com sucesso : " & VBA.vbNewLine & "Origem : " & pAction & VBA.vbNewLine & "Tamanho " & formatFileSize(UBound(HTTP_Response), True)
66              End If
67          End Select
68      End If
69  End With

70  CreateHTTPSession = Array(HTTP_status, HTTP_Response, HTTP_Headers, strErrorMessagem)

71  On Error GoTo 0
72  Exit Function

CreateHTTPSession_Error:
73  If VBA.Err.Number = -2147012889 Then
74      lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.source & "-" & VBA.Err.Description
75      GoTo RetResponse
76  ElseIf VBA.Err = -2147012866 Then
77      lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.source & "-" & VBA.Err.Description
78      GoTo RetResponse
79  ElseIf VBA.Err <> 0 Then
80      lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
81      Debug.Print VBA.Now, cstr_ProcedureName & " - " & strErrorMessagem
82      HTTP_status = 0
83      GoTo RetResponse
        'Quando não há o modulo de exceções
        'Call VBA.MsgBox("Ocorreu um erro ao criar a sessão HTTP para o endereço remoto : " & pAction & VBA.vbNewLine & VBA.vbNewLine & _
         "Detalhes : " & cstr_ProcedureName & "." & VBA.Erl & " - " & lngErrorNumber & "-" & strErrorMessagem)
84  End If
85  Exit Function
    'Debug Mode
86  Resume
End Function

Private Function getHTTPResponse(oHttp As Object) As eHTTP_Status
    On Error Resume Next
    Dim rsp
    rsp = oHttp.ResponseStream
    Debug.Print formatFileSize(UBound(rsp), True)
End Function


Private Sub addHeader(cHeader As Collection, pKey As String, pValue As String)
    On Error Resume Next
    Call cHeader.Add(pValue, pKey)
End Sub

Private Function getHTTPStatus(oHttp As Object) As eHTTP_Status
    On Error Resume Next
'    Debug.Print "GET HTTP Status : " & oHttp.status
    getHTTPStatus = oHttp.status
End Function

Private Function getHTTPStatus_Desc(oHttp As Object) As String
    On Error Resume Next
    getHTTPStatus_Desc = oHttp.StatusText
End Function

Private Function getHTTP_ReadyState(oHttp As Object) As Integer
    On Error Resume Next
    'Debug.Print "GET HTTP Ready State : " & oHttp.ReadyState
    getHTTP_ReadyState = oHttp.ReadyState
End Function

Private Function getHTTP_ReadyState_Desc(eReadyState As eHTTP_ReadyState)
    getHTTP_ReadyState_Desc = VBA.Choose(eReadyState, "Uninitialized - open() has not been called yet.", _
                                                   "Loading - send() has not been called yet.", _
                                                   "Loaded - send() has been called, headers and status are available.", _
                                                   "Interactive - Downloading, responseText holds the partial data.", _
                                                   "Completed - Finished with all operations.")
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cTFW_WebUtil.InternetIsConnected()
' TIPO             : Function
' DATA/HORA        : 05/10/2015 10:37
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Determina se há uma conexão ativa com a Internet
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      05/10/2015 10:37
'---------------------------------------------------------------------------------------
Public Function InternetIsConnected(Optional ByRef ConnType As String) As Boolean
    Dim dwFlags As Long
    Dim WebTest As Boolean
    '---------------------------------------------------------------------------------------
10  On Error GoTo InternetIsConnected_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
20  Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Function cTFW_WebUtil.InternetIsConnected()"
    '---------------------------------------------------------------------------------------
30  ConnType = ""
40  WebTest = InternetGetConnectedState(dwFlags, 0&)
50  Select Case WebTest
    Case dwFlags And CONNECT_LAN: ConnType = "LAN"
60  Case dwFlags And CONNECT_MODEM: ConnType = "Modem"
70  Case dwFlags And CONNECT_PROXY: ConnType = "Proxy"
80  Case dwFlags And CONNECT_OFFLINE: ConnType = "Offline"
90  Case dwFlags And CONNECT_CONFIGURED: ConnType = "Configured"
100 Case Else: ConnType = "N/A"
110 End Select
120 InternetIsConnected = WebTest
Fim:
130 On Error GoTo 0
140 Exit Function

InternetIsConnected_Error:
150 If VBA.Err <> 0 Then
160     lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
170     Debug.Print cstr_ProcedureName, "Linha : " & VBA.Erl() & " - " & strErrorMessagem
180 End If
    GoTo Fim:
    'Debug Mode
190 Resume
End Function
'
'Function InternetIsConnected() As Boolean
'    Dim hOpen As Long
'130 hOpen = InternetOpen(API_AGENT_NAME, INTERNET_OPEN_TYPE_DIRECT,VBA.vbNullString,VBA.vbNullString, 0)
'140 If hOpen = 0 Then
'160     InternetIsConnected = False
'170     Exit Function
'    Else
'        InternetIsConnected = True
'180 End If
'End Function

' Get locally-stored cookies for a specified URL
' If CookieName is omitted, returns all the cookies as
' as semicolon-delimited list of NAME=VALUE pairs
Function GetCookies(ByVal Url As String, Optional CookieName As String) As _
         String
    Dim buffer As String
    Dim Length As Long

    ' prepare the receiving buffer
1   Length = 10240
2   buffer = Space$(Length)
    ' query WinInet for cookies from this URL
    ' a zero value means failure
3   If InternetGetCookie(Url, VBA.vbNullString, buffer, _
                         Length) = 0 Then Exit Function

    ' LENGHT has received the size of returned data
4   buffer = VBA.Left$(buffer, Length)

5   If Len(CookieName) = 0 Then
        ' the entire cookie string was requested
6       GetCookies = buffer
7   Else
        ' extract one single cookie
        Dim cookies() As String
        Dim i As Long, tmp As String

        ' get the individual cookies
8       cookies = VBA.Split(buffer, ";")

        ' search for the right one
9       For i = 0 To UBound(cookies)
            ' trim the leading space, if any
10          tmp = LTrim$(cookies(i))
11          If InStr(1, tmp, CookieName & "=", VBA.vbTextCompare) = 1 Then
                ' we've got it
12              GetCookies = Mid$(tmp, Len(CookieName) + 2)
13              Exit For
14          End If
15      Next
16  End If

End Function

'Envia uma mensagem para o evento 'RegistraLog' para que seja capturado por uma aplicação
Private Sub sendToEvent(pMensagem As String)
    RaiseEvent OnRunning(VBA.Now & " " & pMensagem)
End Sub

Private Sub RaiseEventPrint(Optional pMensagem As String, Optional bDebugPrint As Boolean)
    If bDebugPrint Then Debug.Print pMensagem
    RaiseEvent DebugPrint(pMensagem)
End Sub

'---------------------------------------------------------------
' CREDITO POST VBFORUMS EM - http://www.vbforums.com/showthread.php?334645-Winsock-Making-HTTP-POST-GET-Requests
'---------------------------------------------------------------
' returns as type URL from a string
Private Function ExtractUrl(ByVal strURL As String) As tpURL
    Dim intPos1 As Integer
    Dim intPos2 As Integer
    
    Dim retURL As tpURL
    
    '1 look for a scheme it ends with ://
    intPos1 = InStr(strURL, "://")
    
    If intPos1 > 0 Then
        retURL.Scheme = Mid(strURL, 1, intPos1 - 1)
        strURL = Mid(strURL, intPos1 + 3)
    End If
        
    '2 look for a port
    intPos1 = InStr(strURL, ":")
    intPos2 = InStr(strURL, "/")
    
    If intPos1 > 0 And intPos1 < intPos2 Then
        ' a port is specified
        retURL.Host = Mid(strURL, 1, intPos1 - 1)
        
        If (VBA.IsNumeric(Mid(strURL, intPos1 + 1, intPos2 - intPos1 - 1))) Then
                retURL.Port = CInt(Mid(strURL, intPos1 + 1, intPos2 - intPos1 - 1))
        End If
    ElseIf intPos2 > 0 Then
        retURL.Host = Mid(strURL, 1, intPos2 - 1)
    Else
        retURL.Host = strURL
        retURL.URI = "/"
        
        ExtractUrl = retURL
        Exit Function
    End If
    
    strURL = Mid(strURL, intPos2)
    
    ' find a question mark ?
    intPos1 = InStr(strURL, "?")
    
    If intPos1 > 0 Then
        retURL.URI = Mid(strURL, 1, intPos1 - 1)
        retURL.Query = Mid(strURL, intPos1 + 1)
    Else
        retURL.URI = strURL
    End If
    
    ExtractUrl = retURL
End Function

' url encodes a string
Function URLEncode(ByVal Str As String) As String
    Dim intLen As Integer
    Dim X As Integer
    Dim curChar As Long
    Dim newStr As String
    intLen = Len(Str)
    newStr = ""
    For X = 1 To intLen
        curChar = VBA.Asc(VBA.Mid$(Str, X, 1))
        If (curChar < 48 Or curChar > 57) And _
            (curChar < 65 Or curChar > 90) And _
            (curChar < 97 Or curChar > 122) Then
                            newStr = newStr & "%" & Hex(curChar)
        Else
            newStr = newStr & VBA.Chr(curChar)
        End If
    Next X
    URLEncode = newStr
End Function

' decodes a url encoded string
Function UrlDecode(ByVal Str As String) As String
    Dim intLen As Integer
    Dim X As Integer
    Dim curChar As String * 1
    Dim strCode As String * 2
    Dim newStr As String
    
    intLen = Len(Str)
    newStr = ""
    
    For X = 1 To intLen
        curChar = Mid$(Str, X, 1)
        If curChar = "%" Then
            strCode = "&h" & Mid$(Str, X + 1, 2)
            If VBA.IsNumeric(strCode) Then
                curChar = VBA.Chr(Int(strCode))
            Else
                curChar = ""
            End If
            X = X + 2
        End If
        newStr = newStr & curChar
    Next X
    UrlDecode = newStr
End Function

 'Formata o tamanho de um arquivo em bytes
Private Function formatFileSize(size As Long, Optional bIncluirUnidade As Boolean = True)
    Dim bUnidade As String
    Dim lngSize As Double
    Select Case size
        Case Is >= 1073741824: lngSize = VBA.Round(((size / 1024) / 1024) / 1024, 2): bUnidade = VBA.IIf(bIncluirUnidade, " GB", "")
        Case Is >= 1048576: lngSize = VBA.Round((size / 1024) / 1024, 2): bUnidade = VBA.IIf(bIncluirUnidade, " MB", "")
        Case Is >= 1024 And size < (size * 1024): lngSize = VBA.Round(size / 1024, 2): bUnidade = VBA.IIf(bIncluirUnidade, " KB", "")
        Case Is < 1024: lngSize = size: bUnidade = VBA.IIf(bIncluirUnidade, " Bytes", "")
    End Select
    formatFileSize = lngSize
    If bIncluirUnidade Then formatFileSize = Right(Space(8) & lngSize & bUnidade, 12)
End Function

