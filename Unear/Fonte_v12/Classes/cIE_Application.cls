VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cIE_Application"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_ObjControl As Object
Private m_StrControlValue As String
Private vRetValidacaoProjetos As Variant
Private strStatus As String
Private objDocument As Object
Private m_objieApp As Object

Private c_IETabs As Collection

Public Enum eWindowType
    WindowsExplorer = 1
    InternetExplorer = 2
    OutlookExplorer = 3
    AllWindowsBrowsers = 4
End Enum

Public Enum eGetInstanceType
    NewBrowserInstance = 0
    ByHandle = 1
    ByWindowCaption = 2
    ByURL = 3
    ByContentInPage = 4
    ByProperty = 5
    ByFirstIntance = 6
    ByAll = 7
End Enum

Public Enum eTypeMatchValue
    Exact = 1
    BeginsWith = 2
    EndsWith = 3
    Contains = 4
End Enum

Public Enum eDialogDownload
    BotaoSalvar = 1
    BotaoCancelar = 2
End Enum

Public Event DebugPrint(pMensagem As String)
Public Event OnRunning(pMessage As String)
Private clsUtil As New cUtil

Private Const CLSID_MANAGER_TAB As String = "new:{DF4FCC34-067A-4E0A-8352-4A1A5095346E}"
Private Const CLSID_MEDIUM_TAB As String = "new:{D5E8041D-920F-45e9-B8FB-B1DEB82C6E5E}"

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : NovaTab()
' TIPO             : Function
' DATA/HORA        : 08/04/2014 09:42
' CONSULTOR        : (ADELSON)/ TECNUN - Adelson Rosendo Marques da Silva
' CONTATO          : adelson@tecnun.com.br
' DESCRIÇÃO        : Inicializar uma nova instancia do Internet Explorer.
'                    Caso seja informado a URL, a mesma será carregada na pagina inciial
'---------------------------------------------------------------------------------------
Public Function NovaTab(Optional pUrl As String, Optional ModoInstancia As eGetInstanceType = eGetInstanceType.NewBrowserInstance) As cIE_Tab
10  On Error GoTo iniciarInternetExplorer_Error
    Call Me.PegarTab(pUrl, ModoInstancia, "", Contains)
    Set NovaTab = Me.FirstTab
120 On Error GoTo 0
130 Exit Function

iniciarInternetExplorer_Error:
140 If VBA.Err <> 0 Then Call VBA.MsgBox("Erro ao inicializar o Internet Explorer...." & VBA.vbNewLine & VBA.vbNewLine & VBA.Err.Number & " - " & VBA.Err.Description & "Modulo : IEDrive.iniciarInternetExplorer() > Linha & " & VBA.Erl, VBA.vbCritical)
150 Exit Function
160 Resume
End Function

Function startNewIE(pUrl As String) As cIE_Tab

10  On Error GoTo iniciarInternetExplorer_Error

    Dim win As Object ' SHDocVw.InternetExplorerMedium
    'Dim wb As SHDocVw.WebBrowser
    'Dim wb2 As SHDocVw.WebBrowser_V1
    'win.Visible = 1
    'Set win = VBA.CreateObject("InternetExplorer.Application")
    Set win = VBA.CreateObject(CLSID_MEDIUM_TAB) '("InternetExplorer.Application")
    'Set win = New SHDocVw.InternetExplorerMedium
    win.Visible = True
    'win.Height = 0
    'win.Width = 0
    Call clsUtil.OcultarJanela(win.hwnd)
    Set startNewIE = addNewIeDoc(win, c_IETabs, pUrl)
    'win.Visible = False

120 On Error GoTo 0
130 Exit Function

iniciarInternetExplorer_Error:
140 If VBA.Err <> 0 Then Call VBA.MsgBox("Erro ao inicializar o Internet Explorer...." & VBA.vbNewLine & VBA.vbNewLine & VBA.Err.Number & " - " & VBA.Err.Description & "Modulo : IEDrive.iniciarInternetExplorer() > Linha & " & VBA.Erl, VBA.vbCritical)
150 Exit Function
160 Resume

End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cTFW_WebUtil.getBrowser()
' TIPO             : Function
' DATA/HORA        : 08/11/2014 17:11
' CONSULTOR        : (ADELSON)/ TECNUN - Adelson Rosendo Marques da Silva
' CONTATO          : adelson@tecnun.com.br
' DESCRIÇÃO        : Retorna a instancia de browser
'---------------------------------------------------------------------------------------
Function PegarTab(Optional pUrl As String, _
                  Optional ModoInstancia As eGetInstanceType = eGetInstanceType.ByFirstIntance, _
                  Optional value As Variant, _
                  Optional typeMatchValue As eTypeMatchValue = eTypeMatchValue.Exact) As Collection

    Dim objWindows As Object ' SHDocVw.ShellWindows
    Dim win    As Object
    Dim objWebBrowser As Object
    Dim cBrowsers As New Collection
    Dim pMatchValue As String
    '----------------------------------------------------------------------------------------------------
10  On Error GoTo getBrowser_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
20  Dim dtSartRunProc As Date: dtSartRunProc = Time
    Const cstr_ProcedureName As String = "Function prjArmsFinanceManager.cTFW_WebUtil.getBrowser()"
    '----------------------------------------------------------------------------------------------------

    If ModoInstancia = NewBrowserInstance Then
        'Set win = GetObject("new:{D5E8041D-920F-45e9-B8FB-B1DEB82C6E5E}") 'Instancia atraves do CLSID do Internet Explorer
        Set win = VBA.CreateObject(CLSID_MEDIUM_TAB) '("InternetExplorer.Application")
        Call addNewIeDoc(win, c_IETabs, pUrl)
        GoTo Fim
    End If

    'Recupera todas as janelas atualmenta abertas pelo Windows Explorer, Internet Explorer e tambem do Outlook
    'O GUID {9BA05972-F6A8-11CF-A442-00A0C90A8F39} representa a classe ShellWindows, tambem disponivel na biblioteca SHDOcVw.ShellWindows
30  Set objWindows = VBA.CreateObject("new:{9BA05972-F6A8-11CF-A442-00A0C90A8F39}")

40  Select Case typeMatchValue
        Case eTypeMatchValue.Exact
50          pMatchValue = CStr(value)
60      Case eTypeMatchValue.BeginsWith
70          pMatchValue = CStr(value) & "*"
80      Case eTypeMatchValue.EndsWith
90          pMatchValue = "*" & CStr(value)
100     Case eTypeMatchValue.Contains
110         pMatchValue = "*" & CStr(value) & "*"
120 End Select

    Set c_IETabs = New Collection

    For Each win In objWindows
        VBA.DoEvents
        If win.ReadyState <> 0 Then '
            If VBA.TypeName(win.Document) = "HTMLDocument" Then
                Select Case ModoInstancia
                    Case eGetInstanceType.ByFirstIntance
                        Call addNewIeDoc(win, c_IETabs, pUrl): Exit For
                    Case eGetInstanceType.ByURL
                        If win.LocationURL Like pMatchValue Then Call addNewIeDoc(win, c_IETabs, pUrl)
                    Case eGetInstanceType.ByWindowCaption
                        If win.Document.Title Like pMatchValue Then
                            'Set objWebBrowser = win: cBrowsers.Add win.Document
                            Call addNewIeDoc(win, c_IETabs, pUrl)
                        End If
                    Case eGetInstanceType.ByAll
                        Call addNewIeDoc(win, c_IETabs, pUrl)
                        'cBrowsers.Add win.Document
                End Select
            End If
        End If
    Next win

Fim:
450 Set objWindows = Nothing
460 Set PegarTab = c_IETabs

470 On Error GoTo 0
480 Exit Function

getBrowser_Error:
490 If VBA.Err <> 0 Then
500     lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
        'Quando não há o modulo de exceções
510     Call VBA.MsgBox("Ocorreu um erro inesperado no processamento" & VBA.vbNewLine & VBA.vbNewLine & _
                        "Detalhes : " & cstr_ProcedureName & "." & VBA.Erl & " - " & lngErrorNumber & "-" & strErrorMessagem)
520 End If
530 Exit Function
    'Debug Mode
540 Resume
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cIE_Application.addNewIeDoc()
' TIPO             : Function
' DATA/HORA        : 28/10/2016 23:26
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Atualiza a lista de tabs abertas
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      28/10/2016 23:26
'---------------------------------------------------------------------------------------
Private Function addNewIeDoc(IEApp As Object, ByRef onTabs As Collection, pUrl As String) As cIE_Tab
    Dim IETab As cIE_Tab
    Dim doc As Object

    '---------------------------------------------------------------------------------------
1   On Error GoTo addNewIeDoc_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
2   Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Function cIE_Application.addNewIeDoc()"
    '---------------------------------------------------------------------------------------
3   If onTabs Is Nothing Then Set onTabs = New VBA.Collection
4   Set IETab = New cIE_Tab
5   Set IETab.Browser = Me
    Set IETab.Browser.IEApp = IEApp
6   Set IETab.Application = IEApp

7   If pUrl <> "" Then
        If IEApp.LocationURL <> pUrl Then Call IETab.navigate(pUrl)
9       Set doc = IETab.Document
10  Else
11      Set IETab.Document = doc
12  End If
    
13  IETab.Title = IETab.Document.Title
14  IETab.Url = IEApp.LocationURL
15  IETab.TabKey = onTabs.count + 1 & "|" & IETab.Url
16  Call onTabs.Add(IETab, IETab.TabKey)

    Set addNewIeDoc = IETab

Fim:
17  On Error GoTo 0
18  Exit Function

addNewIeDoc_Error:
19  If VBA.Err <> 0 Then
20      lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
        Debug.Print cstr_ProcedureName, "Linha : " & VBA.Erl() & " - " & strErrorMessagem
22  End If
    GoTo Fim:
    'Debug Mode
23  Resume
End Function

'Recupera a primeira Tab
Public Property Get FirstTab() As cIE_Tab
    If Me.Tabs.count > 0 Then Set FirstTab = Me.Tabs(1)
End Property

'Recupera a ultima Tab
Public Property Get LastTab() As cIE_Tab
    If Me.Tabs.count > 0 Then Set LastTab = Me.Tabs(Me.Tabs.count)
End Property

Sub quit()
10  Me.IEApp.quit
20  Set Me.IEApp = Nothing
End Sub
'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cIEDrive.BroswerIsOpen()
' TIPO             : Function
' DATA/HORA        : 30/10/2014 23:00
' CONSULTOR        : (ADELSON)/ TECNUN - Adelson Rosendo Marques da Silva
' CONTATO          : adelson@tecnun.com.br
' DESCRIÇÃO        : Verifica se o internet explorer esta aberto
'---------------------------------------------------------------------------------------
Function BroswerIsOpen() As Boolean
    Dim lng
10  On Error Resume Next
20  If Not IEApp Is Nothing Then
30      lng = IEApp.hwnd
40      If lng Then BroswerIsOpen = True
50  End If
End Function
'Torna visivel a janela do Browser
Sub ShowWindow()
10  Me.IEApp.Visible = 1
End Sub

Public Property Get Tabs() As Collection
    If c_IETabs Is Nothing Then Set c_IETabs = New VBA.Collection
10  Set Tabs = c_IETabs
End Property

Public Property Set Tabs(obj_IETabs As Collection)
10  Set c_IETabs = obj_IETabs
End Property

Public Property Get IEApp() As Object
10  Set IEApp = m_objieApp
End Property

Public Property Set IEApp(objieApp As Object)
10  Set m_objieApp = objieApp
End Property

Private Sub RaiseEventPrint(Optional pMensagem As String, Optional bDebugPrint As Boolean)
    If bDebugPrint Then Debug.Print pMensagem
    RaiseEvent DebugPrint(pMensagem)
End Sub

'Envia uma mensagem para o evento 'RegistraLog' para que seja capturado por uma aplicação
Private Sub sendToEvent(pMensagem As String)
    RaiseEvent OnRunning(VBA.Now & " " & pMensagem)
End Sub

'''Private Sub AdicionarReferenciaUIAutomation()
'''    Call Application.VBE.ActiveVBProject.References.AddFromFile("C:\Windows\SysWOW64\UIAutomationCore.dll")
'''End Sub
'''
'''Function NotificationBaseIsVisible() As Boolean
'''    Dim btn As IUIAutomationElement
'''    Set btn = GetButtonInNotificationBar(BotaoSalvar)
'''    NotificationBaseIsVisible = Not btn Is Nothing
'''End Function
'''
'''Function CloseNotificationBase()
'''    If NotificationBaseIsVisible Then
'''        Call DownloadDialog_Click(BotaoCancelar)
'''    End If
'''End Function
'''
'''
''''Aguarda ate que a barra de salvamento de downalod esteja disponivel
'''Function WaitDownloadBar()
'''    '---------------------------------------------------------------------------------------
'''    On Error GoTo WaitDownloadBar_Error
'''    Dim lngErrorNumber As Long, strErrorMessagem As String:
'''    Const cstr_ProcedureName As String = "GetButtonInNotificationBar()"
'''    '---------------------------------------------------------------------------------------
'''    Dim btn As IUIAutomationElement
'''    Set btn = GetButtonInNotificationBar(BotaoSalvar)
'''    Do While btn Is Nothing
'''        VBA.DoEvents
'''        Set btn = GetButtonInNotificationBar(BotaoSalvar)
'''        Debug.Print "Aguardando botão Salvar..."
'''    Loop
'''Fim:
'''15  Set WaitDownloadBar = btn
'''
'''16  On Error GoTo 0
'''17  Exit Function
'''
'''WaitDownloadBar_Error:
'''18  If Err <> 0 Then
'''19      lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
'''20      Call VBA.Err.Raise(VBA.Err.Number, cstr_ProcedureName, VBA.vbNewLine & cstr_ProcedureName & " Em" & VBA.IIf(VBA.Erl() <> 0, " > Linha (" & VBA.Erl() & ")", "") & " > " & VBA.Err.Description)
'''21  End If
'''    GoTo Fim:
'''    'Debug Mode
'''22  Resume
'''End Function
'''
''''Essa função é responsavel apenas para localizar o Hwnd do Notification Bar de dwonalod
'''Private Function getDownloadNotificationBar() As Long
'''    Dim wNotificationBar As New cTFW_Window
'''    Dim wIE As New cTFW_Window
'''    'Pega a janela do IE
'''    Call wIE.GetWindow(Me.IEApp.hwnd)
'''    getDownloadNotificationBar = wIE.GetChild("Frame Notification Bar").hwnd
'''End Function
'''
'''Function GetButtonInNotificationBar(pBotao As eDialogDownload) As IUIAutomationElement
'''    '---------------------------------------------------------------------------------------
'''1   On Error GoTo GetButtonInNotificationBar_Error
'''    Dim lngErrorNumber As Long, strErrorMessagem As String:
'''    Const cstr_ProcedureName As String = "GetButtonInNotificationBar()"
'''    '---------------------------------------------------------------------------------------
'''    Dim UIAut                   As IUIAutomation
'''    Dim UIAutNotificationBar    As IUIAutomationElement
'''    Dim btnSalvar               As IUIAutomationElement
'''    Dim iCnd                    As IUIAutomationCondition
'''    Dim lhWnd                   As Long
'''    '---------------------------------------------------------------------------------------
'''
'''    Set UIAut = New CUIAutomation
'''    lhWnd = getDownloadNotificationBar()
'''    If lhWnd = 0 Then Exit Function
'''    'Retorna o UIAutNotificationBar da Notification Bar
'''    Set UIAutNotificationBar = UIAut.ElementFromHandle(ByVal lhWnd)
'''    'Define a condição de como o engine vai localizar o botão
'''    'Atraves da propriedade name = Salvar
'''    'UIA_PropertyIds.UIA_NamePropertyId =30005
'''    Set iCnd = UIAut.CreatePropertyCondition(UIA_PropertyIds.UIA_NamePropertyId, TextoBotao(pBotao))
'''    'Localiza o primeiro elemento encontrado pelo Caption
'''    'TreeScope.TreeScope_Subtree = 7
'''    Set btnSalvar = UIAutNotificationBar.FindFirst(TreeScope.TreeScope_Subtree, iCnd)
'''
'''Fim:
'''15  Set GetButtonInNotificationBar = btnSalvar
'''
'''16  On Error GoTo 0
'''17  Exit Function
'''
'''GetButtonInNotificationBar_Error:
'''18  If VBA.Err <> 0 Then
'''19      lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
'''20      Call VBA.Err.Raise(VBA.Err.Number, cstr_ProcedureName, VBA.vbNewLine & cstr_ProcedureName & " Em" & VBA.IIf(VBA.Erl() <> 0, " > Linha (" & VBA.Erl() & ")", "") & " > " & VBA.Err.Description)
'''21  End If
'''    GoTo Fim:
'''    'Debug Mode
'''22  Resume
'''End Function
'''
''''Envia o clique em um botão da barra de download
'''Function DownloadDialog_Click(pBotao As eDialogDownload)
'''    '---------------------------------------------------------------------------------------
'''1   On Error GoTo DownloadDialog_Click_Error
'''    Dim lngErrorNumber As Long, strErrorMessagem As String:
'''    Const cstr_ProcedureName As String = "DownloadDialog_Click()"
'''    '---------------------------------------------------------------------------------------
'''    'Objetos de Interação com a interface de usuário. Serão utilizados para identificar e Clicar no botão 'Salvar'
'''    Dim btnSalvar               As IUIAutomationElement
'''    Dim InvokePattern           As IUIAutomationInvokePattern
'''    '---------------------------------------------------------------------------------------
'''    Set btnSalvar = GetButtonInNotificationBar(pBotao)
'''    If Not btnSalvar Is Nothing Then
'''        'Cria um objeto que define a ação que será executada pelo elemento
'''        'UIA_PatternIds.UIA_InvokePatternId = 10000
'''        Set InvokePattern = btnSalvar.GetCurrentPattern(UIA_PatternIds.UIA_InvokePatternId)
'''        'Invoca (Executa) a ação (Invoke). No caso, no botão, o padrão é o Click
'''        InvokePattern.Invoke
'''        Debug.Print "Botao Salvar foi clicado !"
'''    End If
'''Fim:
'''15  DownloadDialog_Click = VBA.Err = 0
'''
'''16  On Error GoTo 0
'''17  Exit Function
'''
'''DownloadDialog_Click_Error:
'''18  If Err <> 0 Then
'''19      lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
'''20      Call VBA.Err.Raise(VBA.Err.Number, cstr_ProcedureName, VBA.vbNewLine & cstr_ProcedureName & " Em" & VBA.IIf(VBA.Erl() <> 0, " > Linha (" & VBA.Erl() & ")", "") & " > " & VBA.Err.Description)
'''21  End If
'''    GoTo Fim:
'''    'Debug Mode
'''22  Resume
'''End Function
'''
'''Private Function TextoBotao(btn As eDialogDownload) As String
'''    TextoBotao = VBA.Choose(btn, "Salvar", "Cancelar")
'''End Function
