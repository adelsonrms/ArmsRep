VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cIE_Tab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private m_objieApp As Object

Public Enum eStopWaitAction
    STOP_WHEN_FIND_IGNORE = 0
    STOP_WHEN_FIND_TEXT = 1
    STOP_WHEN_FIND_ID = 2
End Enum

Private objDocument As Object 'MSHTML.HTMLDocument

'Eventos.............
Public Event DebugPrint(pMensagem As String)
Public Event OnRunning(pMessage As String)
Public Event OnFinish()

Private m_objIE_Browser As cIE_Application
Private m_strURL As String
Private m_strTittle As String
Private m_strTabKey As String

Private m_Inputs As Collection
Private m_Controls As Collection


Public Function GetDocument() As Object    '     As MSHTML.HTMLDocument
'---------------------------------------------------------------------------------------
1   On Error GoTo GetDocument_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
2   Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Function cIE_Tab.GetDocument()"
    '---------------------------------------------------------------------------------------
3   If PageIsLoaded() Then
4       Set objDocument = m_objieApp.Document
5       Set GetDocument = objDocument
6   End If

Fim:
7   On Error GoTo 0
8   Exit Function

GetDocument_Error:
9   If Err <> 0 Then
10      lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
        Debug.Print cstr_ProcedureName, "Linha : " & VBA.Erl() & " - " & strErrorMessagem
12  End If
    GoTo Fim:
    'Debug Mode
13  Resume
End Function

Public Property Get Document() As Object ' MSHTML.HTMLDocument
    Call GetDocument
10  Set Document = objDocument
End Property

Public Property Set Document(objDoc As Object) 'MSHTML.HTMLDocument)
10  Set objDocument = objDoc
    Call GetDocument
End Property


Public Function PageIsLoaded()
    On Error Resume Next 'Preciso apenas do True
    If Not m_objieApp Is Nothing Then
10      PageIsLoaded = Not m_objieApp.Busy
    End If
End Function
'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cIEDrive.ControlExists()
' TIPO             : Function
' DATA/HORA        : 30/10/2014 22:59
' CONSULTOR        : (ADELSON)/ TECNUN - Adelson Rosendo Marques da Silva
' CONTATO          : adelson@tecnun.com.br
' DESCRIÇÃO        : Verifica se um controle existe na instancia da pagina
'---------------------------------------------------------------------------------------
Public Function ControlExists(strIDControl As String, Optional oPage As Object) As Boolean
10  On Error Resume Next
    Dim ctr    As Object
20  If oPage Is Nothing Then Set oPage = m_objieApp.Document
30  If PageIsLoaded(oPage) Then
40      Set ctr = oPage.getElementById(strIDControl)
50      ControlExists = Not ctr Is Nothing
60  End If
End Function

Public Property Get Control(strIDControl As String) As Object
10  On Error Resume Next
20  Set m_ObjControl = Nothing
30  Set m_ObjControl = Me.Document.getElementById(strIDControl)
40  If Not m_ObjControl Is Nothing Then
50      m_StrControlValue = m_ObjControl.innertext
60  Else
        Call sendToEvent("cIDDrive.ControlValue() - Não foi encontrado nenhum controle na pagina com o ID  : " & strIDControl)
80  End If
90  Set Control = m_ObjControl
End Property

Function RefreshControlsWithID()
    Set m_Controls = GetControls(, True, True)
    Set m_Inputs = GetControls(, True, True, "INPUT")
End Function

'Recupera todos os elementos na pagina que possui IDs
Public Function GetControls(Optional oParentElement As Object, _
                       Optional bGetAll As Boolean = True, _
                       Optional filterID As Boolean = True, _
                       Optional filterByType As String = "*")
    Dim intCtr As Long
    Dim oElem As Object ' MSHTML.HTMLDivElement
    Dim subElement As Object
    Dim cIECtr As cIE_Control
    Dim colElements As Collection

    If oParentElement Is Nothing Then Set oParentElement = objDocument.DocumentElement

    Set colElements = New VBA.Collection
    Set m_Inputs = New VBA.Collection

    For Each subElement In VBA.IIf(bGetAll, oParentElement.All, oParentElement.Children)
        VBA.DoEvents
        If filterID Then
            If ControlHasAttibute(subElement, "id") Then
                GoSub CriarControl:
            End If
        Else
            GoSub CriarControl:
        End If
    Next subElement

    Set GetControls = colElements

    Exit Function

CriarControl:
    If VBA.UCase(subElement.tagName) Like "*" & VBA.UCase(filterByType) & "*" Then
        'Cria um controle (generico que tenha ID)
        Set cIECtr = addNewControl(colElements, subElement)
    End If

    'Separa os controles INPUT em collection dedicada
    'If cIECtr.TypeElement = "INPUT" Then Call m_Inputs.Add(cIECtr, cIECtr.ID)
    Return
End Function

Public Function ParseControl(element As Object) As cIE_Control
    Dim ctr As New cIE_Control
    With ctr
        Set .element = element
        .ID = getAttributeValue(element, "id")
        .Name = getAttributeValue(element, "Name")
        .TypeElement = VBA.UCase(element.tagName)
        If .TypeElement = "INPUT" Then
            .value = element.value
        Else
            .value = element.innertext
        End If
        Set .IE_Tab = Me
    End With
    Set ParseControl = ctr
End Function

'Adiciona o controle em uma coleção
Private Function addNewControl(ByRef C As Collection, subElement As Object) As cIE_Control
    On ERRRO GoTo Err_AddControl
    Dim ctr As New cIE_Control
    Set ctr = ParseControl(subElement)
    On Error Resume Next
    Call C.Add(ctr, ctr.ID)
    
    If VBA.Err.Number = 457 Then
        Call C.Add(ctr)
        VBA.Err.Clear
    End If
    On ERRRO GoTo Err_AddControl
    Set addNewControl = ctr
    Exit Function
Err_AddControl:
    If VBA.Err <> 0 Then
        MsgBox "addNewControl() > Erro : " & VBA.Error, vbCritical
    End If
    Exit Function
    Stop
    Resume
End Function

'Verifica se um determinado elemento possui o atributo especificado
Public Function ControlHasAttibute(tag As Object, pAttName As String) As Boolean
    On Error Resume Next
    Dim V
    V = tag.getAttribute(pAttName)
    If VBA.Err.Number = 0 Then
        ControlHasAttibute = Not VBA.IsNull(V)
    End If
End Function

'Recupera o valor de um atributo. O metodo padrão '' da bilioteca retorna NULL caso não ache o atributo.
'Isso gera erros inesperados no VBA
Public Function getAttributeValue(tag, pAttName As String)
    Dim V
    V = tag.getAttribute(pAttName)
    If VBA.IsNull(V) Then V = ""
    getAttributeValue = V
End Function

Public Function GetLinks(Optional oParent As Object, Optional sFilterValue As String, Optional sFilterAttribute As String = "class") As Collection
    Dim lnk As New Collection
    If oParent Is Nothing Then Set oParent = Document.DocumentElement
    If sFilterAttribute = "class" Then
        Set lnk = Document.getElementsByClassName(sFilterValue)
    Else
        Set lnk = GetElementsByTag(oParent, "a", Array(sFilterAttribute, sFilterValue))
    End If
    Set GetLinks = lnk
End Function

Public Function GetElementByTag(parent As Object, strTag As String, ParamArray arrFilterBy()) As Object
    Dim div As Object
    paramFilter = arrFilterBy
    Set div = GetElementsByTag(parent, strTag, Array(arrFilterBy(0)(0), arrFilterBy(0)(1)))
    If div.count > 0 Then Set GetElementByTag = div.item(1).element
End Function

'If div.count > 0 Then

'Recupera uma coleção de elementos especificado por uma tag e filtra os elementos que possuirem os atributos especificados
Public Function GetElementsByTag(parent As Object, strTag As String, ParamArray arrFilterBy())
    On Error GoTo Err_GetElement
    Dim colElement As Object ' MSHTML.IHTMLElementCollection
    Dim element As Object 'MSHTML.HTMLGenericElement
    Dim filterElement, cFound As Collection
    Dim f As String
    
    If parent Is Nothing Then Set parent = Me.Document

    Set colElement = parent.getElementsByTagName(strTag)
    Set cFound = New VBA.Collection

    For Each element In colElement
        VBA.DoEvents
        If UBound(arrFilterBy, 1) <> -1 Then
        For Each filterElement In arrFilterBy
            vlr = getAttributeValue(element, VBA.CStr(filterElement(0)))
            If VBA.CStr(vlr) = VBA.CStr(filterElement(1)) Then
                Call addNewControl(cFound, element)
            End If
        Next filterElement
        Else
            Call addNewControl(cFound, element)
        End If
    Next element
    Set GetElementsByTag = cFound
    
    Exit Function
Err_GetElement:
    If VBA.Err <> 0 Then
        Debug.Print "GetElementsByTag() > Erro : " & VBA.Error, vbCritical
    End If
    Exit Function
    Stop
    Resume
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cIEDrive.navigate()
' TIPO             : Sub
' DATA/HORA        : 06/11/2014 10:22
' CONSULTOR        : (ADELSON)/ TECNUN - Adelson Rosendo Marques da Silva
' CONTATO          : adelson@tecnun.com.br
' DESCRIÇÃO        : Força a navegação de uma URL e atualiza as instancias dos objetos
'---------------------------------------------------------------------------------------
Public Sub navigate(pUrl As String, _
             Optional ByRef doc As Object, _
             Optional pMsg As String)
    '----------------------------------------------------------------------------------------------------
10  On Error GoTo navigate_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
20  Dim dtSartRunProc As Date: dtSartRunProc = Time
    Const cstr_ProcedureName As String = "Sub VBAProject.cIEDrive.navigate()"
    '----------------------------------------------------------------------------------------------------
30  Call m_objieApp.navigate(pUrl)
40  Call WaitFor
50  Set doc = objDocument
    If Not doc Is Nothing Then
        Call SalvarHTML(doc.DocumentElement.innerHTML)
    End If
60  On Error GoTo 0
70  Exit Sub

navigate_Error:
80  If VBA.Err <> 0 Then
90      lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
        'Quando não há o modulo de exceções
100     Call VBA.MsgBox("Ocorreu um erro inesperado no processamento" & VBA.vbNewLine & VBA.vbNewLine & _
                        "Detalhes : " & cstr_ProcedureName & "." & VBA.Erl & " - " & lngErrorNumber & "-" & strErrorMessagem, vbCritical)
110 End If
120 Exit Sub
    'Debug Mode
130 Resume
End Sub
'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cIEDrive.ExecScript()
' TIPO             : Sub
' DATA/HORA        : 05/11/2014 13:43
' CONSULTOR        : (ADELSON)/ TECNUN - Adelson Rosendo Marques da Silva
' CONTATO          : adelson@tecnun.com.br
' DESCRIÇÃO        : Executa um Script na pagina e aguarda o carregamento
'---------------------------------------------------------------------------------------
Public Function ExecScript(strScriptCode As String, Optional ByRef doc As Object) As Boolean
'----------------------------------------------------------------------------------------------------
1   On Error GoTo ExecScript_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
2   Dim dtSartRunProc As Date: dtSartRunProc = Time
    Const cstr_ProcedureName As String = "Sub cIEDrive.ExecScript()"
    '----------------------------------------------------------------------------------------------------
3   Call sendToEvent("cIDDrive.ExecScript() - Executando o script : " & strScriptCode)
4   Call WaitFor
5   Call WaitForObject(m_objieApp)
6   Call WaitForObject(Me.Document)
7   Set objDocument = Me.GetDocument()
8   Call WaitForObject(objDocument.parentWindow)
9   Call Me.Document.parentWindow.ExecScript(strScriptCode)
10  Call Me.WaitFor
11  Set objDocument = Me.GetDocument()
12  Call WaitForObject(objDocument)
13  Call SalvarHTML(Me.Document.DocumentElement.innerHTML)
14  On Error GoTo 0
15  ExecScript = VBA.Err = 0
16  Exit Function

ExecScript_Error:
17  If VBA.Err <> 0 Then
18      lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
        'Quando não há o modulo de exceções
19      Call sendToEvent("cIDDrive.ExecScript() - Erro ao processar o script : " & strScriptCode)
20      Call sendToEvent("Detalhes : " & cstr_ProcedureName & "." & VBA.Erl & " - " & lngErrorNumber & "-" & strErrorMessagem)
21      Call VBA.MsgBox("Ocorreu um erro inesperado no processamento" & VBA.vbNewLine & VBA.vbNewLine & _
                        "Detalhes : " & cstr_ProcedureName & "." & VBA.Erl & " - " & lngErrorNumber & "-" & strErrorMessagem, vbCritical)
22  End If
23  Exit Function
    'Debug Mode
24  Resume
End Function

Function SalvarHTML(strHTML As String, Optional pArquivoDestino As String)
    Dim fs As Object ' Scripting.TextStream
    If pArquivoDestino = "" Then
        If VBA.Dir(Environ("temp") & "\IEDrive\", vbDirectory) = "" Then MkDir (Environ("temp") & "\IEDrive\")
        pArquivoDestino = Environ("temp") & "\IEDrive\" & Format(Now, "ddmmyyyy_hhnnss") & ".html"
    End If
    With VBA.CreateObject("Scripting.FileSystemObject")
        Set fs = .createTextFile(pArquivoDestino, True, False)
        Call fs.Write(strHTML)
        Call fs.Close
    End With
End Function

Function WaitForObject(obj As Object) As Boolean
1   On Error Resume Next
    Dim vCheck
2   vCheck = "..."
3   Do
4       VBA.DoEvents
5       Select Case VBA.TypeName(obj)
        Case "HTMLDocument"
6           vCheck = obj.Body.tagName: GoSub CheckValue
7       Case "InternetExplorer", "IWebBrowser2"
8           vCheck = obj.Busy: GoSub CheckValue
9       Case "HTMLWindow", "HTMLWindow2"
10          vCheck = obj.Name: GoSub CheckValue
11      End Select
        'If TypeName(obj) Like "HTMLWindow*" Then vCheck = obj.Name: GoSub CheckValue
12  Loop
13  Exit Function

CheckValue:
14  If vCheck <> "..." Then
15      WaitForObject = True
16      Exit Function
    Else
        Call sendToEvent("WaitForObject() - Aguardando objeto...Error (" & VBA.Error & ")")
17  End If
18  Return
End Function

Function PaginaComErro() As Boolean
    If ControlExists("whatToTryAlign") Then
        PaginaComErro = -1
    End If
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cIEDrive.WaitFor()
' TIPO             : Function
' DATA/HORA        : 30/10/2014 23:00
' CONSULTOR        : (ADELSON)/ TECNUN - Adelson Rosendo Marques da Silva
' CONTATO          : adelson@tecnun.com.br
' DESCRIÇÃO        : Procura por um conteudo em uma pagina
'---------------------------------------------------------------------------------------
Public Function WaitFor(Optional pText As String, _
                        Optional pStopOnFound As eStopWaitAction = STOP_WHEN_FIND_IGNORE, _
                        Optional oParentElement As Object, _
                        Optional pTimeoutSegs As Integer = 10) As Variant

'----------------------------------------------------------------------------------------------------
1   On Error GoTo WaitFor_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
2   Dim dtSartRunProc As Date: dtSartRunProc = Time
    Const cstr_ProcedureName As String = "Function cIEDrive.WaitFor()"
    'Habilitar Log de inicio/fim da rotina (-1 (ou True) / 0 ou False) ??
    Const cEnableLog As Boolean = 0
    '----------------------------------------------------------------------------------------------------

    Dim oParent As Object
    Dim vRetorno As Variant
    Dim dTimeOut As Date
    dTimeOut = VBA.DateAdd("s", pTimeoutSegs, VBA.Now())
    
    'Atualiza a instancia do pagina
3   Set objDocument = GetDocument()

    'Define o escopo do elemento de busca do conteudo
    'Caso não seja informado, será localizado na pagina inteira (DocumentElement)
4   If Not oParentElement Is Nothing Then
5       Set oParent = oParentElement
6   Else
7       If Not objDocument Is Nothing Then
8           Set oParent = objDocument.DocumentElement
9       End If
10  End If
    
    bCancelar = False
    bRodando = False
11  vRetorno = False

12  Do
        bRodando = True
        If bCancelar Then Exit Do
        Call sendToEvent("Aguarda...")
13      VBA.DoEvents
        Set objDocument = Nothing
14      Set objDocument = GetDocument()
        
        If ControlExists("whatToTryAlign") Then
            Debug.Print "# Erro na pagina "
            Exit Do
        Else

15      Select Case pStopOnFound
            'Interrompe o loop ao encontrar o texto informado
        Case eStopWaitAction.STOP_WHEN_FIND_TEXT
16          vRetorno = FindText(oParent, pText)
17          If vRetorno Then Exit Do
            'Interrompe o loop ao encontrar um elemento com o ID
18      Case eStopWaitAction.STOP_WHEN_FIND_ID
19          vRetorno = ControlExists(pText)
20          If vRetorno Then Exit Do
21      Case Else
            Set objDocument = GetDocument()
            If Not objDocument Is Nothing Then
22              vRetorno = objDocument.ReadyState = "complete" And m_objieApp.Busy = False
            End If
23          If vRetorno Then Exit Do
24      End Select
        'Se alcançou o timeout, sair
        If VBA.CDbl(dTimeOut) > 0 Then
            If VBA.Now >= dTimeOut Then
                Debug.Print "WaitFor...Excedeu Timeout !  " & pTimeoutSegs & "Segs"
                Exit Do
            End If
        End If
        End If
25  Loop
    
Fim:
    RaiseEvent OnFinish
    bRodando = False
26  Call GetDocument
27  WaitFor = vRetorno

28  On Error GoTo 0
29  Exit Function

WaitFor_Error:
    Debug.Print VBA.Err.Number, VBA.Err.Description
    Stop
30  If VBA.Err <> 0 Then
31      lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
        'Quando não há o modulo de exceções
32      Call VBA.MsgBox("Ocorreu um erro inesperado no processamento" & VBA.vbNewLine & VBA.vbNewLine & _
                        "Detalhes : " & cstr_ProcedureName & "." & VBA.Erl & " - " & lngErrorNumber & "-" & strErrorMessagem)
33  End If
34  Exit Function
    'Debug Mode
35  Resume
End Function

'Localiza o texto em outro
Private Function FindText(pElement As Object, TextToFind As String) As Boolean
    On Error GoTo Error
    If pElement Is Nothing Then Set pElement = m_objieApp.Document.Body
10  If Not pElement Is Nothing Then
20      FindText = VBA.InStr(pElement.innertext, TextToFind) > 0
30  End If
    Exit Function
Error:
    VBA.MsgBox VBA.Error, VBA.vbCritical
    Exit Function
    Resume
    Stop
End Function
'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cIEDrive.Click()
' TIPO             : Function
' DATA/HORA        : 30/10/2014 23:00
' CONSULTOR        : (ADELSON)/ TECNUN - Adelson Rosendo Marques da Silva
' CONTATO          : adelson@tecnun.com.br
' DESCRIÇÃO        : Efetua o click de um objeto
'---------------------------------------------------------------------------------------
'REVISÃO
'1.0    30/10/2014 23:00 ARMS - Iniciadoa  função
'1.1    06/03/2017 16:00 ARMS - Transformado tipo de argumento 'objItemClicked' de Object
'                               para Variant para que seja possivel passar um ID de
'                               um objeto a ser clicado
'---------------------------------------------------------------------------------------
Public Function Click(objItemClicked) As Boolean
    Dim ctr As Object
    On Error GoTo Err_Click

    If VBA.IsObject(objItemClicked) Then
        Set ctr = objItemClicked
    Else
        Set ctr = Me.Control(CStr(objItemClicked))
        Do While Not Me.ControlExists(CStr(objItemClicked))
            Me.WaitFor
            VBA.DoEvents
            Set ctr = Me.Control(CStr(objItemClicked))
        Loop
    End If

    ctr.Click

    Call WaitFor

    Set objDocument = GetDocument()
    Click = VBA.Err = 0
'    Me.RefreshControlsWithID
    Exit Function
Err_Click:
    If VBA.Err <> 0 Then
        Debug.Print "Error Control.Click() - " & VBA.Err.Description
    End If
    Click = False
End Function

'Envia uma mensagem para o evento 'RegistraLog' para que seja capturado por uma aplicação
Private Sub sendToEvent(pMensagem As String)
    Dim strMsg As String
    strMsg = VBA.Now & " " & pMensagem
    Debug.Print strMsg: RaiseEvent OnRunning(strMsg)
End Sub

Public Property Get Application() As Object
10  Set Application = m_objieApp
End Property

Public Property Set Application(objieApp As Object)
10  Set m_objieApp = objieApp
End Property

Public Property Get Title() As String
    Title = m_strTittle
End Property

Public Property Let Title(pTittle As String)
    m_strTittle = pTittle
End Property

Public Property Get Url() As String
    Url = m_strURL
End Property

Public Property Let Url(pUrl As String)
    m_strURL = pUrl
End Property

Public Property Get TabKey() As String
    TabKey = m_strTabKey
End Property

Public Property Let TabKey(pTabKey As String)
    m_strTabKey = pTabKey
End Property

Public Property Get Browser() As cIE_Application
10  Set Browser = m_objIE_Browser
End Property

Public Property Set Browser(oIEBrowser As cIE_Application)
10  Set m_objIE_Browser = oIEBrowser
End Property

Public Sub Fechar()
    If Not m_objieApp Is Nothing Then
        m_objieApp.quit
        Call RemoverTab
    End If
End Sub

Private Function RemoverTab()
    On Error GoTo Erro
    Dim cNew As New Collection
    Dim i As Long
    For i = 1 To Browser.Tabs.count
        If Me.Browser.Tabs(VBA.CLng(i)).Url Like "*" & Me.Url & "*" Then
            Call Me.Browser.Tabs.Remove(VBA.CLng(i))
            Exit Function
        End If
    Next i
    Exit Function
Erro:
    Debug.Print VBA.Error
    Exit Function
    Stop
    Resume
End Function

Public Property Get inputs() As Collection
    If m_Inputs Is Nothing Then Set m_Inputs = New Collection
    Set inputs = m_Inputs
End Property

Public Property Set inputs(ByVal vNewValue As Collection)
    Set inputs = vNewValue
End Property

Public Property Get Controls() As Collection
    If m_Controls Is Nothing Then Set m_Controls = New Collection
    Set Controls = m_Controls
End Property

Public Property Set Controls(ByVal vNewValue As Collection)
    Set Controls = vNewValue
End Property

Sub Refresh()
    On Error Resume Next
    Call Me.Application.Refresh
End Sub

