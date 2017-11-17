VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTFW_Outlook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Database
Private m_objApplication As Object
Private m_objWorkbooks As Object
Private m_objWorksheets As Object
Private m_objRange As Object
Private m_objThisWorkbook As Object
Private m_objSession As Object
Dim objMail As Object  'Outlook.MailItem

Public Enum OlItemType
    olMailItem = 0
End Enum

Public Enum OlUserPropertyType
    olText = 1
End Enum

Private Enum OlDefaultFolders
    olFolderInbox = 6
    olFolderOutbox = 4
    olFolderSentMail = 5
    olFolderDrafts = 16
    olFolderContacts = 10
End Enum

Public Enum olPastaPadrao
    eCaixaDeEntrada = 6
    eCaixaDeSaida = 4
    eItensEnviados = OlDefaultFolders.olFolderSentMail
    eContatos = OlDefaultFolders.olFolderContacts
End Enum

Public Enum eMetodoBusca
    PropriedadePersonalizada = 1
    Find_PorQuery = 2
    Find_PorID = 3
End Enum

Private Const ProgID_App As String = "Outlook.Application"
Private Sub Class_Initialize()
'  Call InicializarOutlook
End Sub

Private Sub Class_Terminate()
10  Set m_objApplication = Nothing
20  Set m_objWorkbooks = Nothing
30  Set m_objWorksheets = Nothing
40  Set m_objRange = Nothing
50  Set m_objThisWorkbook = Nothing
End Sub

'Inicializa uma instancia do Excel.
'Caso seja passado uma instancia ja criada, aproveita a instancia
Function AbrirOutlook(Optional objInstancia As Object) As Object
10  If objInstancia Is Nothing Then
20      Call InicializarOutlook
30  Else
40      Set m_objApplication = objInstancia
50  End If
End Function

Public Property Get Application() As Object
10  If m_objApplication Is Nothing Then InicializarOutlook
20  Set Application = m_objApplication
End Property

Public Property Set Application(objApplication As Object)
10  Set m_objApplication = objApplication
End Property
'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cEmail.InicializarOutlook()
' TIPO             : Function
' DATA/HORA        : 04/07/2016 10:28
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Inicializa uma nova instancia do Outlook
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      04/07/2016 10:28
'---------------------------------------------------------------------------------------
Private Function InicializarOutlook(Optional bNew As Boolean) As Object
'Inicializa o Outlook
'---------------------------------------------------------------------------------------
10  On Error GoTo InicializarOutlook_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
20  Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Function cEmail.InicializarOutlook()"
    '---------------------------------------------------------------------------------------
30  On Error Resume Next
    'Cria/Recupera a instancia do outlook
40  If Not bNew Then Set m_objApplication = VBA.GetObject(Class:=ProgID_App)
50  If m_objApplication Is Nothing Then Set m_objApplication = VBA.CreateObject(Class:=ProgID_App)
60  Set Me.Session = m_objApplication.Session
70  Set m_objApplication = m_objApplication
Fim:
80  On Error GoTo 0
90  Exit Function

InicializarOutlook_Error:
100 If Err <> 0 Then
110     lngErrorNumber = VBA.Err.Number: strErrorMessagem = "Erro na instalação do Outlook " & vbNewLine & vbNewLine & VBA.Err.Description
120     Debug.Print cstr_ProcedureName, "Linha : " & VBA.Erl() & " - " & strErrorMessagem
130 End If
    GoTo Fim:
    'Debug Mode
140 Resume
End Function

Function PastaPadrao(eTipoPasta As olPastaPadrao) As Object
    Dim oFD As Object
10  On Error Resume Next
20  Set oFD = m_objApplication.Session.GetDefaultFolder(eTipoPasta)
30  If VBA.Err.Number = 0 Then Set PastaPadrao = oFD
End Function

Sub ExibirOutlook(Optional bShow As Boolean = True)
10  Me.Application.Visible = bShow
End Sub

'Força o fechamento de todas as instancias do Excel
Sub FecharProcessoOutlook(Optional bTodas As Boolean = False, Optional bForce As Boolean = True)
    Dim lngRet As Long
10  If bTodas Then
20      lngRet = VBA.Shell("taskkill.exe /im outlook.exe " & VBA.IIf(bForce, " /f", ""))
30      If lngRet <> 0 Then
40          Debug.Print "Erro ao finalizar os processos do Excel em Andamento"
50      End If
60  Else
70      Call m_objApplication.quit
80      Set m_objApplication = Nothing
90  End If
End Sub
'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : FileExists
' TIPO             : Function
' DATA/HORA        : 29/05/2009  - 12:26
' DESENVOLVEDOR    : MONDIAL TECNOLOGIA EM INFORMATICA
'                    www.mondial.com.br
' CONSULTOR        : Adelson Rosendo Marques da Silva
' DESCRIÇÃO        : Verifica se um arquivo existe no computador
'---------------------------------------------------------------------------------------
'
' + HISTÓRIO DE REVISÃO
'---------------------------------------------------------------------------------------
' DATA          | DESCRIÇÃO
'---------------------------------------------------------------------------------------
' 12:26    | Criação da rotina
'---------------------------------------------------------------------------------------
Private Function FileExists(sFileName As String) As Boolean
10  On Error Resume Next
    Dim sTempFileName As String, fso As Object
    'Tenta primeiramente pelo Fso
20  If fso Is Nothing Then Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    'Verifica se o FSO foi instanciado (por algum problema pode não instanciar)
30  If Not fso Is Nothing Then
40      FileExists = fso.FileExists(sFileName)
50  Else
        'Se o Fso não estiver instalado, tenta a função Dir
60      sTempFileName = VBA.Dir(sFileName)
70      FileExists = sTempFileName <> ""
80  End If
    'Limpa a variável
90  Set fso = Nothing
End Function

Private Function FileIsLocked(sFilePath As String) As Boolean
10  On Error GoTo IsOpen
    Dim intFile As Long
20  If Not FileExists(sFilePath) Then Exit Function
30  intFile = VBA.FreeFile
40  Open sFilePath For Binary Access Read Lock Read As intFile: Close intFile
50  FileIsLocked = False: Exit Function
IsOpen:
60  FileIsLocked = True
End Function

Public Property Get Session() As Object
10  Set Session = m_objSession
End Property

Public Property Set Session(objSession As Object)
10  Set m_objSession = objSession
End Property

Public Function DefinirPropriedadePersonalizada(oItem As Object, PropName As String, Optional oPropValue As String)
10  On Error GoTo Erro
    Dim oProp As Object    ' UserProperty
20  If Not oItem Is Nothing Then
30      Set oProp = oItem.UserProperties.Find(PropName, True)
40      If oProp Is Nothing Then Set oProp = oItem.UserProperties.Add(PropName, OlUserPropertyType.olText)
50      If Not oProp Is Nothing Then
60          oProp.value = oPropValue
70          Set DefinirPropriedadePersonalizada = oProp
80      End If
90  End If

100 Exit Function
Erro:
110 If Err <> 0 Then
120     MsgBox Error, vbCritical
130 End If
140 Exit Function
150 Resume

End Function

Sub ExibirEmail(objEmail As Object)
10  If objEmail Is Nothing Then
20      MsgBox "Mensagem não encontrada !", vbExclamation
30  Else
40      objEmail.Display
50  End If
End Sub

Function Localizar(fdProcura As Object, metodoBusca As eMetodoBusca, ParamArray arrCriterios()) As Object
10  On Error GoTo Erro
    Dim iTotal As Integer
    Dim iContagem As Long
    Dim vCheck As String
    Dim oItems As Object
    Dim item As Object
    Dim dtStart As Date
    Dim pValorProcura
    Dim pCriterio As String

20  dtStart = VBA.Now
    Dim sFiltro As String
30  Select Case metodoBusca
    Case eMetodoBusca.PropriedadePersonalizada
40      If fdProcura.Items.count = 0 Then GoTo Fim
50      pCriterio = arrCriterios(0)
60      pValorProcura = arrCriterios(1)
70      Set item = LocalizarPorPropriedade(fdProcura, VBA.CStr(arrCriterios(0)), VBA.CStr(arrCriterios(1)))
80  Case eMetodoBusca.Find_PorQuery
90      If fdProcura.Items.count = 0 Then GoTo Fim
100     sFiltro = MontarCriterioFiltro(pCriterio, pValorProcura, "STR", "=")
110     Set item = LocalizarPorFind(fdProcura, sFiltro)
120 Case eMetodoBusca.Find_PorID
130     Set item = LocalizarPorID(VBA.CStr(arrCriterios(0)))
140 End Select
Fim:
160 Set Localizar = item
170 Exit Function
Erro:
180 If Err <> 0 Then
190     MsgBox Error, vbCritical
200 End If
210 Exit Function
220 Resume

End Function

Private Function LocalizarPorID(pEntryID As String) As Object
    Dim item As Object
10  On Error Resume Next
20  Set item = Application.Session.GetItemFromID(pEntryID)
30  If Not item Is Nothing Then
40      Set LocalizarPorID = item
50  End If
End Function

Private Function LocalizarPorFind(fdProcura As Object, pFiltro As String) As Object
10  On Error GoTo Erro
    Dim cItems As Object    'MailItem

20  Set cItems = fdProcura.Items.Find(pFiltro)
30  Set LocalizarPorFind = cItems.item(1)

40  Exit Function
Erro:
50  If Err <> 0 Then
60      MsgBox Error, vbCritical
70  End If
80  Exit Function
90  Resume
End Function

Private Function LocalizarPorPropriedade(fdProcura As Object, pCriterio As String, pValorProcura As String) As Object
10  On Error GoTo Erro
    Dim iTotal As Integer
    Dim iContagem As Long
    Dim vCheck As String
    Dim oItems As Object    ' Outlook.Items
    Dim item As Object    'MailItem
    Dim dtStart As Date

20  dtStart = VBA.Now
30  Set oItems = fdProcura.Items

40  If oItems.count = 0 Then
50      'Debug.Print "Não há mensagens pendentes de envio. Cancelando..."
60  Else
70      For iTotal = oItems.count To 1 Step -1
80          iContagem = iContagem + 1
90          VBA.DoEvents
100         vCheck = ""
110         On Error Resume Next
120         Set item = oItems.item(iTotal)
130         vCheck = item.UserProperties("IDSENT").value
140         On Error GoTo Erro
150         If vCheck <> "" Then
160             If vCheck = pValorProcura Then
170                 'Debug.Print VBA.Space(3) & "Email encontrado ! - EntryID : " & item.EntryID
180                 Exit For
190             End If
200         End If
210     Next
220 End If
240 Set LocalizarPorPropriedade = item
250 Exit Function
Erro:
260 If Err <> 0 Then
270     MsgBox Error, vbCritical
280 End If
290 Exit Function
300 Resume
End Function

Private Function MontarCriterioFiltro(pCriterio As String, pValorProcura, Optional TIPO As String, Optional operador As String = "=") As String
10  Select Case VBA.UCase(TIPO)
    Case "STR"
20      MontarCriterioFiltro = "[" & pCriterio & "] " & operador & "'" & pValorProcura & "'"
30  Case "DT"
40      MontarCriterioFiltro = "[" & pCriterio & "] " & operador & " '" & VBA.Format(CDate(pValorProcura), "ddddd h:nn AMPM") & "'"
50  Case "NUM"
60      MontarCriterioFiltro = "[" & pCriterio & "] " & operador & " '" & VBA.Format(CDate(pValorProcura), "ddddd h:nn AMPM") & "'"
70  Case "BOOL"
80      MontarCriterioFiltro = "[" & pCriterio & "] " & operador & VBA.IIf(pValorProcura, "TRUE", "FALSE")
90  End Select
End Function

Private Function PegarPastaRaiz(PastaRaiz) As Object    ' Folder
10  On Error GoTo ErrorHandler
    Dim oRaiz As Object    'Folder
    'olFolderInbox = 6
20  If VBA.IsMissing(PastaRaiz) Then
30      Set oRaiz = Application.Session.GetDefaultFolder(OlDefaultFolders.olFolderInbox).Store.GetRootFolder
40  Else
50      Set oRaiz = Application.Session.Folders(Application.Session.GetDefaultFolder(OlDefaultFolders.olFolderInbox).Store.GetRootFolder.Name).Folders(PastaRaiz)
60  End If
70  Set PegarPastaRaiz = oRaiz
80  Exit Function
ErrorHandler:
90  MsgBox Error
100 Exit Function
110 Resume
End Function

'Cria uma pasta no perfil da caixa atual
Function CriarPastaOutlook(pNomePasta As String, Optional PastaRaiz)
10  On Error GoTo ErrorHandler
    Dim oRootFolder As Object    ' Outlook.Folder
    Dim myNotesFolder As Object    'Outlook.Folder
20  If VBA.IsMissing(PastaRaiz) Then
30      Set oRootFolder = PegarPastaRaiz(PastaRaiz)
40  Else
50      Set oRootFolder = PastaRaiz
60  End If
70  If Not PastaExiste(pNomePasta, oRootFolder) Then
80      Set myNotesFolder = oRootFolder.Folders.Add(pNomePasta, olFolderInbox)
90  Else
100     Set myNotesFolder = oRootFolder.Folders(pNomePasta)
110 End If

120 Set CriarPastaOutlook = myNotesFolder
Exit Function
ErrorHandler:
If VBA.Err <> 0 Then
    VBA.MsgBox "Ocorreu um erro inesperado !" & vbNewLine & vbNewLine & "CriarPastaOutlook() - Erro : " & VBA.Error, VBA.vbCritical, "Erro"
    Exit Function
End If
Resume
End Function

'Verifica se uma pasta no Outlook existe
Function PastaExiste(sPasta As String, Optional PastaRaiz) As Boolean
    On Error GoTo ErrorHandler
    Dim oRootFolder As Object    ' Outlook.Folder
10  If VBA.IsMissing(PastaRaiz) Then
20      Set oRootFolder = PegarPastaRaiz(PastaRaiz)
30  Else
40      Set oRootFolder = PastaRaiz
50  End If
60  On Error Resume Next
70  PastaExiste = Not oRootFolder.Folders(sPasta) Is Nothing

Exit Function
ErrorHandler:
If VBA.Err <> 0 Then
    VBA.MsgBox "Ocorreu um erro inesperado !" & vbNewLine & vbNewLine & "CriarPastaOutlook() - Erro : " & VBA.Error, VBA.vbCritical, "Erro"
    Exit Function
End If
Resume
End Function

Private Sub objMail_Send(Cancel As Boolean)
10  Debug.Print VBA.Now(), "Iniciado o envio do email..."
End Sub
'Recupera a caixa de correio padrão do perfil atual
Property Get CaixaDeCorreioPadrao() As String
    On Error Resume Next
    CaixaDeCorreioPadrao = Application.Session.DefaultStore.DisplayName
End Property

'Move uma item de email para uma pasta especificada
Function MoverItemPara(item As Object, pFolder)
10  On Error GoTo ErrorHandler
    Dim oFolder As Object
    If VBA.IsObject(pFolder) Then
        Set oFolder = pFolder
    End If
    If Not oFolder Is Nothing Then Call item.Move(oFolder)
Exit Function
ErrorHandler:
If VBA.Err <> 0 Then
    VBA.MsgBox "Ocorreu um erro inesperado !" & vbNewLine & vbNewLine & "MoverItemPara() - Erro : " & VBA.Error, VBA.vbCritical, "Erro"
    Exit Function
End If
Resume
End Function

'23/07/2017 - Abre a caixa de dialogo para seleção de nomes
Function SelecionarNomes(Optional ByRef pRecipientes As Object) As String
10  On Error GoTo ErrorHandler
    Dim oDialog As Object 'SelectNamesDialog
    Dim nome As String
    Dim iRecp As Integer
    Dim colNomeOK As New VBA.Collection
    Dim colNomeErro As New VBA.Collection
    
    Set oDialog = Me.Application.Session.GetSelectNamesDialog
    With oDialog
        If Not VBA.IsMissing(pRecipientes) Then .Recipients = pRecipientes
        If .Display() Then
            For iRecp = 1 To .Recipients.count
                If .Recipients.item(iRecp).resolved Then
                    Call colNomeOK.Add(.Recipients.item(iRecp).Name)
                    nome = nome & ";" & .Recipients.item(iRecp).Name
                Else
                    Call colNomeErro.Add(.Recipients.item(iRecp).Name)
                End If
            Next
            nome = VBA.Mid(nome, 2)
        End If
    End With
    SelecionarNomes = VBA.Array(nome, colNomeOK, colNomeErro)
Exit Function
ErrorHandler:
If VBA.Err <> 0 Then
    If VBA.Err = -2147467259 Then
        Resume Next
    Else
        VBA.MsgBox "Ocorreu um erro inesperado !" & vbNewLine & vbNewLine & "SelecionarNomes() - Erro : " & VBA.Error, VBA.vbCritical, "Erro"
    End If
    Exit Function
End If
Resume
End Function

