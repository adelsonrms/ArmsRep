VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTFW_Window"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' PROJETO/MODULO   : pTFW_WindowHandle.cTFW_Window
' TIPO             : Módulo de classe
' DATA/HORA        : 29/03/2017 08:42
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Concentra ações especificas de uma janela do Windows
'---------------------------------------------------------------------------------------
' + Historico de Revisão do Módulo
' **************************************************************************************
'   Versão    Data/Hora             Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      29/03/2017 08:42
'---------------------------------------------------------------------------------------
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type POINTAPI
    X As Long
    y As Long
End Type

Private Type WINDOWPLACEMENT
    Length As Long
    flags As Long
    showCmd As Long
    ptMinPosition As POINTAPI
    ptMaxPosition As POINTAPI
    rcNormalPosition As RECT
End Type

'Suporte para Office 64 Bits
Private m_sDimension As String
#If VBA7 Then
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

    Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
    Private Declare PtrSafe Function GetDlgItem Lib "user32.dll" (ByVal hDlg As Long, ByVal nIDDlgItem As Long) As Long
    Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
    Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
    Private Declare PtrSafe Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
    Private Declare PtrSafe Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
    Private Declare PtrSafe Function GetActiveWindowA Lib "user32" Alias "GetActiveWindow" () As Long
    Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
    Private Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal y As Long) As Long
    Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
#Else
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
    Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
    Private Declare Function GetDlgItem Lib "user32.dll" (ByVal hDlg As Long, ByVal nIDDlgItem As Long) As Long
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
    Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
    Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
    Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
    Private Declare Function GetActiveWindowA Lib "user32" Alias "GetActiveWindow" () As Long
    Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
    Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal y As Long) As Long
    Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

    Private Declare Function IsChild Lib "user32" (ByVal hWndParent As Long, ByVal hwnd As Long) As Long

#End If

Private Const MK_LBUTTON = &H1
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const SW_MINIMIZE = 6, SW_NORMAL = 1, SW_MAXIMIZE = 3, SW_RESTORE = 9
Private Const WM_RBUTTONDBLCLK = &H206 'Double-click
Private Const WM_RBUTTONDOWN = &H204 'Button down
Private Const WM_RBUTTONUP = &H205 'Button up

Private Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Private Declare Function SetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Sub SetCursorPosA Lib "user32" Alias "SetCursorPos" (ByVal X As Integer, ByVal y As Integer)
Private Declare Sub SetWindowPosA Lib "user32" Alias "SetWindowPos" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Sub mouse_event Lib "user32.dll" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


'~~> Constants for pressing left button of the mouse
Private Const MOUSEEVENTF_LEFTDOWN As Long = &H2
'~~> Constants for Releasing left button of the mouse
Private Const MOUSEEVENTF_LEFTUP As Long = &H4
 

Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Private Const SW_SHOWMAXIMIZED = 3

'Constantes necessárias à algumas ações
Private Const EM_REPLACESEL = &HC2
'Ação do Click
Private Const BM_CLICK As Long = &HF5&
'Envia um texto
Private Const WM_SETTEXT = &HC
Private Const GW_CHILD = 5
Private Const BM_SETCHECK As Long = &HF1&
Private Const BST_CHECKED As Long = &H1&
Private Const TCM_SETCURFOCUS As Long = &H1330&

'Variaveis locais
Private m_lhWnd As Long
Private m_sCaption As String
Private m_sClassName As String
Private m_objChildren As Collection
Private lngRet As Long

'Enum para o tipo de Botão
Public Enum DefualtButton
    OKButton = 1
    CancelButton = 2
    OtherButton = 3
End Enum

Private Sub Class_Terminate()
10  Set m_objChildren = Nothing
End Sub


Private Function NewWindow(Optional pCaption As String)
10  m_lhWnd = FindWindow(VBA.vbNullString, pCaption)
End Function
'------------------------------------------------------------------
'Atualiza a coleção com todas as janela filhas
'------------------------------------------------------------------
Public Function LoadChildren(pWindows As Collection)
10  Set m_objChildren = pWindows
End Function
'------------------------------------------------------------------
'Recupera a CLASSNAME (tipo da janela)
'------------------------------------------------------------------
Public Function PegarClass(Optional lngWnd As Long) As String
    Dim RetVal As Long
    Dim lpClassName As String
10  If lngWnd = 0 Then lngWnd = Me.hwnd
20  lpClassName = Space(256)
30  RetVal = GetClassName(lngWnd, lpClassName, 256)
40  PegarClass = VBA.Left$(lpClassName, RetVal)
End Function
'------------------------------------------------------------------
'Define o handle da janela a ser analisada
'------------------------------------------------------------------
Public Function GetWindow(Optional pCaptionOrHandle) As Long
    Dim lResult As Long
10  If VBA.IsNumeric(pCaptionOrHandle) Then
20      lResult = VBA.CLng(pCaptionOrHandle)
30  Else
40      Me.Caption = VBA.CStr(pCaptionOrHandle)
50      lResult = FindWindow(VBA.vbNullString, VBA.CStr(pCaptionOrHandle))
60  End If
    Me.hwnd = lResult
70  GetWindow = lResult
End Function

'------------------------------------------------------------------
'Define o handle da janela a ser analisada
'------------------------------------------------------------------
Public Function GetByClassName(Optional pClassName) As Long
    Dim lResult As Long
10  If VBA.IsNumeric(pClassName) Then
20      lResult = VBA.CLng(pClassName)
30  Else
50      lResult = FindWindowEx(ByVal 0&, ByVal 0&, VBA.CStr(pClassName), VBA.vbNullString)
60  End If
    Me.hwnd = lResult
70  GetByClassName = lResult
End Function

Function GetChild(pClassName As String) As cTFW_Window
    Dim lResult As Long
    Dim oNewWindow As New cTFW_Window
    Dim strText As String
    
    lResult = getWindowInfo(Me.hwnd, pClassName)
    oNewWindow.hwnd = lResult
    strText = VBA.String$(100, Chr$(0))
    lngRet = GetClassName(oNewWindow.hwnd, strText, 100)
    oNewWindow.ClassName = VBA.Left$(strText, lngRet)
    
    strText = VBA.String$(100, Chr$(0))
    lngRet = GetWindowText(lResult, strText, 100)
    
    If lngRet > 0 Then
        oNewWindow.Caption = VBA.Left$(strText, lngRet)
    Else
        oNewWindow.Caption = "N/A"
    End If
    Set GetChild = oNewWindow
End Function

Private Function getWindowInfo(lParent As Long, pClassName As String) As Long
    Dim lResult As Long
    Dim strText As String
    lResult = FindWindowEx(ByVal lParent, ByVal 0&, VBA.vbNullString, VBA.vbNullString)
    While lResult <> 0
        strText = VBA.String$(100, Chr$(0))
        lngRet = GetClassName(lResult, strText, 100)
        If VBA.Left$(strText, lngRet) = pClassName Then
            getWindowInfo = lResult
            GoTo Fim:
        Else
                lResult = FindWindowEx(ByVal lParent, ByVal lResult, VBA.vbNullString, VBA.vbNullString)
        End If
    Wend
Fim:
End Function



'------------------------------------------------------------------
'Envia o comando de CLICK para um botão na janela. Identifica o HANDLE do botão atraves da função GetDlgItem
'------------------------------------------------------------------
Function ClickOn(pButton As DefualtButton, Optional hWndButton As Long) As Long
10  Select Case pButton
    Case DefualtButton.OKButton
20      hWndButton = GetDlgItem(Me.hwnd, &H1&)
30  Case DefualtButton.CancelButton
40      hWndButton = GetDlgItem(Me.hwnd, &H2&)
50  Case DefualtButton.OtherButton
60      hWndButton = hWndButton
70  End Select
80  Call validaRetorno(hWndButton, "GetDlgItem()")
90  If hWndButton <> 0 Then
100     lngRet = SetFocusAPI(hWndButton)
110     ClickOn = APIMessage(hWndButton, BM_CLICK)
120 End If
End Function

Function Activate() As Long
    Call ShowWindow(m_lhWnd, SW_SHOWMAXIMIZED)
    Call SetWindowPosA(m_lhWnd, 0, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE)
    Call SetForegroundWindow(m_lhWnd)
End Function


Sub MouseClick(PosX As Long, PosY As Long, Optional pWaitClick As Long = 50)
    Call SetCursorPosA(PosX, PosY)
    Call mouse_event(MOUSEEVENTF_LEFTDOWN, PosX, PosY, 0, 0)
    Call mouse_event(MOUSEEVENTF_LEFTUP, PosX, PosY, 0, 0)
    Call Sleep(pWaitClick) 'Aguarda 1 segundo o processamento do click
End Sub

Function Wait(Optional pTime As Long = 1000)
    Call Sleep(pTime)
End Function

Function SetFocus()
    Debug.Print SetFocusAPI(Me.hwnd)
End Function
'------------------------------------------------------------------------------------------------------------------------------------
'Localiza o primeiro elemento EDIT (caixa de edição de texto ou TEXTBOX em uma janela
'------------------------------------------------------------------------------------------------------------------------------------
Function FindButton() As cTFW_Window
    Dim wnd As cTFW_Window
    Dim index As Integer
10  For index = 1 To Me.Children.count
20      Set wnd = Me.Children(index)
30      If wnd.ClassName = "Button" Then
40          Set FindButton = Me.Children(index)
50          Exit Function
60      End If
70  Next index
End Function
'------------------------------------------------------------------------------------------------------------------------------------
'Localiza o primeiro elemento EDIT (caixa de edição de texto ou TEXTBOX em uma janela
'------------------------------------------------------------------------------------------------------------------------------------
Function FindEdit() As cTFW_Window
    Dim wnd As cTFW_Window
    Dim index As Integer
10  For index = 1 To Me.Children.count
20      Set wnd = Me.Children(index)
30      If wnd.ClassName = "Edit" Then
40          Set FindEdit = Me.Children(index)
50          Exit Function
60      End If
70  Next index
End Function
'------------------------------------------------------------------------------------------------------------------------------------
'Salva o texto em um determinado TextBox (atraves do HANDLE)
'------------------------------------------------------------------------------------------------------------------------------------
Function SetText(Optional sText As String, Optional lEditHwnd As Long) As Long
10  If lEditHwnd = 0 Then lEditHwnd = Me.hwnd
20  SetText = APIMessage(lEditHwnd, WM_SETTEXT, CStr(sText))
End Function

Private Function APIMessage(pHwnd As Long, pMessage As Long, Optional pValue)
10  If VBA.IsMissing(pValue) Then pValue = 0&
20  lngRet = SendMessage(pHwnd, CInt(pMessage), 0, pValue)
30  Call validaRetorno(lngRet, "APIMessage > SendMessage()")
40  APIMessage = lngRet
End Function

Private Function validaRetorno(lngRt As Long, Optional InvokedBy As String, Optional pMensagem As String)
10  If lngRt <> 0 Then
20      Debug.Print "Invoked_By :" & InvokedBy & " - Return : " & lngRt & " - OK"
30  Else
40      Debug.Print "Invoked_By :" & InvokedBy & " - Return : " & lngRt & " - " & pMensagem
50  End If
End Function

Private Sub GetWindowSize()
    Dim Rec As RECT, Point As POINTAPI
    GetWindowRect Me.hwnd, Rec
    GetCursorPos Point
    m_sDimension = "Left=" & Rec.Left & " / " & "Right=" & Rec.Right & " / " & "Top=" & Rec.Top & " / " & "Bottom=" & Rec.Bottom
End Sub

Function LoadProperties()
    Me.ClassName = Me.PegarClass(hwnd)
    Call GetWindowSize
End Function

'------------------------------------------------------------------------------------------
'PROPRIEDADES
'------------------------------------------------------------------------------------------
Public Property Get hwnd() As Long
10  hwnd = m_lhWnd
End Property

Public Property Let hwnd(ByVal lhWnd As Long)
10  m_lhWnd = lhWnd
End Property

Public Property Get Caption() As String
10  Caption = m_sCaption
End Property

Public Property Let Caption(ByVal sCaption As String)
10  m_sCaption = sCaption
End Property

Public Property Get ClassName() As String
10  ClassName = m_sClassName
End Property

Public Property Let ClassName(ByVal sClassName As String)
10  m_sClassName = sClassName
End Property

Public Property Get Children() As Collection
10  Set Children = m_objChildren
End Property

Public Property Set Children(objChildren As Collection)
10  Set m_objChildren = objChildren
End Property

Public Property Get Dimension() As String
    Dimension = m_sDimension
End Property

Public Property Let Dimension(ByVal sDimension As String)
    m_sDimension = sDimension
End Property

