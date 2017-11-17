Attribute VB_Name = "APISupport"
Option Explicit
'---------------------------------------------------------------------------------------
' PROJETO/MODULO   : pTFW_WindowHandle.mTFW_APISupport
' TIPO             : M�dulo
' DATA/HORA        : 29/03/2017 16:43
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRI��O        : Fun�oes APIs de suporte necess�rias para interagir com as jabnelas do Windows
'---------------------------------------------------------------------------------------
'
' + Historico de Revis�o do M�dulo
' **************************************************************************************
'   Vers�o    Data/Hora             Autor           Descri�ao
'---------------------------------------------------------------------------------------
' * 1.00      29/03/2017 16:43
'---------------------------------------------------------------------------------------
Public Declare PtrSafe Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private colWindows As Collection
'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : mTFW_APISupport.EnumChildProc()
' TIPO             : Function
' DATA/HORA        : 29/03/2017 08:38
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRI��O        : Essa fun��o executa em um Loop que enumera todas as janela filhas (Child) que est�o na janela 'hWnd'
'---------------------------------------------------------------------------------------
'
' + Historico de Revis�o
' **************************************************************************************
'   Vers�o    Data/Hora           Autor           Descri�ao
'---------------------------------------------------------------------------------------
' * 1.00      29/03/2017 08:38
'---------------------------------------------------------------------------------------
Public Function EnumChildProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
    Dim objWindow As cTFW_Window
    '---------------------------------------------------------------------------------------
10  On Error GoTo EnumChildProc_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
20  Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Function mTFW_APISupport.EnumChildProc()"
    '---------------------------------------------------------------------------------------
30  Set objWindow = New cTFW_Window
40  Call objWindow.GetWindow(VBA.CLng(hwnd))
50  Call objWindow.LoadProperties

60  If colWindows.count = 0 Then
70      Call colWindows.Add(objWindow, "1")
80  Else
90      Call colWindows.Add(objWindow, VBA.CStr(hwnd))
100 End If
110 EnumChildProc = 1    'Continua a Enumara��o
Fim:
120 On Error GoTo 0
130 Exit Function

EnumChildProc_Error:
140 If Err <> 0 Then
150     lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
160 End If
    GoTo Fim:
    'Debug Mode
170 Resume
End Function
'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : mTFW_APISupport.GetWindows()
' TIPO             : Function
' DATA/HORA        : 29/03/2017 16:44
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRI��O        : Carrega uma Collection de todas as janelas (criando instancias da class cWindow) filhas
'---------------------------------------------------------------------------------------
'
' + Historico de Revis�o
' **************************************************************************************
'   Vers�o    Data/Hora           Autor           Descri�ao
'---------------------------------------------------------------------------------------
' * 1.00      29/03/2017 16:44
'---------------------------------------------------------------------------------------
Function GetWindows(lngHWnd As Long) As Collection
    Dim lngRet As Long
    '---------------------------------------------------------------------------------------
10  On Error GoTo GetWindows_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
20  Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Function mTFW_APISupport.GetWindows()"
    '---------------------------------------------------------------------------------------
30  Set colWindows = New VBA.Collection
40  lngRet = APISupport.EnumChildWindows(lngHWnd, AddressOf APISupport.EnumChildProc, ByVal 0&)
50  Set GetWindows = colWindows
Fim:
60  On Error GoTo 0
70  Exit Function
GetWindows_Error:
80  If Err <> 0 Then
90      lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
100 End If
    GoTo Fim:
    'Debug Mode
110 Resume
End Function
