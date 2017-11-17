VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cUtil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const WM_CLOSE = &H10
Private Const conHwndTopmost = -1
Private Const conSwpNoActivate = &H10
Private Const conSwpShowWindow = &H40
Private Const HWND_BOTTOM = 1
Private Const SWP_NOSIZE = &H1
Private Const SWP_DRAWFRAME = &H20
Private Const SWP_HIDEWINDOW = &H80
Private Const SW_HIDE = 0
Private Const SW_MAXIMIZE = 3
Private Const SW_SHOW = 5
Private Const SW_MINIMIZE = 6
Private Const SW_RESTORE = 9

Private Declare PtrSafe Function FindWindow% Lib "User" (ByVal lpClassName As Any, ByVal lpWindowName As Any)
Private Declare PtrSafe Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare PtrSafe Function GetClassName Lib "user32.dll" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Private Declare PtrSafe Function ShowWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare PtrSafe Function SetWindowText Lib "user32.dll" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare PtrSafe Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private OldText As String

Function RetHandle(Str As String) As Long
    Dim i As Integer
    Dim TempStr As String
    
    For i = 1 To Len(Str)
       TempStr = Right(Str, i)
       If Left(TempStr, 1) = " " Then
          ' found the space
          RetHandle = CLng(LTrim(TempStr))
          Exit Function
       End If
    Next i
End Function

Sub OcultarJanela(hwd As Long)
    Call ShowWindow(hwd, 6)
End Sub

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : AuxFileSystem.MkFullDirectory()
' TIPO             : Function
' DATA/HORA        : 15/02/2016 15:51
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Cria toda a estrutura de pastas de um diretório informado
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      15/02/2016 15:51
' * 1.02      05/10/2016 10:20    Adelson          Redução, organização e otimização
'---------------------------------------------------------------------------------------
Function MkFullDirectory(strDir As String, Optional bSilent As Boolean = True) As Boolean
    Dim sDiretorio As String
    Dim vDiretorios As Variant
    Dim i As Integer
    
    strDir = VBA.Replace(strDir, "/", "\")
    
1   On Error GoTo ErrCriaPasta
2   If VBA.Left(strDir, 2) = "\\" Then
3       Call VBA.MkDir(strDir)
4       MkFullDirectory = VBA.Err = 0
5       Exit Function
6   End If
7   vDiretorios = VBA.Split(strDir, "\")
8   For i = LBound(vDiretorios) To UBound(vDiretorios)
9       sDiretorio = VBA.IIf(sDiretorio = "", vDiretorios(i), sDiretorio & "\" & vDiretorios(i))
        'Cria a pasta, caso ainda não exista
10      If VBA.Dir(sDiretorio, VBA.vbDirectory) = "" Then Call VBA.MkDir(sDiretorio)
11  Next
12  If VBA.Dir(sDiretorio, VBA.vbDirectory) = "" Then
13      If Not bSilent Then VBA.MsgBox "Criada com sucesso !", VBA.vbInformation
14      MkFullDirectory = True
15  Else
16      If Not bSilent Then VBA.MsgBox "Erro ao tentar cria a estrutura de pasta!" & VBA.Chr(10) & sDiretorio, VBA.vbCritical
17      MkFullDirectory = False
18  End If

19  Exit Function

ErrCriaPasta:
20  If VBA.Err <> 0 Then
21      If Not bSilent Then
22          VBA.MsgBox "Erro ao criar a pasta : " & strDir, VBA.vbCritical, "Shell"
23          MkFullDirectory = False
24          Exit Function
25          Resume
26      End If
27  End If
End Function
