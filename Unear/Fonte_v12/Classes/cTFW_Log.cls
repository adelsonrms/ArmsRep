VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTFW_Log"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------------------------------------------
' PROJETO/MODULO   : TFWCliente.cLog
' TIPO             : Módulo de classe
' DATA/HORA        : 01/10/2015 17:58
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Funções relativa a manimulação de Logs
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão do Móduto
' **************************************************************************************
'   Versão    Data/Hora             Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      01/10/2015 17:58
'---------------------------------------------------------------------------------------
Public cOldValues As Variant
Public cNewValues As Variant
Private m_sFileLog As String
Private m_sPathLog As String
Private m_objFSO As Object ' FileSystemObject

Private Sub Class_Initialize()
    Set m_objFSO = VBA.CreateObject("Scripting.FileSystemObject")
    Call configuraLocalLog
End Sub

Private Sub Class_Terminate()
    Set m_objFSO = Nothing
End Sub
'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : Sessao.RegistrarLog()
' TIPO             : Sub
' DATA/HORA        : 29/04/2014 15:59
' CONSULTOR        : (ADELSON)/ TECNUN - Adelson Rosendo Marques da Silva
' CONTATO          : adelson@tecnun.com.br
' DESCRIÇÃO        : Registra uma entrada no log de eventos
'---------------------------------------------------------------------------------------
'REVISÃO
' 27/09/2015 01:44 - ARMS   Inclusão de parametro para escolher incluir ou nao o calculo de tempo no log
'---------------------------------------------------------------------------------------
Public Sub RegistrarLog(pMessage As String, _
                    Optional pFileOutPut As String = "", _
                    Optional bStartNewFile As Boolean = False, _
                    Optional bPrintDatetime As Boolean = True, _
                    Optional bDebugPrint As Boolean, _
                    Optional bCloseHifen As Boolean = False, _
                    Optional pCaracter As String = "-", _
                    Optional bShowFileInNotePad As Boolean = False, _
                    Optional bIncluirCalculoTermpo As Boolean)

    Dim NumFileLog As Integer
    Dim msgLog As String
    Dim FileNameLog As String
    Dim sTime As String

10  If VBA.Dir(CurrentProject.Path & "\Log\", VBA.vbDirectory) = "" Then VBA.MkDir CurrentProject.Path & "\Log\"

20  If pFileOutPut <> "" Then
30      FileNameLog = pFileOutPut
40  Else
50      FileNameLog = m_sFileLog
60  End If

70  If bIncluirCalculoTermpo Then
80      sTime = CalcularTempoDecorrido(VBA.Now - dtFirstTime)
90      If sTime = "00 segs" Then
100         sTime = " @ START " & sTime
110         dtLastTime = dtFirstTime
120     End If

130     msgLog = VBA.IIf(bPrintDatetime, VBA.Now & VBA.vbTab & _
                ChaveUsuario & VBA.vbTab & _
                Right(Space(16) & CalcularTempoDecorrido(VBA.Now - dtLastTime), 16) & " | " & Right(Space(16) & sTime, 16) & VBA.vbTab, "") & pMessage
140 Else
150     msgLog = VBA.IIf(bPrintDatetime, VBA.Now & VBA.vbTab & VBA.Environ("ComputerName") & "\" & VBA.Environ("Username") & VBA.vbTab, "") & pMessage
160 End If

170 NumFileLog = openRegistrarLog(FileNameLog, bStartNewFile, bPrintDatetime)

180 If bCloseHifen Then Print #NumFileLog, String(300, pCaracter)
190 Print #NumFileLog, msgLog
200 If bCloseHifen Then Print #NumFileLog, String(300, pCaracter)
210 Close #NumFileLog

220 If bDebugPrint Then Debug.Print msgLog

230 dtLastTime = Time

240 If bShowFileInNotePad Then openTextView FileNameLog

End Sub

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : Sessao.openRegistrarLog()
' TIPO             : Function
' DATA/HORA        : 29/04/2014 15:58
' CONSULTOR        : (ADELSON)/ TECNUN - Adelson Rosendo Marques da Silva
' CONTATO          : adelson@tecnun.com.br
' DESCRIÇÃO        : Abre um novo arquivo de logo
'---------------------------------------------------------------------------------------
Private Function openRegistrarLog(pathLog As String, Optional bClear As Boolean = False, Optional PrintDatetime As Boolean = True) As Integer
    On Error Resume Next
    Dim NumFile As Integer
    NumFile = VBA.FreeFile()

    'Reinicia o log caso esteja cheio
    If VBA.Dir(pathLog) <> "" Then
        If FileSystem.FileLen(pathLog) > 300000 Then
            Name pathLog As pathLog & "_" & VBA.Format(VBA.Now, "yyyymmddhhnnss") & ".txt"
            bClear = True
        End If
    End If

    If VBA.Dir(pathLog) = "" Or bClear Then
        Open pathLog For Output As #NumFile
        If PrintDatetime Then
            Print #NumFile, VBA.Left("DATA / HORA" & String(Len(VBA.Now), " "), Len(VBA.Now)) & VBA.vbTab & "MAQUINA \ USUARIO" & VBA.vbTab & "LOG GERADO PELA FERRAMENTA : " & CodeDb.Name & VBA.Now
            Print #NumFile, String(120, "-")
        End If
    Else
        Open pathLog For Append As #NumFile
    End If
    openRegistrarLog = NumFile
End Function

'Formata o tempo decorrido de um processo
Function CalcularTempoDecorrido(dtTimeToFormat As Date)
    Dim ret    As String
    If dtTimeToFormat = VBA.Now Then CalcularTempoDecorrido = "(Start)"
    If VBA.Hour(dtTimeToFormat) > 0 Then ret = VBA.Format(VBA.Hour(dtTimeToFormat), "00") & " hrs"
    If VBA.Hour(dtTimeToFormat) = 0 And VBA.Minute(dtTimeToFormat) > 0 Then ret = ret & VBA.IIf(ret <> "", " e ", "") & VBA.Format(VBA.Minute(dtTimeToFormat), "00") & " min"
    If VBA.Hour(dtTimeToFormat) = 0 And VBA.Minute(dtTimeToFormat) > 0 And VBA.Second(dtTimeToFormat) Then ret = ret & VBA.IIf(ret <> "", " e ", "") & VBA.Format(VBA.Second(dtTimeToFormat), "00") & " segs"
    If VBA.Hour(dtTimeToFormat) = 0 And VBA.Minute(dtTimeToFormat) = 0 And VBA.Second(dtTimeToFormat) Then ret = VBA.Format(VBA.Second(dtTimeToFormat), "00") & " segs"
    If ret = "" Then ret = "00 segs"
    CalcularTempoDecorrido = ret
End Function

'Exibe o conteudo texto no Notepad
Private Sub openTextView(source As String)
    If VBA.Dir(source) <> "" Then
        VBA.Shell "notepad.exe " & source, VBA.vbMaximizedFocus
    End If
End Sub

'Formata o tamanho de um arquivo em bytes
Function formatFileSize(size As Long, Optional bIncluirUnidade As Boolean = True)
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

Function FileSize(pFilePath As String) As Long
    If FileExists(pFilePath) Then FileSize = FileSystem.FileLen(pFilePath)
End Function

Function FileDateTime(pFilePath As String) As Date
    If FileExists(pFilePath) Then FileDateTime = FileSystem.FileDateTime(pFilePath)
End Function

'Determina onde o código está sendo executado. Se Excel ou Access
Function getHostPath() As Object
    On Error Resume Next
    Dim oApp   As Object
    Set oApp = Access.Application
    If oApp.Name = "Microsoft Excel" Then
        Set getHostPath = oApp.ThisWorkbook
    Else
        Set getHostPath = oApp.CurrentProject
    End If
End Function

Public Property Get FileLog() As String
    FileLog = m_sFileLog
End Property

Public Property Let FileLog(ByVal sFileLog As String)
    m_sFileLog = sFileLog
End Property

Public Property Get pathLog() As String
    pathLog = m_sPathLog
End Property

Public Property Let pathLog(ByVal sPathLog As String)
    m_sPathLog = sPathLog
End Property

Public Property Get fso() As Object ' FileSystemObject
    If m_objFSO Is Nothing Then Set m_objFSO = VBA.CreateObject("Scripting.FileSystemObject")
    Set fso = m_objFSO
End Property
'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cLog.configuraLocalLog()
' TIPO             : Function
' DATA/HORA        : 01/10/2015 17:46
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Configura a localização dos arquivos de log.
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      01/10/2015 17:46
'---------------------------------------------------------------------------------------
Private Function configuraLocalLog() As Boolean
    Dim Host As Object
    Dim sFolder As String
    Dim sName As String
    '----------------------------------------------------------------------------------------------------
10  On Error GoTo configuraLocalLog_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
20  Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Function cLog.configuraLocalLog()"
    '----------------------------------------------------------------------------------------------------

30  Set Host = getHostPath
40  If Not Host Is Nothing Then
50      sFolder = Host.Path() & "\Log"
60      sName = Host.Name & ".txt"
70  Else
80      sFolder = Environ("temp")
90      sName = VBA.Format(VBA.Now, "yyyymmdd_hhnnss") & ".txt"
100 End If
110 If Not fso.FolderExists(sFolder) Then fso.CreateFolder (sFolder)
120 m_sPathLog = sFolder
130 m_sFileLog = m_sPathLog & "\" & sName


Fim:
140 On Error GoTo 0
150 Exit Function

configuraLocalLog_Error:
160 If VBA.Err <> 0 Then
170     lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
180     Debug.Print VBA.Now, cstr_ProcedureName & "() - Ocorreu um erro inesperado nessa rotina. ERROR : " & strErrorMessagem
190 End If
    GoTo Fim:
    'Debug Mode
200 Resume
End Function
