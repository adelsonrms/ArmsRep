Attribute VB_Name = "AuxFileSystem"
'---------------------------------------------------------------------------------------
' PROJETO      : ConsolidadoSQLServer
' MÓDULO       : AuxFileSystem.
' TIPO         : Módulo
' DATA - HORA  : 10/3/2011 - 17:37
' AUTOR        : Mondial Tecnologia em Informática LTDA
'                  Website: www.mondial.com.br
' CONSULTOR    : Adelson Rosendo Marques da Silva
'                  Email:   adelson.silva@mondial.com.br
' COMENTÁRIOS  : Contém funções genericas para uso do Shell (Arquivos, Pastas, etc)
'
'---------------------------------------------------------------------------------------
' + HISTÓRIO DE REVISÃO
'---------------------------------------------------------------------------------------
' DATA / HORA           | DESCRIÇÃO
'---------------------------------------------------------------------------------------
' 10/3/2011  - 17:37   | Incluído função que extrai o nome completo de uma pasta mapeada
'---------------------------------------------------------------------------------------
'
Option Explicit

Public Enum eFileInfo
    dataCriacao = 1
    dataModificacao = 2
    Tamanho = 3
    nome = 4
    extensao = 5
    Drive = 6
    pasta = 6
End Enum

Public Enum eSortOrder
    DIR_SO_ASC = 1
    DIR_SO_DESC = 2
End Enum


Private Type GUID
    data1      As Long
    data2      As Integer
    data3      As Integer
    data4(7)   As Byte
End Type

Private nRowStart As Long
Private nColStart As Long
Private nRowEnd As Long
Private nColEnd As Long
Private fso    As Object
Private wsh    As Object

'----------------------------------------------------------------------------------------------
' FUNÇÕES PARA CAIXA DE DIALOGO DE CORES DO SISTEMA
'----------------------------------------------------------------------------------------------
Private Type BrowseInfo
    hwndOwner  As Long
    pIDLRoot   As Long
    pszDisplayName As Long
    lpszTitle  As Long
    ulFlags    As Long
    lpfnCallback As Long
    lParam     As Long
    iImage     As Long
End Type

Private Declare PtrSafe Function CoCreateGuid Lib "OLE32.DLL" (pGuid As GUID) As Long
Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private m_CurrentDirectory As String   'The current directory

'Requer adaptação para uso com o FSO
Function getUNCFullName(strPath As String, Optional strFormat As String)
    getUNCFullName = FormatPath(strPath, , "UNC")
End Function

'[i_acd_heard]---------------------------------------------------------------------------------------
' PROCEDIMENTO     : CriarAtalho
' TIPO             : Function
' DATA/HORA        : 21/12/2009  - 14:41
' DESENVOLVEDOR    : MONDIAL TECNOLOGIA EM INFORMATICA
'                    www.mondial.com.br
' CONSULTOR        : Adelson Rosendo Marques da Silva
' DESCRIÇÃO        : Cria um atalho para um arquivo especificado
'
'---------------------------------------------------------------------------------------
' + HISTÓRIO DE REVISÃO
'---------------------------------------------------------------------------------------
' DATA / HORA           | DESCRIÇÃO
'---------------------------------------------------------------------------------------
' 21/12/2009  - 14:41   | Criação da rotina
'[f_acd_heard]---------------------------------------------------------------------------------------

Public Function CriarAtalho(sPath As String, Optional LocalDestinoDoAtalho As String, Optional DESCRIÇÃO As String, Optional vIcone) As Boolean
    Dim oNovoAtalho As Object
    Dim sNomeArquivo As String
1   On Error GoTo CriarAtalho_Err

2   If fso Is Nothing Then Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    'Instancia o processador de scripts
3   Set wsh = VBA.CreateObject("wscript.shell")
    'define o local de destino do atalho, caso não seja informado uma pasta
4   If LocalDestinoDoAtalho = "" Then
        'Usa a area de trabalho
5       LocalDestinoDoAtalho = wsh.SpecialFolders("Desktop")
6   End If
    'Extrai o nome do arquivo
7   sNomeArquivo = fso.GetBaseName(sPath)

    'Cria um objeto Shortcut um novo objeto Atalho
8   Set oNovoAtalho = wsh.CreateShortcut(LocalDestinoDoAtalho & "\" & sNomeArquivo & ".lnk")
    'define os seus parametros
9   With oNovoAtalho
10      .TargetPath = wsh.ExpandEnvironmentStrings(sPath)
11      .Description = DESCRIÇÃO
        'Se for informado um icone...
12      If Not VBA.IsMissing(vIcone) Then
13          .IconLocation = CStr(vIcone)
14      End If
        'Salva o atalho no local indicado
15      .Save
16  End With

17  CriarAtalho = VBA.Err = 0
18  Set fso = Nothing
    'Finaliza a rotina
19  On Error GoTo 0
20  Exit Function

    'Trata a ocorrencia de erros não previsíveis
CriarAtalho_Err:
21  If VBA.Err <> 0 Then
22      If VBA.MsgBox("Erro não tratado ao executar uma ação no procedimento." & VBA.Chr(10) & _
                  "Detalhes : " & VBA.Chr(10) & _
                  "Descrição do Erro : " & VBA.Error & VBA.Chr(10) & _
                  VBA.IIf(Erl <> 0, "Linha onde ocorreu o erro : " & Erl & VBA.Chr(10), "") & _
                  VBA.Chr(10) & _
                  "Rotina : CriarAtalho" & VBA.Chr(10) & _
                  "Módulo: AuxFileSystem.", VBA.vbCritical, "Erro do Sistema") = VBA.vbYes Then
            'Stop
            'Resume
23      End If
24  End If
End Function

'Função encapsulada para retornar o objeto Object
Function getFSOFile(FileName As String) As Object    ' Object
1   With VBA.CreateObject("Scripting.FileSystemObject")
2       If .FileExists(FileName) Then
3           Set getFSOFile = .GetFile(FileName)
4       End If
5   End With
End Function

Function getNomeBase(FileName As String) As String
    On Error Resume Next
1   With VBA.CreateObject("Scripting.FileSystemObject")
3       getNomeBase = .GetBaseName(FileName)
5   End With
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : AuxImportText.PegarPasta()
' TIPO             : Function
' DATA/HORA        : 01/09/2015 18:40
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Retorna o endereço da pasta mae de um arquivo ou pasta
'---------------------------------------------------------------------------------------
Function PegarPasta(pCaminho As String)
    Dim ofso As Object
    Set ofso = VBA.CreateObject("Scripting.FileSystemObject")
    PegarPasta = FormatPath(ofso.GetParentFolderName(pCaminho), ofso, "UNC")
    Set ofso = Nothing
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : AuxAplicacao.FormatPath()
' TIPO             : Function
' DATA/HORA        : 16/12/2015 15:46
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Formata  um caminho para o padrão UNC
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      16/12/2015 15:46
'---------------------------------------------------------------------------------------
Function FormatPath(sPath As String, Optional f As Object, Optional tipoFormato As String = "Normal") As String
    Dim drv As Object ' Scripting.Drive
    Dim sDrive As String

    '---------------------------------------------------------------------------------------
10  On Error GoTo FormatPath_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
20  Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Function AuxAplicacao.FormatPath()"
    '---------------------------------------------------------------------------------------

30  If tipoFormato <> "UNC" Then FormatPath = sPath: Exit Function

40  If f Is Nothing Then Set f = VBA.CreateObject("Scripting.FileSystemObject")
50  With f
60      sDrive = .GetDriveName(sPath)
70      If .DriveExists(sDrive) Then
80          Set drv = .GetDrive(sDrive)
90          If drv.DriveType = 3 Then
100             FormatPath = VBA.Replace(sPath, sDrive, drv.ShareName)
110         Else
120             FormatPath = sPath
130         End If
140     Else
150         FormatPath = sPath
160     End If
170 End With

180 Set f = Nothing
Fim:
190 On Error GoTo 0
200 Exit Function

FormatPath_Error:
210 If VBA.Err <> 0 Then
220     lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
        'Debug.Print cstr_ProcedureName, "Linha : " & VBA.Erl() & " - " & strErrorMessagem
230     Call Excecoes.TratarErro(strErrorMessagem, lngErrorNumber, cstr_ProcedureName, VBA.Erl(), , False)
240 End If
    GoTo Fim:
    'Debug Mode
250 Resume
End Function

Function GetFileInfo(FileName As String, Optional TypeInfo As eFileInfo = 1) As Variant
    Dim flObject As Object
1     Set flObject = getFSOFile(FileName)
2     If Not flObject Is Nothing Then
3       With flObject
4           Select Case TypeInfo
            Case eFileInfo.dataCriacao
5               GetFileInfo = .DateCreated
6           Case eFileInfo.dataModificacao
7               GetFileInfo = .DateLastModified
8           Case eFileInfo.Tamanho
9               GetFileInfo = .size
10          Case eFileInfo.extensao
11              GetFileInfo = VBA.CreateObject("Scripting.FileSystemObject").GetExtensionName(.size)
12          Case eFileInfo.Drive
13              GetFileInfo = .Drive.DriveLetter
14          Case eFileInfo.pasta
15              GetFileInfo = .ParentFolder.Path
16          Case eFileInfo.nome
17              GetFileInfo = .Name
18          End Select
19      End With
20    Else
21      GetFileInfo = ""
22    End If
End Function

'Exclui um arquivo passando o endereço do arquivo a ser excluido.
'Valida se :
'   1 - O arquivo a ser excluído existe
'   2 - Caso ocorre erro, mostra mensagem ou nao.
'   3 - Retorna True ou False
Function DeleteFile(sFile As String, Optional bShowMsg As Boolean = False) As Boolean
    On Error Resume Next
    If FileExists(sFile) Then
        VBA.Kill sFile
        If VBA.Err = 70 Then
            If bShowMsg Then VBA.MsgBox "O Arquivo '" & sFile & "' esta aberto. Feche-o antes de continuar." & VBA.Chr(10) & " O Processo será interrompido.", VBA.vbCritical
            Debug.Print "Erro ao excluir o arquivo '" & sFile & "'" & VBA.vbNewLine & VBA.Err & " - " & VBA.Error
            DeleteFile = False
            Exit Function
        End If
        DeleteFile = True
    End If
End Function

Function CanRename(sFilePath As String) As Boolean
    On Error Resume Next
    Name sFilePath As sFilePath
    CanRename = VBA.Err = 0
End Function

'Verifica se o arquivo pode ser aberto ou esta em uso
Function FileIsUsed(sFilePath As String) As Boolean
    On Error GoTo IsOpen
    Dim intFile As Long
    If FileExists(sFilePath) Then
        intFile = VBA.FreeFile
        Open sFilePath For Random Lock Read Write As intFile: Close intFile
        FileIsUsed = False: Exit Function
    Else
        FileIsUsed = False: Exit Function
    End If

IsOpen:
    FileIsUsed = True
End Function

Function WorkBookIsOpened(sFilePath As String) As Boolean
    On Error GoTo NotIsOpen
    Dim sFile  As String
    Dim sFolder As String
    Dim vInfo  As Variant, i

    vInfo = VBA.Split(sFilePath, "\")
    sFolder = ""
    For i = 0 To UBound(vInfo) - 1
        sFolder = sFolder & vInfo(i) & "\"
    Next

    sFile = "~$" & vInfo(UBound(vInfo))

    WorkBookIsOpened = FileExists(sFolder & sFile): Exit Function

NotIsOpen:
    WorkBookIsOpened = False
End Function

Function getWorkbookOpenFileTemp(sFilePath As String)
    On Error GoTo NotIsOpen
    Dim sFile  As String
    Dim sFolder As String
    Dim vInfo  As Variant, i

    vInfo = VBA.Split(sFilePath, "\")
    sFolder = ""
    For i = 0 To UBound(vInfo) - 1
        sFolder = sFolder & vInfo(i) & "\"
    Next

    sFile = "~$" & vInfo(UBound(vInfo))

    If FileExists(sFolder & sFile) Then
        getWorkbookOpenFileTemp = sFolder & sFile
    End If

    Exit Function

NotIsOpen:
    getWorkbookOpenFileTemp = ""

End Function

Function FileIsLocked(sFilePath As String) As Boolean
    On Error GoTo IsOpen
    Dim intFile As Long
    If Not FileExists(sFilePath) Then Exit Function
    intFile = VBA.FreeFile
    Open sFilePath For Binary Access Read Lock Read As intFile: Close intFile
    FileIsLocked = False: Exit Function
IsOpen:
    FileIsLocked = True
End Function

Function FolderDisponivel(sPath As String) As Boolean
    On Error Resume Next
    VBA.Kill sPath & "\new_folder.txt"
    VBA.Err.Clear
    Open sPath & "\new_folder.txt" For Output As #1
    Print #1, "new_folder"
    Close #1
    VBA.Kill sPath & "\new_folder.txt"
    FolderDisponivel = VBA.Err = 0
End Function

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
Function FileExists(sFileName As String) As Boolean
    On Error Resume Next
    Dim sTempFileName As String
    'Tenta primeiramente pelo Fso
    If fso Is Nothing Then Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    'Verifica se o FSO foi instanciado (por algum problema pode não instanciar)
    If Not fso Is Nothing Then
        FileExists = fso.FileExists(sFileName)
    Else
        'Se o Fso não estiver instalado, tenta a função Dir
        sTempFileName = VBA.Dir(sFileName)
        FileExists = sTempFileName <> ""
    End If
    'Limpa a variável
    Set fso = Nothing
End Function

Function FolderExists(sFileName As String) As Boolean
    On Error Resume Next
    Dim sTempFileName As String
    'Tenta primeiramente pelo Fso
    If fso Is Nothing Then Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    'Verifica se o FSO foi instanciado
    If Not fso Is Nothing Then
        FolderExists = fso.FolderExists(sFileName)
    Else
        'Se o Fso não estiver instalado, tenta a função Dir
        sTempFileName = VBA.Dir(sFileName)
        FolderExists = sTempFileName <> ""
    End If
    'Limpa a variável
    Set fso = Nothing
End Function

Sub DeleteFolder(sFileName As String)
    If fso Is Nothing Then Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    'Verifica se o FSO foi instanciado
    If fso.FolderExists(sFileName) Then fso.DeleteFolder sFileName
    'Limpa a variável
    Set fso = Nothing
End Sub

Function getCountFilesInFolder(sPath As String) As Boolean
    If fso Is Nothing Then Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    'Verifica se o FSO foi instanciado
    If fso.FolderExists(sPath) Then
        getCountFilesInFolder = fso.GetFolder(sPath).Files.count
    End If
End Function

Function GetDirectoryPath(sFileName As String) As String
    On Error Resume Next
    Dim sTempFileName As String
    'Tenta primeiramente pelo Fso
    If fso Is Nothing Then Set fso = VBA.CreateObject("Scripting.FileSystemObject")
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
' * 1.03      10/07/2017 10:10    Adelson          Identificado e corrigido problema na
'                                                  criação de pastas que contem o endereço
'                                                  UNC com o servidor começando em \\Servidor
'---------------------------------------------------------------------------------------
Function MkFullDirectory(ByVal strDir As String, Optional bSilent As Boolean = True) As Boolean
    Dim sDiretorio As String
    Dim vDiretorios As Variant
    Dim i As Integer
    Dim sRoot As String
    Dim sDemais As String

1   On Error GoTo ErrCriaPasta
2   If VBA.Left(strDir, 2) = "\\" Then
        sRoot = "\\" & VBA.Split(VBA.Mid(strDir, 3), "\")(0)
    Else
        sRoot = VBA.Split(strDir, "\")(0)
6   End If

    sDemais = VBA.Mid(strDir, Len(sRoot) + 1)
    strDir = sDemais
    
7   vDiretorios = VBA.Split(strDir, "\")
    
    On Error Resume Next
8   For i = LBound(vDiretorios) To UBound(vDiretorios)
        'If i = 0 Then sDiretorio = sRoot & "\" & vDiretorios(i)
9       sDiretorio = VBA.IIf(sDiretorio = "", vDiretorios(i), sDiretorio & "\" & vDiretorios(i))
        'Cria a pasta, caso ainda não exista
10      If VBA.Dir(sRoot & "\" & sDiretorio, VBA.vbDirectory) = "" Then
            Call VBA.MkDir(sRoot & "\" & sDiretorio)
        End If
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
'---------------------------------------------------------------------------------------
' PROJETO      : armsXLToolPack_prj
' PROCEDIMENTO : GetPath
' TIPO         : Function
' DATA - HORA  : 21/05/2008 17:15
' ULTIMA REVISÃO : 15/07/2009 12:51
' AUTOR        : Adelson Rosendo Marques da Silva
'                Email:   adelson.silva@mondial.com.br
' COMENTÁRIOS  : Localiza e retorna o caminho de um diretório por meio do objeto BroseForFolder
'---------------------------------------------------------------------------------------
'
Function getFolderPath(Optional DESCRIÇÃO As String, Optional OldValue As String, Optional bUNCFullName As Boolean = False) As String
    Dim objPasta As Object
    Dim sh     As Object
    Set sh = VBA.CreateObject("Shell.Application")
    If DESCRIÇÃO <> "" Then
        Set objPasta = sh.BrowseForFolder(0, DESCRIÇÃO, 0, 0)
    Else
        Set objPasta = sh.BrowseForFolder(0, VBA.vbNullString, 0, 0)
    End If

    If objPasta Is Nothing Then
        getFolderPath = OldValue
        Exit Function
    End If
    getFolderPath = objPasta.Self.Path
    If bUNCFullName Then getFolderPath = (objPasta.Self.Path)
End Function

'Cria um arquivo texto com um texto qualquer usandoa função nativa #print
Function createTextFile(sFile As String, strText As String) As String
    Dim intFile As Integer
    intFile = VBA.FreeFile()
    Open sFile For Output As intFile
    Print #intFile, strText
    Close #intFile
    createTextFile = sFile
End Function

'Abre um arquivo qualquer no Notepad usando o Shell nativo VBA
Function OpenFileInNotepad(sFile As String)
    If FileExists(sFile) Then
        VBA.Shell "NOTEPAD.EXE " & sFile, VBA.vbNormalFocus:
    Else
        VBA.MsgBox "O Arquivo solicitado não foi encontrado. Pode ter sido excluído ou renomeado !", VBA.vbCritical, "Não encontrado"
    End If
End Function

'Abre um arquivo qualquer no Notepad usando o ShellExecute
Function OpenFile(sFile As String)
    Dim r      As Long
    If FileExists(sFile) Then
        r = ShellExecute(0, "OPEN", sFile, "", "", 0)
    Else
        VBA.MsgBox "O Arquivo solicitado não foi encontrado. Pode ter sido excluído ou renomeado !", VBA.vbCritical, "Não encontrado"
    End If
End Function

'Especificação conforme : https://ss64.com/nt/dir.html
'   [sorted]   Sorted by /O[:]sortorder
'   /O:N   Name                  /O:-N   Name
'   /O:S   file Size             /O:-S   file Size
'   /O:E   file Extension        /O:-E   file Extension
'   /O:D   Date & time           /O:-D   Date & time
'   /O:G   Group folders first   /O:-G   Group folders last
'   several attributes can be combined e.g. /O:GEN
'
'   [time] /T:  the time field to display & use for sorting
'
'   /T:C   Creation
'   /T:A   Last Access
'   /T:W   Last Written (default)
Function GetFiles(sPath As String, sFiltro As String, Optional ClassificarPorColuna As eFileInfo = eFileInfo.nome, Optional OrdemClassificacao As eSortOrder = eSortOrder.DIR_SO_ASC) As Collection
    Dim objShell As Object ' IWshRuntimeLibrary.WshShell
    Dim objStdOut As Object 'IWshRuntimeLibrary.WshExec
    Dim TextOutput As Object 'IWshRuntimeLibrary.TextStream
    
    Dim rLinha As String
    Dim vResultado
    Dim vLinha
    Dim vFInfo
    Dim strTemp As String
    
    'Montagem do comando
    Dim sCmdDIR As String
    Dim sOrderSort As String
    Dim sSortColumn As String
    Dim strSortData As String
    
    Set GetFiles = New Collection
    If sPath = "" Then Exit Function
    
    strTemp = VBA.Environ("TEMP") & "\~LIST_FILES_" & VBA.Format(VBA.Now, "yyyy-mm-dd_hh-nn-ss") & ".TXT"
    'Local da pasta
    sCmdDIR = "CMD /C "
    
    'Parametros
    sCmdDIR = sCmdDIR & "DIR /Q /A:-D" 'Sem Diretorios
    
    Select Case ClassificarPorColuna
    Case eFileInfo.nome
        sSortColumn = "N"
    Case eFileInfo.dataCriacao
        sSortColumn = "D"
        strSortData = "/T:A"
    Case eFileInfo.dataModificacao
        sSortColumn = "D"
        strSortData = "/T:W"
    End Select
    
    If OrdemClassificacao = DIR_SO_DESC Then sOrderSort = "-"
    
    sCmdDIR = sCmdDIR & " /O:" & sOrderSort & sSortColumn 'Classificação
    'DIR "O:\5900-RISCO\AREA COMUM\homolog_qlik" /A:-D /O:-D /T:A
    'Filtros por ocorrencia
    sCmdDIR = sCmdDIR & VBA.Space(1) & VBA.Chr(34) & sPath & "\" & "*" & sFiltro & "*" & VBA.Chr(34)
    
'    Call VBA.Shell(sCmdDIR & " >" & strTemp, VBA.vbHide)
    'GoSub ObtemTexto:
    'Do While UBound(vResultado) = -1: GoSub ObtemTexto: Loop
    
    Set objShell = VBA.CreateObject("WScript.Shell")
    With objShell
        Set objStdOut = .Exec(sCmdDIR)
        Set TextOutput = objStdOut.StdOut
    End With
    
    vResultado = VBA.Split(TextOutput.ReadAll, VBA.vbNewLine)
    
    
    For Each vLinha In vResultado
        rLinha = VBA.Trim(VBA.CStr(vLinha))
        If rLinha <> VBA.vbNullString Then
            If VBA.IsDate(VBA.Left(rLinha, 10)) Then
                vFInfo = VBA.Array(sPath & "\" & VBA.Trim(VBA.Mid(rLinha, 60)), _
                                   VBA.CDate(VBA.Left(rLinha, 17)), _
                                   VBA.CDbl(VBA.Mid(rLinha, 18, 18)), _
                                   VBA.Trim(VBA.Mid(rLinha, 36, 24)), _
                                   VBA.Trim(VBA.Mid(rLinha, 60)))
                Call GetFiles.Add(vFInfo)
            End If
        End If
    Next
    
    Do While FileIsLocked(strTemp): Loop
    Call DeleteFile(strTemp)
    
    Exit Function
    
ObtemTexto:
    vResultado = VBA.Split(AuxFileSystem.GetTextFromFile(strTemp), VBA.vbNewLine)
    Return
End Function

'Recupera o nome do procedimento onde a seleção esta
Function ActiveProcedureName(Optional sCompenentName As String, _
                             Optional bFullName As Boolean = False) As String
    ActiveProcedureName = sCompenentName
    On Error Resume Next
    Dim oCM    As Object
    If sCompenentName = "" Then sCompenentName = Access.Application.VBE.SelectedVBComponent.Name
    Set oCM = Access.Application.VBE.ActiveVBProject.VBComponents(sCompenentName).CodeModule
    If VBA.Err <> 0 Then
        ActiveProcedureName = "(Restrição de acesso ao modelo de objetos VBE está ativada."
    Else
        Call oCM.CodePane.GetSelection(nRowStart, nColStart, nRowEnd, nColEnd)
        If bFullName Then
            ActiveProcedureName = CurrentProject.Name & "!" & CurrentProject.vbProject.Name & "." & sCompenentName & "." & oCM.ProcOfLine(nRowStart, 0) & "()"
        Else
            ActiveProcedureName = oCM.ProcOfLine(nRowStart, 0)
        End If
    End If
End Function

Function PathIsFolder(sPath As String) As Boolean
    On Error Resume Next
    PathIsFolder = (VBA.GetAttr(sPath) And VBA.vbDirectory) <> 0
End Function

'As funções a seguir são usadas para compactar um endereço
Function CompactPath(sPath As String, Optional intTamanhoSpace As Integer = 40)
    On Error Resume Next
    If Len(sPath) > intTamanhoSpace Then
        sPath = VBA.Left(sPath, intTamanhoSpace - Len("..." & VBA.Dir(sPath, VBA.vbDirectory))) & "\...\" & VBA.Dir(sPath, VBA.vbDirectory)
    End If
    CompactPath = sPath
End Function

'------------------------------------------------------------------------------------------------------------------------------------
'#### FUNÇÕES QUE UTILIZAM O FSO ####
'------------------------------------------------------------------------------------------------------------------------------------
Function getFolderSH(sPath As String) As Object ' Scripting.Folder
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    If PathIsFolder(sPath) Then
        If fso.FolderExists(sPath) Then Set getFolderSH = fso.GetFolder(sPath)
    Else
        If fso.FileExists(sPath) Then Set getFolderSH = fso.GetFile(sPath).ParentFolder
    End If
End Function

Function CopiarArquivoParaPasta(pArquivoOrigem As String, _
                                pPastaDestino As String, _
                                Optional pNovoNome As String, _
                                Optional pMover As Boolean = False, _
                                Optional temp_file As Boolean) As String
10  On Error GoTo CopiarArquivoParaPasta_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
20  Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Function AuxFileSystem.CopiarArquivoParaPasta()"
    '----------------------------------------------------------------------------------------------------
    Dim fs As FileSystemObject, oFile As File
    Dim strArquivoDestino As String
    
    Set fs = VBA.CreateObject("Scripting.FileSystemObject")
    If fs.FileExists(pArquivoOrigem) Then
        Set oFile = fs.GetFile(pArquivoOrigem)
        If VBA.Right(pPastaDestino, 1) = "\" Then pPastaDestino = VBA.Mid(pPastaDestino, 1, Len(pPastaDestino) - 1)
        'Endereço do arquivo de destino
        strArquivoDestino = pPastaDestino & "\" & VBA.IIf(pNovoNome <> "", pNovoNome, oFile.Name)
        'Move ou Copia
        If pMover Then
            Call oFile.Move(strArquivoDestino)
        Else
            Call oFile.Copy(strArquivoDestino, True)
        End If
        'Sucesso ou nao
        If fs.FileExists(strArquivoDestino) Then CopiarArquivoParaPasta = strArquivoDestino
    End If
    Set fs = Nothing
    
    Exit Function
    
CopiarArquivoParaPasta_Error:
    If VBA.Err <> 0 Then
        Debug.Print "Erro ao copiar o arquivo para pasta de destino : " & VBA.Err.Description
    End If
End Function
'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : AuxFileSystem.CopiarPastaPara()
' TIPO             : Function
' DATA/HORA        : 15/06/2015 16:17
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Copia todo o conteúdo de uma pasta para outra
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      15/06/2015 16:17    Adelson         Criação/Atualização do procedimento
'---------------------------------------------------------------------------------------
Function CopiarPastaPara(origem As String, destino As String, Optional bExcluirOrigem As Boolean = False) As String
    Dim fs As Object, oFile As Object
10  On Error GoTo CopiarPastaPara_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
20  Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Function AuxFileSystem.CopiarPastaPara()"
    '----------------------------------------------------------------------------------------------------
30  Set fs = VBA.CreateObject("Scripting.FileSystemObject")
40  If fs.FolderExists(origem) Then
50      Set oFile = fs.GetFolder(origem)
60      Call oFile.Copy(destino, True)
70      If bExcluirOrigem Then fs.DeleteFolder origem, True
80  End If
90  Set fs = Nothing

Fim:
100 On Error GoTo 0
110 Exit Function

CopiarPastaPara_Error:
120 If VBA.Err <> 0 Then
130     lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
140     Call Excecoes.TratarErro(strErrorMessagem, lngErrorNumber, cstr_ProcedureName, VBA.Erl)
150 End If
    GoTo Fim:
    'Debug Mode
160 Resume
End Function
'------------------------------------------------------------------------------------------------------------------------------------
'#### FUNÇÕES QUE UTILIZAM O FSO ####
'------------------------------------------------------------------------------------------------------------------------------------
Function ComapctarEnderecoCompleto(pasta As String, NomeArquivo As String, Optional TamanhoTotal As Integer = 40)
    On Error Resume Next
    If Right(pasta, 1) = "\" Then pasta = Mid(pasta, 1, Len(pasta) - 1)
    If Len(pasta & NomeArquivo) > TamanhoTotal Then
        ComapctarEnderecoCompleto = VBA.Left(pasta, TamanhoTotal - Len("...\" & NomeArquivo)) & "...\" & NomeArquivo
    Else
        ComapctarEnderecoCompleto = pasta & "\" & NomeArquivo
    End If
    If VBA.Err <> 0 Then ComapctarEnderecoCompleto = pasta & "\" & NomeArquivo
End Function

'Retorna o nome do usuário
Function GetComputerName() As String
    'Tenta recupera o usuário pela Variável de Ambiente
    GetComputerName = Environ("ComputerName")
    'Se não consegui, deixa quieto
End Function
'
'------------------------------------------------------------------------------------------------------------------------
' DATA/HORA        : 14/3/2012 - 15:00
' AUTOR            : Adelson Rosendo Marques da Silva
' OBJETIVO         : Exclui varios arquivos de uma só vez
'------------------------------------------------------------------------------------------------------------------------
Sub DelBatchFiles(strPathFolder As String, Optional strFilter As String)
    On Error Resume Next
    If FolderExists(strPathFolder) Then Call VBA.Kill(strPathFolder & "\" & strFilter & "*.*")
End Sub

'------------------------------------------------------------------------------------------------------------------------
' DATA/HORA        : 14/3/2012 - 15:00
' AUTOR            : Adelson Rosendo Marques da Silva
' OBJETIVO         : Abre a caixa de dialogo BrowseForFolder com um diretório selecionado
'------------------------------------------------------------------------------------------------------------------------
Public Function BrowseForFolder(Title As String, StartDir As String) As String
'    Dim lpIDList As Long
'    Dim szTitle As String
'    Dim sBuffer As String
'    Dim tBrowseInfo As BrowseInfo
'
'    m_CurrentDirectory = StartDir & VBA.vbNullChar
'    szTitle = Title
'
'    With tBrowseInfo
'        .hwndOwner = 0
'        .lpszTitle = lstrcat(szTitle, "")
'        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN + BIF_STATUSTEXT
'        .lpfnCallback = GetAddressofFunction(AddressOf BrowseCallbackProc)
'    End With
'
'    lpIDList = SHBrowseForFolder(tBrowseInfo)
'    If (lpIDList) Then
'        sBuffer = VBA.Space(MAX_PATH)
'        SHGetPathFromIDList lpIDList, sBuffer
'        sBuffer = VBA.Left(sBuffer, InStr(sBuffer, VBA.vbNullChar) - 1)
'        BrowseForFolder = sBuffer
'    Else
'        BrowseForFolder = ""
'    End If

End Function



'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : AuxFileSystem.CaixaDeDialogo()
' TIPO             : Function
' DATA/HORA        : 06/10/2015 20:02
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Função generica para exibir uma caixa de dialog de seleção de arquivos e pastas.
'                    -------------------------------------
'                    Rerefencias necessárias :
'                    -------------------------------------
'                    - Biblioteca : Office
'                    - Objeto     : FileDialog (variavel dlg)
'                    - Objeto     : FileDialogSelectedItems (variavel arquivo)
'                    - Constante  : msoFileDialogViewProperties (3)
'                    - Constante  : msoFileDialogFolderPicker (4)
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      06/10/2015 20:02
' * 1.02      04/10/2016 10:40    Adelson Silva    Convertido as constantes em valores
'---------------------------------------------------------------------------------------
'
Function CaixaDeDialogo(msoTipoDialog As OFFICE_MsoFileDialogType, _
                        Optional pTitulo As String, _
                        Optional pSelecaoMultiplaArquivos As Boolean = False, _
                        Optional TextoBotao As String, _
                        Optional Filtros As Variant = "Arquivos de Texto;*.txt*|Todos os Arquivos;*.*", _
                        Optional pDiretorioInicial As String) As Variant

'---------------------------------------------------------------------------------------
10  On Error GoTo CaixaDeDialogo_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
20  Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Function AuxFileSystem.CaixaDeDialogo()"
    '---------------------------------------------------------------------------------------
    Dim dlg         As Object
    Dim vFiltros    As Variant, i As Integer
    Dim arquivo     As Object
    Dim vFile       As Variant

30  Set dlg = Access.Application.FileDialog(msoTipoDialog)

40  With dlg
50      .Title = pTitulo
60      .AllowMultiSelect = pSelecaoMultiplaArquivos
70      .ButtonName = TextoBotao
80      .Filters.Clear
90      .InitialView = 1
100     .InitialFileName = pDiretorioInicial
110     If msoTipoDialog <> 4 Then

120         vFiltros = VBA.Split(Filtros, "|")
130         For i = LBound(vFiltros) To UBound(vFiltros)
140             Call dlg.Filters.Add(VBA.Split(vFiltros(i), ";")(0), VBA.Split(vFiltros(i), ";")(1))
150         Next

160     End If
170     Call .show
180     If .SelectedItems.count > 0 Then
190         Set arquivo = .SelectedItems
            If pSelecaoMultiplaArquivos Then
200             ReDim vFile(arquivo.count - 1)
210             For i = 0 To arquivo.count - 1
220                 vFile(i) = arquivo(i + 1)
230             Next
            Else
                vFile = arquivo.item(1)
            End If
240         .Filters.Clear
250     End If
260 End With
270 CaixaDeDialogo = vFile

Fim:
280 On Error GoTo 0
290 Exit Function

CaixaDeDialogo_Error:
300 If VBA.Err <> 0 Then
310     lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
320     Call Excecoes.TratarErro(strErrorMessagem, lngErrorNumber, cstr_ProcedureName, VBA.Erl(), True, False)
330 End If
    GoTo Fim:
    'Debug Mode
340 Resume
350 Set dlg = Nothing
End Function
'------------------------------------------------------------------------------------------------------------------------
' DATA/HORA        : 14/3/2012 - 15:00
' AUTOR            : Adelson Rosendo Marques da Silva
' OBJETIVO         : Gera um GUID. Códificação aleatoria do Windows
'------------------------------------------------------------------------------------------------------------------------
Public Function getNewGUID(Optional intTamanho As Integer = 36, Optional bIncluirSeparador As Boolean = False) As String
    Dim udtGUID As GUID
    If (CoCreateGuid(udtGUID) = 0) Then
        getNewGUID = _
        VBA.String(8 - VBA.Len(VBA.Hex$(udtGUID.data1)), "0") & Hex$(udtGUID.data1) & "-" & _
                     VBA.String(4 - Len(Hex$(udtGUID.data2)), "0") & Hex$(udtGUID.data2) & _
                     VBA.String(4 - Len(Hex$(udtGUID.data3)), "0") & Hex$(udtGUID.data3) & "-" & _
                     VBA.IIf((udtGUID.data4(0) < &H10), "0", "") & Hex$(udtGUID.data4(0)) & _
                     VBA.IIf((udtGUID.data4(1) < &H10), "0", "") & Hex$(udtGUID.data4(1)) & _
                     VBA.IIf((udtGUID.data4(2) < &H10), "0", "") & Hex$(udtGUID.data4(2)) & _
                     VBA.IIf((udtGUID.data4(3) < &H10), "0", "") & Hex$(udtGUID.data4(3)) & "-" & _
                     VBA.IIf((udtGUID.data4(4) < &H10), "0", "") & Hex$(udtGUID.data4(4)) & _
                     VBA.IIf((udtGUID.data4(5) < &H10), "0", "") & Hex$(udtGUID.data4(5)) & _
                     VBA.IIf((udtGUID.data4(6) < &H10), "0", "") & Hex$(udtGUID.data4(6)) & _
                     VBA.IIf((udtGUID.data4(7) < &H10), "0", "") & Hex$(udtGUID.data4(7))
    End If
    If Not bIncluirSeparador Then getNewGUID = VBA.Replace(getNewGUID, "-", "")
    If intTamanho > 0 Then getNewGUID = VBA.Left(getNewGUID, intTamanho)
End Function
'--------------------------------------------------------------------------------------------------------------------------------
'Autor      : Adelson Silva - Mondial Informatica (adelsons@gmail.com)
'Data/Hora  : 03/09/2013 12:10
'Descrição  : Pega a extenção de um arquivo especifico
'---------------------------------------------------------------------------------------------------------------------------------
Public Function GetFileExtention(pFileName As String)
    With VBA.CreateObject("Scripting.FileSystemObject")
        GetFileExtention = .GetExtensionName(pFileName)
    End With
End Function
'--------------------------------------------------------------------------------------------------------------------------------
'Autor      : Adelson Silva - Mondial Informatica (adelsons@gmail.com)
'Data/Hora  : 03/09/2013 12:10
'Descrição  : Abre um arquivo texto e obtem todo o seu conteudo em memoria
'---------------------------------------------------------------------------------------------------------------------------------
Public Function GetTextFromFile(pFile As String)
    Dim sContent As String, fso As Object, ts As Object
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(pFile).OpenAsTextStream(1, -2)
    On Error Resume Next
    sContent = ts.ReadAll
    ts.Close
    GetTextFromFile = sContent
End Function


