VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTFW_Excel"
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

Private Const ProgID_App As String = "Excel.Application"

Public Enum ExcelFileFormat
    xlExcel_XLS = 56 'xls (00-03) > xlExcel8 = 56
    xlExcel_XLSB = 50 'xlsb > xlExcel12 = 50
    xlExcel_XLSX = 51 'xlsx > xlOpenXMLWorkbook = 51
    xlExcel_XLSM = 52 'xlsm > xlOpenXMLWorkbookMacroEnabled = 52
End Enum

Public Enum ExcelXlDirection
    xlDirection_Down = -4121
    xlDirection_Toleft = -4159
    xlDirection_Toright = -4161
    xlDirection_up = -4162
End Enum

Public Enum eEstiloTabelaLight
    TableStyleLight1 = 1
    TableStyleLight2 = 2
    TableStyleLight3 = 3
    TableStyleLight4 = 4
    TableStyleLight5 = 5
    TableStyleLight6 = 6
    TableStyleLight7 = 7
    TableStyleLight8 = 8
    TableStyleLight9 = 9
    TableStyleLight10 = 10
    TableStyleLight11 = 11
    TableStyleLight12 = 12
    TableStyleLight13 = 13
    TableStyleLight14 = 14
    TableStyleLight15 = 15
    TableStyleLight16 = 16
    TableStyleLight17 = 17
    TableStyleLight18 = 18
    TableStyleLight19 = 19
    TableStyleLight20 = 20
    TableStyleLight21 = 21
End Enum

Public Enum eEstiloTabelaMedium
    TableStyleMedium1 = 1
    TableStyleMedium2 = 2
    TableStyleMedium3 = 3
    TableStyleMedium4 = 4
    TableStyleMedium5 = 5
    TableStyleMedium6 = 6
    TableStyleMedium7 = 7
    TableStyleMedium8 = 8
    TableStyleMedium9 = 9
    TableStyleMedium10 = 10
    TableStyleMedium11 = 11
    TableStyleMedium12 = 12
    TableStyleMedium13 = 13
    TableStyleMedium14 = 14
    TableStyleMedium15 = 15
    TableStyleMedium16 = 16
    TableStyleMedium17 = 17
    TableStyleMedium18 = 18
    TableStyleMedium19 = 19
    TableStyleMedium20 = 20
    TableStyleMedium21 = 21
End Enum

Private Function PegarNomeEstiloTabelaLight(eEstiloLight As eEstiloTabelaLight)
    PegarNomeEstiloTabelaLight = VBA.Choose(eEstiloLight, "TableStyleLight1", "TableStyleLight2", "TableStyleLight3", "TableStyleLight4", "TableStyleLight5", "TableStyleLight6", "TableStyleLight7", "TableStyleLight8", "TableStyleLight9", "TableStyleLight10", "TableStyleLight11", "TableStyleLight12", "TableStyleLight13", "TableStyleLight14", "TableStyleLight15", "TableStyleLight16", "TableStyleLight17", "TableStyleLight18", "TableStyleLight19", "TableStyleLight20", "TableStyleLight21")
End Function

Private Function PegarNomeEstiloTabelaMedium(eEstiloMedium As eEstiloTabelaMedium)
    PegarNomeEstiloTabelaMedium = VBA.Choose(eEstiloMedium, "TableStyleMedium1", "TableStyleMedium2", "TableStyleMedium3", "TableStyleMedium4", "TableStyleMedium5", "TableStyleMedium6", "TableStyleMedium7", "TableStyleMedium8", "TableStyleMedium9", "TableStyleMedium10", "TableStyleMedium11", "TableStyleMedium12", "TableStyleMedium13", "TableStyleMedium14", "TableStyleMedium15", "TableStyleMedium16", "TableStyleMedium17", "TableStyleMedium18", "TableStyleMedium19", "TableStyleMedium20", "TableStyleMedium21")
End Function

Private Sub Class_Initialize()
  '  Call NewExcel
End Sub

Private Sub Class_Terminate()
    Set m_objApplication = Nothing
    Set m_objWorkbooks = Nothing
    Set m_objWorksheets = Nothing
    Set m_objRange = Nothing
    Set m_objThisWorkbook = Nothing
End Sub

'Inicializa uma instancia do Excel.
'Caso seja passado uma instancia ja criada, aproveita a instancia
Function AbrirExcel(Optional objInstancia As Object) As Object
    If objInstancia Is Nothing Then
        Call NewExcel
    Else
        Set m_objApplication = objInstancia
    End If
End Function

Public Property Get Application() As Object
    Set Application = m_objApplication
End Property

Public Property Set Application(objApplication As Object)
    Set m_objApplication = objApplication
End Property

'Abre uma WorkBook
Friend Function openWorkbook(Optional pFileName As String, _
                             Optional bReadOnly As Boolean = False, _
                             Optional bUpdateLinks As Boolean = True, _
                             Optional sPassword As String = VBA.vbNullString) As Object
10  On Error GoTo openWorkbook_Error

    Dim lngErrorNumber As Long, strErrorMessagem As String
20  Dim dtSartRunProc As Date: dtSartRunProc = Time
    Const cstr_ProcedureName As String = "cExcel.openWorkbook()"


30  If m_objApplication Is Nothing Then
40      Call NewExcel
50  End If

60  m_objApplication.DisplayAlerts = False

70  If pFileName <> "" Then
80      Set m_objThisWorkbook = Me.Workbooks.Open(FileName:=pFileName, UpdateLinks:=bUpdateLinks, ReadOnly:=bReadOnly, password:=sPassword)
90  Else
100     Set m_objThisWorkbook = Me.Workbooks.Add(1)
110 End If
120 Set openWorkbook = m_objThisWorkbook

130 Exit Function

openWorkbook_Error:
140 If VBA.Err <> 0 Then
150     Debug.Print cstr_ProcedureName, VBA.Err.Number & "-" & VBA.Err.Description
160     Call VBA.Err.Raise(VBA.Err.Number, cstr_ProcedureName, "Em : " & cstr_ProcedureName & " Em " & VBA.IIf(VBA.Erl() <> 0, " > (" & VBA.Erl() & ")", "") & " > " & "Ocorreu um erro na abertura do arquivo " & vbNewLine & vbNewLine & VBA.Err.Description)
170 End If
End Function

'Fecha uma Workbook
Friend Sub closeWorkbook(ByRef wkInstance As Object, Optional pSave As Boolean = False, Optional bCLoseXL As Boolean = True, Optional bManterBackup As Boolean = False)
    On Error Resume Next
    Dim xApp   As Object
    If wkInstance Is Nothing Then Set wkInstance = m_objThisWorkbook
    If wkInstance Is Nothing Then Exit Sub
    wkInstance.CheckCompatibility = False
    If bCLoseXL Then Set xApp = wkInstance.Application
    If bManterBackup Then
        Call CriarBackupWorkbook(wkInstance.FullName)
    End If
    Call wkInstance.Close(pSave)
    If bCLoseXL Then xApp.quit
    Set xApp = Nothing
    Set wkInstance = Nothing
End Sub


'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : ExcelExport.WorkbookSaveAs()
' TIPO             : Sub
' DATA/HORA        : 15/09/2014 22:18
' CONSULTOR        : (ADELSON)/ TECNUN - Adelson Rosendo Marques da Silva
' CONTATO          : adelson@tecnun.com.br
' DESCRIÇÃO        : Função Generica para salvar uma pasta de trabalho no formato adequado
'---------------------------------------------------------------------------------------
Sub WorkbookSaveAs(oWkb As Object, pFileName As String, Optional bCalculateBeforeSave As Boolean = False, Optional pPassword As String)
    Dim sExt As String, vExt As Variant

    '## MZTools Custom - Variáveis de Ambiente
    '----------------------------------------------------------------------------------------------------
    On Error GoTo WorkbookSaveAs_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
    Dim dtSartRunProc As Date: dtSartRunProc = Time
    Const cstr_ProcedureName As String = "cExcel.WorkbookSaveAs()"
    'Habilitar Log de inicio/fim da rotina (-1 (ou True) / 0 ou False) ??
    Const cEnableLog As Boolean = -1
    '----------------------------------------------------------------------------------------------------

    'vExt = VBA.Split(pFileName, ".")
    'sExt = vExt(UBound(vExt))
    sExt = GetExcelExtention(pFileName)
    
    If FileExists(pFileName) Then
        If FileIsLocked(pFileName) Then
            pFileName = PegarPasta(pFileName) & "\" & VBA.Format(VBA.Now, "YYYYMMDD-HHNNSS") & sExt
            Call VBA.MsgBox("O Arquivo esta em uso e não pode ser salvo. Um arquivo temporário será salvo no mesmo local." & VBA.vbNewLine & VBA.vbNewLine & pFileName, VBA.vbExclamation, "Arquivo em uso")
        Else
            Call VBA.Kill(pFileName)
        End If
    End If
    
    'Constantes dos tipos de acordo com a extenção.
    'xlsb > xlExcel12 = 50
    'xlsx > xlOpenXMLWorkbook = 51
    'xlsm > xlOpenXMLWorkbookMacroEnabled = 52
    'xls (00-03) > xlExcel8 = 56

    Select Case LCase(sExt)
        Case "xls"
            Call oWkb.SaveAs(FileName:=pFileName, FileFormat:=ExcelFileFormat.xlExcel_XLS, password:=pPassword)
        Case "xlsb"
            Call oWkb.SaveAs(FileName:=pFileName, FileFormat:=ExcelFileFormat.xlExcel_XLSB, password:=pPassword)
        Case "xlsm"
            Call oWkb.SaveAs(FileName:=pFileName, FileFormat:=ExcelFileFormat.xlExcel_XLSM, password:=pPassword)
        Case "xlsx"
            Call oWkb.SaveAs(FileName:=pFileName, FileFormat:=ExcelFileFormat.xlExcel_XLSX, password:=pPassword)
    End Select

    On Error GoTo 0
    Exit Sub

WorkbookSaveAs_Error:
    If VBA.Err <> 0 Then
        lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
        Call VBA.Err.Raise(VBA.Err.Number, cstr_ProcedureName, "Em : " & cstr_ProcedureName & "() Em " & VBA.IIf(VBA.Erl() <> 0, " > Linha (" & VBA.Erl() & ")", "") & " > " & strErrorMessagem)
    End If
    Exit Sub
    'Debug Mode
    Resume
End Sub
'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cExcel.saveWorkbookCopy()
' TIPO             : Function
' DATA/HORA        : 06/05/2016 15:03
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : A partir de uma instancia de um Workbook aberto, salva uma pasta de trabalho em outro lugar
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      06/05/2016 15:03
' * 1.01      13/03/2017 15:03    A instancia do Excel será recuperada a partir da
'                                 workbook de origem
' * 1.02      20/03/2017 10:03    Incluido o argumento 'bReOpenCopy' para que seja
'                                 possivel reabri a copia somente se necessário
'---------------------------------------------------------------------------------------
Function saveWorkbookCopy(FromWkb As Object, _
                          Optional FileName As String, _
                          Optional bCloseAfterSave As Boolean = False, _
                          Optional strDeleteSheetsExcept As String, _
                          Optional bReOpenCopy As Boolean = True) As Object
    
    Dim oWkbTemp As Object
    '---------------------------------------------------------------------------------------
1   On Error GoTo saveWorkbookCopy_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
2   Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "cExcel.saveWorkbookCopy()"
    '---------------------------------------------------------------------------------------

3   FileName = VBA.Replace(FileName, "/", "")
4   If FileName = "" Then
        FileName = Environ("temp") & "\" & VBA.Format(VBA.Now, "~yyymmddhhnnss") & "." & GetExcelExtention(FromWkb.Name)
    Else
    
    End If

5   If VBA.Dir(FileName) <> "" Then VBA.Kill FileName
6   Call FromWkb.SaveCopyAs(CStr(FileName))

7   If bCloseAfterSave Then
8       Call Me.closeWorkbook(FromWkb, False, True)
9   Else
        If bReOpenCopy Then
10          Set oWkbTemp = Me.openWorkbook(CStr(FileName))
11          Set saveWorkbookCopy = oWkbTemp
        End If
12  End If
    
    
Fim:
13  On Error GoTo 0
14  Exit Function

saveWorkbookCopy_Error:
15  If VBA.Err <> 0 Then
16      lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
        Debug.Print cstr_ProcedureName, "Linha : " & VBA.Erl() & " - " & strErrorMessagem
18  End If
    GoTo Fim:
    'Debug Mode
19  Resume
End Function


Friend Function NewExcel() As Object
    Set m_objApplication = VBA.CreateObject("Excel.Application")
    Set NewExcel = m_objApplication
End Function

Friend Property Get Workbooks() As Object
    Set Workbooks = m_objApplication.Workbooks
End Property

Private Property Get Worksheets() As Object
    Set Worksheets = m_objApplication.Worksheets
End Property

Friend Property Get Range() As Object
    Set Range = m_objApplication.Range
End Property

Friend Property Get ThisWorkbook() As Object
    Set ThisWorkbook = m_objThisWorkbook
End Property

Friend Property Set ThisWorkbook(objThisWorkbook As Object)
    Set m_objThisWorkbook = objThisWorkbook
End Property

Private Function GetExcelExtention(pFileName As String)
    GetExcelExtention = Right(pFileName, Len(pFileName) - VBA.InStrRev(pFileName, "."))
End Function

Sub ShowHideExcel(Optional bShow As Boolean = True)
    Me.Application.Visible = bShow
End Sub

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cExcel.ExportarParaPlanilha()
' TIPO             : Sub
' DATA/HORA        : 19/03/2015 17:35
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Transfere os dados para uma planilha
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      19/03/2015 17:35    Adelson         Criação/Atualização do procedimento
' * 2.00      11/05/2017 11:18    Adelson         Reorganização e nomeação de parametros para nomes mais intuitivos
'---------------------------------------------------------------------------------------
Public Function ExportarParaPlanilha(pOrigem As String, _
                              Optional pDestino, _
                              Optional wkbOrigem, _
                              Optional pIncluirCabecalho As Boolean, _
                              Optional pOpcoes, _
                              Optional pSalvarArquivoComo As String, _
                              Optional pExibirArquivo As Boolean = False, _
                              Optional pColarFormulas)
                              
    '----------------------------------------------------------------------------------------------------
1   On Error GoTo ExportarParaPlanilha_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String: Const cstr_ProcedureName As String = "cExcel.ExportarParaPlanilha()"
2   Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    '----------------------------------------------------------------------------------------------------
    
    Dim rngLista As Object, lst As Object
    Dim wkbDestino As Object, strPlanilha As String
    Dim sNameList As String, sLayout As String
    Dim wks As Object
    Dim bFormatPadrao As Boolean
    Dim colunasFormarComoValor As String
    Dim rngInicial As Object, strCelula As String

    '--------------------------------------------------------------------
    'Uma nova workbook será criada
    '--------------------------------------------------------------------
3   If VBA.IsMissing(wkbOrigem) Then
4       Set wkbDestino = Me.openWorkbook()
    ElseIf wkbOrigem Is Nothing Then
        Set wkbDestino = Me.openWorkbook()
5   Else
        'Utiliza a instancia de uma Workbook ja criada
6       If VBA.IsObject(wkbOrigem) And VBA.TypeName(wkbOrigem) = "Workbook" Then
7           Set wkbDestino = wkbOrigem
8       Else
            'Abre um arquivo existente
9           If VBA.varType(wkbOrigem) = VBA.vbString Then
10              If InStr(wkbOrigem, "\") > 0 Then
11                  Set wkbDestino = Me.openWorkbook(pFileName:=CStr(wkbOrigem))
12              End If
13          End If
14      End If
15  End If

    '--------------------------------------------------------------------
    'Define a planilha de destino. Uma nova ou planilha existente.
    '--------------------------------------------------------------------
16  If VBA.IsMissing(pDestino) Or VBA.IsNull(pDestino) Then
17      Set wks = wkbDestino.Worksheets.Add()
19      Set pDestino = wks
20  Else
21      If VBA.IsObject(pDestino) Then
            If VBA.TypeName(pDestino) = "Worksheet" Then
22              Set wks = pDestino
                Set rngInicial = wks.Range("A1")
            ElseIf VBA.TypeName(pDestino) = "Range" Then
                Set wks = pDestino.Worksheet
                Set rngInicial = pDestino
            End If
23      Else
            'Se nao for um objeto, pode uma string com o nome da planilha ou o endereço de uma celula
            If VBA.IsNull(pDestino) Then pDestino = VBA.vbNullString
            'Se a string vinher com um ! significa que esta sendo informado um endereço de celula, precisamos quebra o nome da celula e o nome da planilha
            If VBA.InStr(VBA.CStr(pDestino), "!") > 0 Then
                strPlanilha = VBA.Split(VBA.CStr(pDestino), "!")(0)
                strCelula = VBA.Split(VBA.CStr(pDestino), "!")(1)
            Else
                strPlanilha = VBA.CStr(pDestino)
                strCelula = "A1"
            End If
            
            On Error Resume Next
24          Set wks = wkbDestino.Worksheets(strPlanilha)
            'Erro 9, a planilha não existe
            If VBA.Err.Number = 9 Then
                Set wks = wkbDestino.Worksheets.Add()
                If pDestino <> "" Then wks.Name = VBA.Left(strPlanilha, 33) 'Limita em 33 caracteres (o maximo que o excel suporta)
            End If
            'Define a celula inicial
            Set rngInicial = wks.Range(strCelula)
26      End If
27  End If

    'Reinicia o tratamento de erros
    On Error GoTo ExportarParaPlanilha_Error
28  bFormatPadrao = colunasFormarComoValor = ""
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If rngInicial Is Nothing Then Set rngInicial = wks.Cells(1, 1)
    
29  Call TransferToExcelFromRecorset(pOrigem, rngInicial, pIncluirCabecalho, False, "CopyFromRecordSet", colunasFormarComoValor)
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    'Configura o resultado com uma lista
30  If Not VBA.IsMissing(pOpcoes) Then Set lst = ConfigurarLista(rngInicial.offset(-1, 1), pOpcoes)
    
    'Executa as ações de tratamento de formulas enviadas
    If Not VBA.IsMissing(pColarFormulas) And Not VBA.IsEmpty(wkbOrigem) Then
        If VBA.IsArray(pColarFormulas) Then
            Call ColarFormulas(wks, pColarFormulas)
        End If
    End If
    
    'Posiciona o cursor do mouse na celula A1
31  With wks.parent.Windows(1)
32      .ScrollRow = 1
33      .ScrollColumn = 1
34  End With

35  wkbDestino.REFRESHALL

    If pSalvarArquivoComo <> "" Then
        Call Me.WorkbookSaveAs(wks.parent, pSalvarArquivoComo, False, "")
        If Not pExibirArquivo Then
            Call Me.closeWorkbook(wks.parent, , True)
        End If
    End If
    
    If pExibirArquivo Then
        If VBA.MsgBox("Exibir arquivo gerado ?" & VBA.IIf(pSalvarArquivoComo = "", "", VBA.vbNewLine & VBA.vbNewLine & _
                      "Salvo em : " & pSalvarArquivoComo), VBA.vbQuestion + VBA.vbYesNo, "Gerar Excel") = VBA.vbYes Then
            Call Me.ShowHideExcel(True)
        End If
    End If

Fim:
36  On Error GoTo 0
37  Exit Function

ExportarParaPlanilha_Error:
38  If VBA.Err <> 0 Then
39      lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
        Debug.Print cstr_ProcedureName, VBA.Err.Number & "-" & VBA.Err.Description
        MsgBox VBA.vbNewLine & cstr_ProcedureName & " Em" & VBA.IIf(VBA.Erl() <> 0, " >(" & VBA.Erl() & ")", "") & " > " & strErrorMessagem, vbCritical
        Call VBA.Err.Raise(VBA.Err.Number, cstr_ProcedureName, VBA.vbNewLine & cstr_ProcedureName & " Em" & VBA.IIf(VBA.Erl() <> 0, " >(" & VBA.Erl() & ")", "") & " > " & strErrorMessagem)
        GoTo Fim
41  End If
    'Debug Mode
42  Resume
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cExcel.ExportToPDF()
' TIPO             : Function
' DATA/HORA        : 25/11/2015 17:46
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Exporta um documento (Planilha ou Workbbok) para PDF
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      25/11/2015 17:46
'---------------------------------------------------------------------------------------
Function ExportToPDF(objDocument As Object, Optional pFileOutPut As String, Optional bOpenAfterPublish As Boolean)
    '---------------------------------------------------------------------------------------
10  On Error GoTo ExportToPDF_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
20  Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "cExcel.ExportToPDF()"
    '---------------------------------------------------------------------------------------
    
    pFileOutPut = VBA.Replace(pFileOutPut, "/", "")
    
30  If pFileOutPut = "" Then
40      If VBA.TypeName(objDocument) = "Worksheet" Then
50          pFileOutPut = objDocument.parent.FullName
60      ElseIf VBA.TypeName(objDocument) = "Workbook" Then
70          pFileOutPut = objDocument.FullName
80      End If
90      pFileOutPut = VBA.Left(pFileOutPut, VBA.InStrRev(pFileOutPut, ".")) & ".PDF"
100 End If

    Call MkFullDirectory(PegarPasta(pFileOutPut))

110 Call objDocument.ExportAsFixedVBA.Format(Type:=0, _
                                         FileName:=pFileOutPut, _
                                         Quality:=0, _
                                         IncludeDocProperties:=True, _
                                         IgnorePrintAreas:=False, _
                                         OpenAfterPublish:=bOpenAfterPublish)

Fim:
120 On Error GoTo 0
130 Exit Function

ExportToPDF_Error:
140 If VBA.Err <> 0 Then
150     lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
        Debug.Print cstr_ProcedureName, "Linha : " & VBA.Erl() & " - " & strErrorMessagem
170 End If
    GoTo Fim:
    'Debug Mode
180 Resume

End Function


'Array(1, 9, 9, 9, 1, 9, 9, 9, 1, 1, 1, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9)
Function CriarConexaoTexto(strArquivo As String, _
                           Optional configColunas, _
                           Optional rngInicio, _
                           Optional bUpdate As Boolean = False, _
                           Optional strSalvarComo As String, _
                           Optional bExibirArquivo)
                           
1   On Error GoTo CriarConexaoTexto_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
2   Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "cExcel.CriarConexaoTexto()"
    
    Dim wks As Object
    Dim wkbDestino As Object
    Dim cnTXT As Object
    
    Me.Application.Calculation = -4135
    
    If VBA.IsMissing(rngInicio) Then
        Set wkbDestino = Me.openWorkbook()
        Set wks = wkbDestino.Worksheets(1)
        Set rngInicio = wks.Range("$A$1")
    Else
        If VBA.IsObject(rngInicio) And VBA.TypeName(rngInicio) = "Range" Then
            Set wks = rngInicio.Worksheet
            Set wkbDestino = wks.parent
        Else
            Set rngInicio = Me.Application.Range(CStr(rngInicio))
            Set wks = rngInicio.Worksheet
            Set wkbDestino = wks.parent
        End If
    End If

    Set cnTXT = wks.QueryTables.Add(Connection:="TEXT;" & strArquivo, Destination:=rngInicio)
    
    With cnTXT
        .Name = VBA.Dir(strArquivo)
        .AdjustColumnWidth = False
        .PreserveFormatting = False
        .FieldNames = True 'Não pega o cabeçalho
        .SaveData = False   'Quebra o link
        .TextFileStartRow = 1 'Inicia na 2 pq tem cabeçalho (deve ser ignorado)
        .TextFileSemicolonDelimiter = True
        .TextFilePlatform = 65001
        .TextFileParseType = 1 'xlDelimited
        .TextFileTextQualifier = 1 'xlTextQualifierDoubleQuote
        .TextFileColumnDataTypes = configColunas
        .TextFileTrailingMinusNumbers = True
        If bUpdate Then .Refresh BackgroundQuery:=False
    End With
    
    If strSalvarComo <> "" Then
        Call Me.WorkbookSaveAs(wks.parent, strSalvarComo, False, "")
    End If
    
    If bExibirArquivo Then
        If VBA.MsgBox("Exibir arquivo gerado ?" & VBA.IIf(strSalvarComo = "", "", VBA.vbNewLine & VBA.vbNewLine & _
                      "Salvo em : " & strSalvarComo), VBA.vbQuestion + VBA.vbYesNo, "Gerar Excel") = VBA.vbYes Then
            Call Me.ShowHideExcel(True)
        End If
    End If
    
Fim:
36  On Error GoTo 0
37  Exit Function

CriarConexaoTexto_Error:
38  If VBA.Err <> 0 Then
39      lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
41  End If
    GoTo Fim:
    'Debug Mode
42  Resume
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : mdlRotinasManutenção.TransferData()
' TIPO             : Function
' DATA/HORA        : 21/10/2014 11:29
' CONSULTOR        : (ADELSON)/ TECNUN - Adelson Rosendo Marques da Silva
' CONTATO          : adelson@tecnun.com.br
' DESCRIÇÃO        : Transfere os dados de um Recordset em para o Excel
'---------------------------------------------------------------------------------------
Private Function TransferToExcelFromRecorset(pQueryName As String, _
                                     rngDest As Object, _
                                     bIncludeHeader As Boolean, _
                                     bIncludeNumeracao As Integer, _
                                     metodoCopia As String, _
                                     ByRef colunaValores As String, _
                                     ParamArray Parameters()) As Variant
    Dim arrQuery As Variant
    Dim vRetExec As Variant
    Dim wks As Object, rs As Object
    Dim i As Integer
    Dim lngRowsToTransfer As Long
    Dim arrDados
    Dim bPegaFormat As Boolean

10  On Error GoTo TransferToExcelFromRecorset_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
20  Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "cExcel.TransferToExcelFromRecorset()"
    '----------------------------------------------------------------------------------------------------
    'bPegaFormat = colunaValores = ""
120 Set wks = rngDest.Worksheet

    'Se começar com SELECT, a instrução é SQL
    If UCase(pQueryName) Like "SELECT*" Then
        Set rs = CurrentDb.OpenRecordset(pQueryName)
    'Se for tabela, tambem pega o Recordset direto
    ElseIf PegarTipoDeObjeto(pQueryName) = acTable Then
        Set rs = CurrentDb.OpenRecordset(pQueryName)
    Else
        'Se for consulta (com ou sem parametros)
140     arrQuery = PegaQueryDef(pQueryName, Parameters)
150     Set rs = arrQuery(0).OpenRecordset()
    End If
    
    If Not rs.EOF Then: rs.MoveLast: rs.MoveFirst
    
160 If bIncludeHeader Then
        'Limpa a partir da celula inicial
        Set rngDest = PegarLista(rngDest, 0, 0, True)
        rngDest.ClearContents
        Set rngDest = rngDest.Cells(1, 1)
        
180     For i = 0 To rs.Fields.count - 1

190         rngDest.offset(0, i) = rs.Fields(i).Name 'wks.cells(rngDest.row, rngDest.Column + (i + 1)).value = rs.Fields(i).Name

            If bPegaFormat Then
            Select Case rs.Fields(i).Type
                Case dao.DataTypeEnum.dbDouble, dao.DataTypeEnum.dbDecimal, dao.DataTypeEnum.dbSingle
                colunaValores = colunaValores & ";" & VBA.Split(wks.Cells(1, i + 1).Address, "$")(1)
            End Select
            End If
200     Next
        Set rngDest = rngDest.offset(1)
210 End If
    
    If bPegaFormat Then colunaValores = Mid(colunaValores, 2)
    lngRowsToTransfer = rs.RecordCount
    rngDest.Worksheet.Activate
    Call rngDest.CopyFromRecordset(rs)
    rngDest.Worksheet.parent.Worksheets(1).Activate

380 If bIncludeNumeracao Then
400     With wks
410         .Range("a1").End(-4161).offset(0, 1).value = "Linha"
420         .Range("a1").End(-4161).offset(1).Resize(wks.Range("a1").End(-4121).row - 1).formular1c1 = "=ROW()-1"
430         .Range("a1").End(-4161).offset(1).Resize(wks.Range("a1").End(-4121).row - 1) = wks.Range("a1").End(-4161).offset(1).Resize(wks.Range("a1").End(-4121).row - 1).value
440     End With
450 End If

Fim:
480 On Error GoTo 0: Exit Function

TransferToExcelFromRecorset_Error:
490 If VBA.Err <> 0 Then
500     lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
        Debug.Print cstr_ProcedureName, VBA.Err.Number & "-" & VBA.Err.Description
        Call VBA.Err.Raise(VBA.Err.Number, cstr_ProcedureName, VBA.vbNewLine & cstr_ProcedureName & " Em" & VBA.IIf(VBA.Erl() <> 0, " > Linha (" & VBA.Erl() & ")", "") & " > " & VBA.Err.Description)
        GoTo Fim
520 End If
    GoTo Fim:
    'Debug Mode
530 Resume
End Function
'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cExcel.ConfigurarLista()
' TIPO             : Function
' DATA/HORA        : 19/03/2015 17:36
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Configura um intervalo com uma lista de dados
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      19/03/2015 17:36    Adelson         Criação/Atualização do procedimento
'---------------------------------------------------------------------------------------
Private Function ConfigurarLista(pToRange As Object, vConfigLista) As Object
    Dim lst As Object, rCol As Object
10  On Error GoTo ConfigurarLista_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
20  Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "cExcel.ConfigurarLista()"
    '----------------------------------------------------------------------------------------------------
    Dim sColunaAddress As String, sFormat As String
    Dim sColunasFormatarNumero As String
    Dim sLayoutList As String, sNameList As String
    'Layout padrão
    sLayoutList = "TableStyleLight1"
    
    'Configura o resultado com uma lista
    If Not VBA.IsMissing(vConfigLista) Then
        If VBA.IsArray(vConfigLista) Then
            sNameList = VBA.Split(CStr(vConfigLista(0)), "=")(1)
            sLayoutList = VBA.Split(CStr(vConfigLista(1)), "=")(1)
            sColunasFormatarNumero = VBA.Split(CStr(vConfigLista(2)), "=")(1)
            If sColunasFormatarNumero <> "" Then sColunasFormatarNumero = VBA.Split(sColunasFormatarNumero, ";")(0) & ":" & VBA.Split(sColunasFormatarNumero, ";")(UBound(VBA.Split(sColunasFormatarNumero, ";")))
        End If
        
        Set pToRange = pToRange.Worksheet.Range("A1").Resize(pToRange.Worksheet.Cells(pToRange.Worksheet.Rows.count, pToRange.Worksheet.Range("A1").Column).End(ExcelXlDirection.xlDirection_up).row, pToRange.Worksheet.Range("A1").End(ExcelXlDirection.xlDirection_Toright).Column)
       ' Set lst = ConfigurarLista(pToRange, vConfigLista)
    End If
    
    pToRange.Worksheet.AutoFilterMode = False
    
    'Delete a tabela existente
    On Error Resume Next
    Set lst = pToRange.Worksheet.ListObjects(sNameList)
    
    On Error GoTo ConfigurarLista_Error
    If lst Is Nothing Then
        Set lst = pToRange.Worksheet.ListObjects.Add(1, pToRange, , 1)
        lst.Name = sNameList
    End If
    
    'Ajusta o intervalo da lista de acordo com a endereço desejado
    If lst.Range.Address <> pToRange.Address Then
        Call lst.Resize(pToRange)
    End If
    
40  lst.TableStyle = sLayoutList
50  lst.showtotals = True

60  If sColunasFormatarNumero <> "" Then
        If InStr(sColunasFormatarNumero, "=") > 0 Then
            sColunaAddress = VBA.Split(sColunasFormatarNumero, "=")(0)
            sFormat = VBA.Split(sColunasFormatarNumero, "=")(1)
        Else
            sColunaAddress = sColunasFormatarNumero
            sFormat = "_-* #,##0.00_-"
        End If
70      Set rCol = pToRange.Worksheet.Range(sColunaAddress)
80      rCol.Range("A2").Resize(lst.ListRows.count + 1, rCol.Columns.count).NumberFormat = sFormat
        For Each rCol In pToRange.Worksheet.Range(sColunaAddress).Columns
            lst.ListColumns(rCol.Column).TotalsCalculation = 1
        Next
90  End If

100 pToRange.entirecolumn.AutoFit
110 Set ConfigurarLista = lst

Fim:
120 On Error GoTo 0
130 Exit Function

ConfigurarLista_Error:
140 If VBA.Err <> 0 Then
150     lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
        Debug.Print strErrorMessagem
160     'Call mTFW_Excecoes.TratarErro(strErrorMessagem, lngErrorNumber, cstr_ProcedureName, VBA.Erl)
170 End If
    GoTo Fim:
    'Debug Mode
180 Resume
End Function

Function ColarFormulas(wks As Object, vArrayFormulas)
    Dim iCol As Integer
    Dim rngFormula As Object
    For iCol = 0 To UBound(vArrayFormulas) Step 3
        Set rngFormula = PegarLista(wks.Range(VBA.CStr(vArrayFormulas(iCol))), 0, 0, False)
        rngFormula.formular1c1 = VBA.CStr(vArrayFormulas(iCol + 1))
        'O terceiro index informa se será colado valor ou nao
        If vArrayFormulas(iCol + 2) = True Then
            rngFormula.value = rngFormula.value
        End If
    Next iCol
End Function
'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : ConexaoDB.PegaQueryDef()
' TIPO             : Function
' DATA/HORA        : 03/06/2014 13:49
' CONSULTOR        : (ADELSON)/ TECNUN - Adelson Rosendo Marques da Silva
' CONTATO          : adelson@tecnun.com.br
' DESCRIÇÃO        : Recupera a instancia de um objeto QueryDef com argumentos parametros
'---------------------------------------------------------------------------------------
Private Function PegaQueryDef(ByVal strProc As String, ParamArray arrParametros() As Variant) As Variant
    Dim adoComando As Object ' dao.QueryDef
    Dim adoParametro As Object    'New ADODB.Parameter
    Dim btContador As Byte
    Dim pName  As String
    Dim pValue As String
    Dim codError As Integer, strMessage As String

    On Error GoTo PegaQueryDef_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
    Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "cExcel.PegaQueryDef()"
    
    arrParametros = Acertar_Array_Parametros(arrParametros)
    Set adoComando = CurrentDb.QueryDefs(strProc)
    With adoComando
        If UBound(arrParametros) >= 0 Then
            If UBound(arrParametros) = 0 Then
                If VBA.IsMissing(arrParametros(0)) Then GoTo Fim
            End If
            For btContador = 0 To UBound(arrParametros) Step 1
                pName = btContador
                pValue = Nz(arrParametros(btContador), "NULL")
                .Parameters(btContador).value = VBA.IIf(VBA.IsEmpty(arrParametros(btContador)), Null, arrParametros(btContador))
            Next btContador
        End If
    End With

    If VBA.Err.Number = 0 Then
        codError = 0
        strMessage = VBA.Now & " - Consulta/Comando '" & strProc & "' pronto pra ser executado !"
    End If
Fim:
    PegaQueryDef = Array(adoComando, Array(codError, strMessage))
    On Error GoTo 0
    Exit Function
PegaQueryDef_Error:
    Set adoComando = Nothing
    codError = VBA.Err.Number
    strMessage = VBA.Err.Description

    If VBA.Err.source = "DAO.Parameters" And codError = 3265 Then
        strMessage = "Parametro na posição '" & pName & "' com valor : '" & pValue & "' não encontrado na definição da consulta '" & strProc & "'" & VBA.vbNewLine & "  ### Há um erro na definição da consulta e passagem dos parametros..."
    End If

    If VBA.Err <> 0 Then
        Debug.Print cstr_ProcedureName, VBA.Err.Number & "-" & VBA.Err.Description
        Call VBA.Err.Raise(VBA.Err.Number, cstr_ProcedureName, VBA.vbNewLine & cstr_ProcedureName & " Em" & VBA.IIf(VBA.Erl() <> 0, " > Linha (" & VBA.Erl() & ")", "") & " > " & strErrorMessagem)
    End If
    GoTo Fim:
    Resume
End Function

'----------------------------------------------------------------------------------------------------
'Funções de Apoio
'----------------------------------------------------------------------------------------------------
Private Function ArrayBidimensional(ByVal arrParametros As Variant) As Boolean
On Error Resume Next 'Resume next necessario, pois o erro pode acontecer
Dim blResultado     As Boolean
Dim Var             As Variant
    Var = arrParametros(0)(0, 0)
    If VBA.Err.Number = 0 Then
        blResultado = True
    Else
        blResultado = False
    End If
    ArrayBidimensional = blResultado
End Function

'---------------------------------------------------------------------------------------
' Modulo    : ConexaoDB / Módulo de classe
' Rotina    : ContemMaisUbounds() / Function
' Autor     : Jefferson
' Data      : 07/11/2012 - 15:54
' Proposta  : Função para verificar se o arrParametros contém mais de uma dimensao (ubound)
'---------------------------------------------------------------------------------------
Private Function ContemMaisUbounds(ByVal arrParametros As Variant) As Boolean
On Error Resume Next 'Resume next necessario, pois o erro pode acontecer
Dim blResultado     As Boolean

    If UBound(arrParametros(0)) >= 0 Then
        If VBA.Err.Number = 0 Then blResultado = True
    Else
        blResultado = False
    End If
    ContemMaisUbounds = blResultado
End Function

'---------------------------------------------------------------------------------------
' Modulo    : AuxArray / Módulo
' Rotina    : AcertarArrayParamentros() / Function
' Autor     : Jefferson
' Data      : 07/11/2012 - 15:55
' Proposta  : Função para acertar o array caso ele tenha mais de uma dimensao
'---------------------------------------------------------------------------------------
Private Function Acertar_Array_Parametros(ByVal arrParametros As Variant) As Variant
On Error GoTo TratarErro
Dim arrAux          As Variant
Dim btContador      As Byte

    If ContemMaisUbounds(arrParametros) Then
        If ArrayBidimensional(arrParametros) Then
            ReDim arrAux(UBound(arrParametros(0), 2))
            For btContador = 0 To UBound(arrParametros(0), 2) Step 1
                arrAux(btContador) = arrParametros(0)(1, btContador)
            Next btContador
        Else
            ReDim arrAux(UBound(arrParametros(0)))
            For btContador = 0 To UBound(arrParametros(0)) Step 1
                arrAux(btContador) = arrParametros(0)(btContador)
            Next btContador
        End If
    Else
        arrAux = arrParametros
    End If
    Acertar_Array_Parametros = arrAux

Exit Function
TratarErro:
    If VBA.Err <> 0 Then
        Debug.Print VBA.Err.Description
        'Call mTFW_Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "ConexaoDB.PegaQueryDef()", VBA.Erl, , False)
    End If
    Resume
End Function

'Força o fechamento de todas as instancias do Excel
Sub FecharProcessoExcel(Optional bTodas As Boolean = False, Optional bForce As Boolean = True)
    Dim lngRet As Long
    If bTodas Then
        lngRet = VBA.Shell("taskkill.exe /im excel.exe " & VBA.IIf(bForce, " /f", ""))
        If lngRet <> 0 Then
            Debug.Print "Erro ao finalizar os processos do Excel em Andamento"
        End If
    Else
        Call m_objApplication.quit
        Set m_objApplication = Nothing
    End If
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
    On Error Resume Next
    Dim sTempFileName As String, fso As Object
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

Private Function FileIsLocked(sFilePath As String) As Boolean
    On Error GoTo IsOpen
    Dim intFile As Long
    If Not FileExists(sFilePath) Then Exit Function
    intFile = VBA.FreeFile
    Open sFilePath For Binary Access Read Lock Read As intFile: Close intFile
    FileIsLocked = False: Exit Function
IsOpen:
    FileIsLocked = True
End Function

'Determina, de acordo com o nome do objeto, qual o seu tipo : Form, Query, Table, Report
Private Function PegarTipoDeObjeto(objName As String) As AcObjectType
    Dim objType As String
1   objType = CLng(Nz(DLookup("Type", "MSysObjects", "Name='" & VBA.Replace(VBA.Replace(objName, "[", ""), "]", "") & "'")))
2   Select Case objType
    Case -32768
3       PegarTipoDeObjeto = AcObjectType.acForm
4   Case 5
5       PegarTipoDeObjeto = AcObjectType.acQuery
6   Case 1, 6
7       PegarTipoDeObjeto = AcObjectType.acTable
8   Case -32764
9       PegarTipoDeObjeto = AcObjectType.acReport
10  Case -32761
11      PegarTipoDeObjeto = AcObjectType.acReport
12  End Select
End Function

Function CriarNovaPlanilha(ParamArray pNome()) As Object
    On Error Resume Next
    Dim wks As Object
    Dim pN
    For Each pN In pNome
        Set wks = Me.ThisWorkbook.Worksheets.Add()
        wks.Name = VBA.CStr(pN)
    Next
    Set CriarNovaPlanilha = wks
End Function

Function ExcluirPlanilhas(ParamArray pNome()) As Object
    On Error Resume Next
    Dim wks As Object
    Dim pN
    Dim bDisplayAlerts As Boolean
    bDisplayAlerts = Me.Application.DisplayAlerts
    Me.Application.DisplayAlerts = False
    For Each pN In pNome
        Call Me.ThisWorkbook.Worksheets(VBA.CStr(pN)).Delete
    Next
    Set CriarNovaPlanilha = wks
    Me.Application.DisplayAlerts = bDisplayAlerts
End Function

Function CriarBackupWorkbook(strFileSource As String)
    Dim strTempBackup As String
    strTempBackup = PegarPasta(strFileSource) & "\BKP_" & VBA.Format(VBA.Now, "yyyymmdd_hhnnss") & "_" & VBA.Dir(strFileSource)
    Call VBA.FileCopy(strFileSource, strTempBackup)
    CriarBackupWorkbook = VBA.Dir(strTempBackup) <> ""
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cExcel.PegarPasta()
' TIPO             : Function
' DATA/HORA        : 01/09/2015 18:40
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Retorna o endereço da pasta mae de um arquivo ou pasta
'---------------------------------------------------------------------------------------
Private Function PegarPasta(pCaminho As String)
    Dim ofso As Object
    Set ofso = VBA.CreateObject("Scripting.FileSystemObject")
    PegarPasta = ofso.GetParentFolderName(pCaminho)
    Set ofso = Nothing
End Function
'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cExcel.PegarPasta()
' TIPO             : Function
' DATA/HORA        : 10/02/2017 10:23
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Retora a instancia de um range baseado nos parametros de iniciais
'---------------------------------------------------------------------------------------

Public Function PegarLista(rngInicio As Object, Optional OffsetRow As Integer = 0, Optional OffsetCol As Integer = 0, Optional bExpandirColunas As Boolean = True) As Object
    On Error Resume Next
    If rngInicio.offset(1) = "" Then
        Set PegarLista = rngInicio.Worksheet.Range(rngInicio, rngInicio.End(xlDirection_Toright).offset(OffsetRow, OffsetCol))
        Exit Function
    End If
    If bExpandirColunas Then
        Set PegarLista = rngInicio.Worksheet.Range(rngInicio.End(xlDirection_Down), rngInicio.End(xlDirection_Toright).offset(OffsetRow, OffsetCol))
    Else
        Set PegarLista = rngInicio.Worksheet.Range(rngInicio.offset(OffsetRow, OffsetCol), rngInicio.End(xlDirection_Down))
    End If
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cExcel.CopyRangeRef()
' TIPO             : Function
' DATA/HORA        : 10/02/2017 10:23
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : A partir da referencia de um range, retorna a instancia de um outro range baseado no tamanho do range de entrada
'---------------------------------------------------------------------------------------
Function CopyRangeRef(FromRange As Object, toRangeRef As Object) As Object
    Set CopyRangeRef = toRangeRef.Resize(FromRange.Rows.count, FromRange.Columns.count)
End Function
