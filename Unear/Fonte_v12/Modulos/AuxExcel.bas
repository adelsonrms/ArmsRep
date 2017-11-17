Attribute VB_Name = "AuxExcel"
Option Compare Database

'---------------------------------------------------------------------------------------
' Modulo....: AuxExcel / Módulo
' Rotina....: SetSheet() / Function
' Autor.....: Jefferson
' Contato...: jefferson.dantas@mondial.com.br
' Data......: 08/11/2012 - 12:09
' Empresa...: Mondial Tecnologia em Informática LTDA.
' Descrição.: Rotina para setar uma sheet de acordo com o seu nome, ela procura tanto
'             no name quanto no codename da sheet.
'---------------------------------------------------------------------------------------
Function SetSheet(ByRef wbk As Object, ByVal strNome As String) As Object
    On Error GoTo TratarErro
    Dim sht    As Object

    For Each sht In wbk.Worksheets
        If sht.Name = strNome Or sht.codename = strNome Then
            Set SetSheet = sht
            Exit Function
        End If
    Next sht
    Exit Function
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxExcel.SetSheet()", Erl)
End Function

'---------------------------------------------------------------------------------------
' Modulo....: AuxExcel / Módulo
' Rotina....: FechaWBK / Sub
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 15/08/2013
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina que fecha o Workbook fazendo algumas verificações e salvando se
'             necessário e também a mesma verificação para fechar a aplicação
'---------------------------------------------------------------------------------------
Public Sub FecharWBK(ByRef wbk As Object, Optional ByVal Salva As Boolean = False, _
                     Optional ByRef xlsapp As Object = Nothing, Optional ByVal FechaAplicacao As Boolean = False)
On Error Resume Next                                        'Em caso de erro a rotina deve continuar a execução
Dim wbkEach As Object

    Call wbk.Close(SaveChanges:=Salva)                      'não precisa testar o if is nothing, pois tem resume next
    Call Publicas.RemoverObjetosMemoria(wbk)

    If Not VBA.IsMissing(xlsapp) And FechaAplicacao Then
        For Each wbkEach In xlsapp.Workbooks
            Call FecharWBK(wbkEach, False)
        Next wbkEach
        xlsapp.quit                                         'não precisa testar o if is nothing, pois tem resume next
        Call Publicas.RemoverObjetosMemoria(xlsapp)
    End If
    
End Sub

'---------------------------------------------------------------------------------------
' Modulo....: AuxExcel / Module
' Rotina....: GetUsedRange() / Function
' Autor.....: Jefferson
' Contato...: jefferdantas@gmail.com
' Data......: 12/27/2012 (mdy)
' Empresa...: Planilhando
' Descrição.: Returns the real used range from a worksheet
'---------------------------------------------------------------------------------------
Public Function GetUsedRange(ByRef sht As Object) As Object
On Error GoTo TreatError
Dim rngAux      As Object
Dim ValueAux    As Long
Dim ReadRow     As Long
Dim ReadCol     As Long
Dim CountRows   As Long
Dim CountCols   As Long
Dim MaxRow      As Long
Dim MaxCol      As Long
    
    If Not sht Is Nothing Then
        Call RemoverFiltro(sht)
        With sht
            ReadRow = VBA.IIf(.usedrange.Rows.count + 10 >= .Rows.count, .usedrange.Rows.count, .usedrange.Rows.count + 10)
            ReadCol = VBA.IIf(.usedrange.Columns.count + 10 >= .Columns.count, .usedrange.Columns.count, .usedrange.Columns.count + 10)
            ValueAux = 0
    
            For CountRows = 1 To .usedrange.Rows.count Step 1
                ValueAux = .Cells(CountRows, ReadCol).End(xlToLeft).Column  '
                If ValueAux > MaxCol Then MaxCol = ValueAux
            Next CountRows
    
            ValueAux = 0
            For CountCols = 1 To .usedrange.Columns.count Step 1
                ValueAux = .Cells(ReadRow, CountCols).End(xlUp).row  '
                If ValueAux > MaxRow Then MaxRow = ValueAux
            Next CountCols
                
            If MaxRow = 1 And MaxCol = 1 Then
                Set rngAux = .Cells(1, 1)
            Else
                Set rngAux = .Range(.Cells(1, 1), .Cells(MaxRow, MaxCol))
            End If
            Set GetUsedRange = rngAux
        End With
    End If
On Error GoTo 0
    
Exit Function
TreatError:
    Resume Next
    
End Function

'---------------------------------------------------------------------------------------
' Modulo....: xlWorksheets / Módulo
' Rotina....: RemoverFiltro() / sub
' Autor.....: Fernando Couto Fernandes
' Contato...: Fernando.Fernandes@Outlook.com.br
' Data......: 12/20/2013
' Descrição.: This routine will remove filter in the given worksheets
'---------------------------------------------------------------------------------------
Public Sub RemoverFiltro(ParamArray arrWorksheets() As Variant)
On Error GoTo TreatError
Dim sht As Object
Dim cnt As Long
    
    For cnt = 0 To UBound(arrWorksheets) Step 1
        If Not arrWorksheets(cnt) Is Nothing Then
        
            If VBA.TypeName(arrWorksheets(cnt)) = "Worksheet" Then
            
                Set sht = arrWorksheets(cnt)
                
                With sht
                    If .AutoFilterMode Then .AutoFilterMode = False
                End With
                
                Set sht = Nothing
                
            End If
            
        End If
    Next cnt
    
Exit Sub
TreatError:
    Resume Next
End Sub

'---------------------------------------------------------------------------------------
' Modulo....: AuxExcel / Módulo
' Rotina....: CriarTitulos / Function
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 14/08/2013
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina para a importação dos DE_Paras conforme arquivo em excel.
'             Incluir posição ser para concatenar a posição da coluna com o nome da coluna
'---------------------------------------------------------------------------------------
Public Function CriarTitulos(ByVal arrDados As Variant, Optional ByVal linha As Long = 1, _
                             Optional ByVal IncluiPosicao As Boolean = False) As Object 'Scripting.Dictionary
On Error GoTo TrataErro
Dim dicAux          As Object
Dim Titulo          As String
Dim contador        As Integer
    
    Set dicAux = VBA.CreateObject("Scripting.Dictionary")

    If VBA.IsArray(arrDados) Then
        For contador = 1 To UBound(arrDados, 2) Step 1
            Titulo = RemoverQuebrasDeLinha(VBA.UCase(VBA.Trim(arrDados(linha, contador))))

            If IncluiPosicao Then Titulo = contador & "|" & Titulo
            If Not dicAux.Exists(Titulo) Then
                Call dicAux.Add(Titulo, contador)
            End If
        Next contador
    End If
    
    Set CriarTitulos = dicAux
    Call Publicas.RemoverObjetosMemoria(dicAux)
    
Exit Function
TrataErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxExcel.CriarTitulos()", Erl)
End Function

'---------------------------------------------------------------------------------------
' Modulo....: AuxExcel / Módulo
' Rotina....: AbrirWBK / Function
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 16/08/2013
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina para abrir um wbk, tanto por um File quanto por dialog, há ainda
'             a possibilidade de verificar se uma sheet exist como validação do wbk.
'---------------------------------------------------------------------------------------
Public Function AbrirWBK(ByRef xlsapp As Object, ByRef fsoFile As Object, _
                         ByRef sht As Object, ByVal strNomePlan As String, _
                         Optional password As String = VBA.vbNullString, _
                         Optional ByVal WriteResPassword As String = VBA.vbNullString, _
                         Optional ByVal ReadOnly As Boolean = False) As Object
On Error GoTo TrataErro
Dim wbk         As Object
Dim fso         As Object
Set fso = VBA.CreateObject("Scripting.FileSystemObject")

TENTARNOVAMENTE:
    If Not fsoFile Is Nothing Then GoTo AbrirArquivo
    
    With Access.Application.FileDialog(Office.MsoFileDialogType.msoFileDialogFilePicker)
        .Title = "Escolha um arquivo para exportação dos dados"
        .InitialView = msoFileDialogViewDetails
        .Filters.Clear
        .Filters.Add "Arquivos em Excel ", "*.xls*,*.csv"
        .show
        
        If Not .SelectedItems.count = 0 Then
            Set fso = VBA.CreateObject("Scripting.FileSystemObject")
            Set fsoFile = fso.GetFile(.SelectedItems(1))
AbrirArquivo:
            Set wbk = PegarWBK(xlsapp:=xlsapp, caminho:=fsoFile.Path, password:=password, WriteResPassword:=WriteResPassword, ReadOnly:=ReadOnly)
            
            If Not wbk Is Nothing Then
                xlsapp.Visible = False
                If Not strNomePlan = VBA.vbNullString Then
                    Set sht = SetSheet(wbk, strNomePlan)
                    If sht Is Nothing Then
                        If fso Is Nothing Then
                            Call AuxExcel.FecharWBK(wbk, False)
                        ElseIf VBA.MsgBox("O arquivo escolhido não contém a planilha """ & strNomePlan & """" & VBA.vbNewLine & _
                                      "Deseja escolher outro arquivo?" & VBA.vbNewLine & VBA.vbNewLine & _
                                      "Nota: Se escolher ""Não"" a exportação será cancelada", VBA.vbYesNo + VBA.vbQuestion, _
                                      "Planilha não encontrada") = VBA.vbYes Then
                            Call AuxExcel.FecharWBK(wbk, False)
                            GoTo TENTARNOVAMENTE
                        Else
                            Call AuxExcel.FecharWBK(wbk, False, xlsapp, True)
                        End If
                    End If
                End If
            End If
        Else
        
            If VBA.MsgBox("Nenhum arquivo foi escolhido." & VBA.vbNewLine & "Deseja cancelar a exportação?", VBA.vbYesNo + VBA.vbQuestion, "Escolher Arquivo") = VBA.vbNo Then
                GoTo TENTARNOVAMENTE
            End If
            
        End If
    End With
Fim:
    Set AbrirWBK = wbk
    
Exit Function
TrataErro:

    If Not fsoFile Is Nothing Then
        Call Excecoes.TratarErro(VBA.Err.Description & " - NomeArquivo: " & fsoFile.Name, VBA.Err.Number, "AuxExcel.AbrirWBK()", Erl)
    Else
        Call Excecoes.TratarErro(VBA.Err.Description & " - FALHA NO FSO.", VBA.Err.Number, "AuxExcel.AbrirWBK()", Erl)
    End If
    Call AuxExcel.FecharWBK(wbk, False, xlsapp, True)
    GoTo Fim
End Function

'---------------------------------------------------------------------------------------
' Modulo....: AuxExcel / Módulo
' Rotina....: PegarWBK / Function
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 22/08/2013
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina que tenta abrir o arquivo excel e se não conseguir ele retorna vazio
'---------------------------------------------------------------------------------------
Private Function PegarWBK(ByRef xlsapp As Object, ByVal caminho As String, _
                          Optional password As String = VBA.vbNullString, _
                          Optional ByVal WriteResPassword As String = VBA.vbNullString, _
                          Optional ReadOnly As Boolean = False) As Object
                          
On Error Resume Next 'Resume next necessário
Dim wbk         As Object

    If Not password = VBA.vbNullString And Not WriteResPassword = VBA.vbNullString Then
        Set wbk = xlsapp.Workbooks.Open(FileName:=caminho, UpdateLinks:=False, ReadOnly:=ReadOnly, password:=password, WriteResPassword:=WriteResPassword)
        
    ElseIf Not password = VBA.vbNullString Then
        Set wbk = xlsapp.Workbooks.Open(FileName:=caminho, UpdateLinks:=False, ReadOnly:=ReadOnly, password:=password)
        
    ElseIf Not WriteResPassword = VBA.vbNullString Then
        Set wbk = xlsapp.Workbooks.Open(FileName:=caminho, UpdateLinks:=False, ReadOnly:=ReadOnly, WriteResPassword:=WriteResPassword)
        
    Else
        Set wbk = xlsapp.Workbooks.Open(FileName:=caminho, UpdateLinks:=False, ReadOnly:=ReadOnly)
        
    End If
    
    If VBA.Err.Number = 0 Then Set PegarWBK = wbk Else Set PegarWBK = Nothing
    
End Function

'---------------------------------------------------------------------------------------
' Modulo....: AuxExcel / Módulo
' Rotina....: PreencherTitulos / Sub
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 26/08/2013
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina para criar a linha de titulo de uma tabela de acordo com um recordset
' Enums utilizados:
'   - Excel.XlBorderWeight.xlMedium = -4138
'   - Excel.XlBorderWeight.xlHairline = 1
'---------------------------------------------------------------------------------------
Public Sub PreencherTitulos(ByRef sht As Object, ByRef rs As Object, _
                            ByVal LinhaTitulo As Long, ByVal ColunaTitulo As Integer, _
                            Optional ByVal Formatar As Boolean = False)
On Error GoTo TratarErro
Dim contador                As Integer
Dim rng                     As Object 'Excel.Range

    ReDim arrTitulos(0 To 0) As String

    With sht
        For contador = 0 To rs.Fields.count - 1 Step 1
            ReDim Preserve arrTitulos(0 To contador) As String
            arrTitulos(contador) = rs.Fields(contador).Name
        Next contador
        .Range(.Cells(LinhaTitulo, ColunaTitulo), .Cells(LinhaTitulo, ColunaTitulo + contador - 1)).value = arrTitulos

        If Formatar Then
            With .Rows(LinhaTitulo)
                .Font.Bold = True
                .VerticalAlignment = xlEnums.xlCenter '- 4108   'xlCenter
                .HorizontalAlignment = xlEnums.xlCenter '-4108    'xlCenter
            End With
            Set rng = .Range(.Cells(LinhaTitulo, 1), .Cells(LinhaTitulo, contador))
            Call AuxExcel.InserirBordas_Ext_Int(rng, xlEnums.xlMedium, 1) '-4138, 1)
        End If
    End With

On Error GoTo 0
Exit Sub
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxExcel.PreencherTitulos", Erl)
End Sub



'---------------------------------------------------------------------------------------
' Modulo....: AuxExcel / Módulo
' Rotina....: InserirBordas_Ext_Int / Sub
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 26/08/2013
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina que cria as bordas externas e internas com seus respectivos tamanhos
'---------------------------------------------------------------------------------------
Public Sub InserirBordas_Ext_Int(ByRef rngCells As Object, ByVal TamanhoBordaExt As Integer, _
                                 ByVal TamanhoBordaInt As Integer)
On Error GoTo TratarErro

    Call InserirBordas(rngCells, TamanhoBordaExt, xlEdgeTop, xlEdgeBottom, xlEdgeLeft, xlEdgeRight)
    Call InserirBordas(rngCells, TamanhoBordaInt, xlInsideVertical, xlInsideHorizontal)
    
On Error GoTo 0
Exit Sub
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxExcel.InserirBordas_Ext_Int", Erl)
End Sub


'---------------------------------------------------------------------------------------
' Modulo....: AuxExcel / Módulo
' Rotina....: InserirBordas / Sub
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 26/08/2013
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina que coloca bordas nas diversas posicoes de um range
' Enums Utilizados:
'   - Excel.XlBorderWeight
'---------------------------------------------------------------------------------------
Public Sub InserirBordas(ByRef rngCells As Object, ByVal TamanhoBorda As Integer, _
                         ParamArray BordaIndex() As Variant)
On Error GoTo TratarErro
Dim contador As Integer

    For contador = 0 To UBound(BordaIndex) Step 1
        Call InserirBordaIndividual(rngCells, BordaIndex(contador), TamanhoBorda)
    Next contador

On Error GoTo 0
Exit Sub
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxExcel.InserirBordas", Erl)
End Sub

'---------------------------------------------------------------------------------------
' Modulo....: AuxExcel / Módulo
' Rotina....: InserirBordaIndividual / Sub
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 26/08/2013
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina para criar bordas em um range com a opção de cores
'---------------------------------------------------------------------------------------
Public Sub InserirBordasColoridas(ByRef rngCells As Object, ByVal TamanhoBorda As Integer, _
                                  ByVal ThemeColor As Integer, _
                                  ByVal TintAndShade As Double, _
                                  ParamArray BordaIndex() As Variant)
On Error GoTo TratarErro
Dim contador        As Integer

    For contador = 0 To UBound(BordaIndex) Step 1
        Call InserirBordaIndividual(rngCells, BordaIndex(contador), TamanhoBorda, _
                                    ThemeColor, TintAndShade)
    Next contador

On Error GoTo 0
Exit Sub
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxExcel.InserirBordas", Erl)
End Sub

'---------------------------------------------------------------------------------------
' Modulo....: AuxExcel / Módulo
' Rotina....: InserirBordaIndividual / Sub
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 26/08/2013
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina para criar bordas em um range
' Enums utilizados:
'   - Excel.XlBorderWeight
'   - Excel.XlBordersIndex
'   - Excel.XlLineStyle.xlContinuous = 1
'   - Excel.Constants.xlAutomatic = -4105
'---------------------------------------------------------------------------------------
Private Sub InserirBordaIndividual(ByRef rngCells As Object, ByVal Borda As Integer, _
                                   ByVal Tamanho As Integer, _
                                   Optional ByVal ThemeColor As Integer = -1, _
                                   Optional ByVal TintAndShade As Double = -1)
On Error GoTo TratarErro

    With rngCells.Borders(Borda)
        .LineStyle = 1
        If ThemeColor = -1 And TintAndShade = -1 Then
        .ColorIndex = xlEnums.xlAutomatic '-4105
        Else
            .ThemeColor = ThemeColor
            .TintAndShade = TintAndShade
        End If
        .Weight = Tamanho
    End With

On Error GoTo 0
Exit Sub
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxExcel.InserirBordaIndividual", Erl)
End Sub

'---------------------------------------------------------------------------------------
' Modulo....: AuxExcel / Módulo
' Rotina....: CriarTitulos / Sub
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 27/08/2013
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina para criar um titulo mesclado conforme a necessidade de cada relatório
' Enums Utilizados:
'   - Excel.XlBorderWeight
'   - Excel.XlBordersIndex.xlEdgeTop = 8
'   - Excel.XlBordersIndex.xlEdgeBottom = 9
'   - Excel.XlBordersIndex.xlEdgeLeft = 7
'   - Excel.XlBordersIndex.xlEdgeRight = 10
'   - Excel.XlBorderWeight.xlMedium = -4138
'---------------------------------------------------------------------------------------
Public Sub CriarTituloNoRelatorio(ByRef sht As Object, ByVal TextoTitulo As String, _
                                  ByVal LinhaInicio As Long, ByVal ColunaInicio As Integer, _
                                  Optional ByVal altura As Long = 1, Optional ByVal Largura As Integer = 1)
On Error GoTo TratarErro
Dim rng             As Object

    With sht
        Set rng = .Range(.Cells(LinhaInicio, ColunaInicio), .Cells(LinhaInicio + altura - 1, ColunaInicio + Largura - 1))
        With rng
            .Merge
            .value = TextoTitulo
            With .Font
                .Bold = True
                .size = 18
            End With
            .HorizontalAlignment = xlEnums.xlCenter '-4108    'Excel.Constants.xlCenter
            .VerticalAlignment = xlEnums.xlCenter '-4108      'Excel.Constants.xlCenter
            Call AuxExcel.InserirBordas(rng, xlEnums.xlMedium, xlEnums.xlEdgeTop, xlEnums.xlEdgeBottom, xlEnums.xlEdgeLeft, xlEnums.xlEdgeRight) '-4138, 8, 9, 7, 10)
        End With
    End With
    Call Publicas.RemoverObjetosMemoria(rng)
On Error GoTo 0
Exit Sub
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxExcel.CriarTitulos", Erl)
End Sub


'---------------------------------------------------------------------------------------
' Modulo....: AuxExcel / Módulo
' Rotina....: SalvarWBK / Sub
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 26/08/2013
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina para salvar o arquivo em excel
'---------------------------------------------------------------------------------------
Public Function SalvarComoWBK(ByRef wbk As Object, ByVal caminho As String, _
                              ByVal NomeArquivo As String) As Boolean
On Error GoTo TratarErro
Dim fso             As Object
Dim Resultado       As Boolean
Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    If Not wbk Is Nothing Then
        With fso
            If .FileExists(.BuildPath(caminho, NomeArquivo)) Then
                Call .DeleteFile(.BuildPath(caminho, NomeArquivo), True)
            End If
            Call wbk.SaveAs(.BuildPath(caminho, NomeArquivo))
            Resultado = True
        End With
    End If
Fim:
    Call Publicas.RemoverObjetosMemoria(fso)
    SalvarComoWBK = Resultado
On Error GoTo 0
Exit Function
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxExcel.SalvarWBK", Erl, MostrarStop:=False)
    Resultado = False:  GoTo Fim
End Function




'---------------------------------------------------------------------------------------
' Modulo....: auxExcel / Module
' Rotina....: Corresp() / Function
' Autor.....: Fernando Fernandes
' Contato...: fernando@tecnun.com.br
' Data......: 28/09/2015 (dmy)
' Empresa...: Tecnun
' Descrição.: Essa rotina encontra um item num intervalo, e devolve a posição
'---------------------------------------------------------------------------------------
Public Function Corresp(ByRef xlsapp As Object, _
                        ByVal OQue As Variant, _
                        ByRef Aonde As Object) As Long
10    On Error GoTo TratarErro
          
20        If Not xlsapp Is Nothing Then
30           Corresp = xlsapp.WorksheetFunction.match(OQue, Aonde, 0)
40        End If

Fim:
50    On Error GoTo 0
60    Exit Function
TratarErro:
70        With Aonde
80            Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxExcel.Corresp(" & OQue & "," & .parent.Name & "!" & .Address & ")", Erl)
90        End With
100       Corresp = 0
110       GoTo Fim
End Function



'---------------------------------------------------------------------------------------
' Modulo....: auxExcel / Module
' Rotina....: ArrumarNomeParaPlanilha() / Function
' Autor.....: Fernando Fernandes
' Contato...: fernando@tecnun.com.br
' Data......: 02/04/2015 (dmy)
' Empresa...: Tecnun
' Descrição.: Essa rotina corrige um texto para poder ser usado como nome de uma planilha
'---------------------------------------------------------------------------------------
Public Function ArrumarNomeParaPlanilha(ByVal nome As String) As String
On Error GoTo TratarErro

    nome = VBA.Replace(nome, "\", ".")
    nome = VBA.Replace(nome, "/", ".")
    nome = VBA.Replace(nome, ":", ".")
    nome = VBA.Replace(nome, "*", ".")
    nome = VBA.Replace(nome, "?", ".")
    ArrumarNomeParaPlanilha = VBA.Left(nome, 31)
    
On Error GoTo 0
Exit Function
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "Exportar.ArrumarNomeParaPlanilha()", Erl)
End Function


'---------------------------------------------------------------------------------------
' Modulo....: auxExcel / Module
' Rotina....: Arrumar() / Function
' Autor.....: Fernando Fernandes
' Contato...: fernando@tecnun.com.br
' Data......: 03/15/2015 (mdy)
' Empresa...: Tecnun
' Descrição.: Usa o trim da planilha que é melhor que o Trim do VBA.
'             Recebe um texto e retorna um texto "arrumado".
'---------------------------------------------------------------------------------------
Public Function Arrumar(ByRef xlsapp As Object, ByRef sValor As Variant) As String
On Error GoTo TratarErro
Dim wsf As Object
Dim num As Long
Dim tam As Long

    If Not xlsapp Is Nothing Then
        Set wsf = xlsapp.WorksheetFunction
        Arrumar = wsf.Trim(sValor)
        Set wsf = Nothing
    Else
        Do While tam <> VBA.Len(sValor)
            tam = VBA.Len(sValor)
            sValor = VBA.Replace(sValor, "  ", " ", 1, -1, VBA.vbTextCompare)
        Loop
        Arrumar = VBA.Trim(sValor)
        
    End If
On Error GoTo 0
Exit Function
TratarErro:
    Arrumar = sValor
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxExcel.Arrumar", Erl)
    Resume Next
End Function

'
'---------------------------------------------------------------------------------------
' Modulo....: auxExcel / Module
' Rotina....: PegarComentario() / Function
' Autor.....: Fernando Fernandes
' Contato...: fernando@tecnun.com.br
' Data......: 28/09/2015 (dmy)
' Empresa...: Tecnun
' Descrição.: Essa rotina devolve o comentário de uma célula, se houver
'---------------------------------------------------------------------------------------
Public Function PegarComentario(ByRef Celula As Object) As String
On Error Resume Next 'Necessário pois se não houver comentário, não precisa tratar o erro

    If Not Celula Is Nothing Then
        Set Celula = Celula.Cells(1)
        PegarComentario = Celula.comment.Text
    End If

On Error GoTo 0
End Function

'Ferramenta Desp Adm - BRAM
Public Function pegarCaminhoModelo(ByVal NomeArquivo As String, ByVal NomeRelatorio As String) As String
Dim fso             As Object
Dim caminho         As String
Set fso = VBA.CreateObject("Scripting.FileSystemObject")

    caminho = PegaEndereco_Templates() & "\" & NomeArquivo
    If Not fso.FileExists(caminho) Then
        Call AuxMensagens.MessageBoxMaster("F034", NomeArquivo, caminho)
        pegarCaminhoModelo = VBA.vbNullString
        Exit Function
    End If
    pegarCaminhoModelo = caminho
End Function

Public Function abrirModeloRelatorio(ByVal wbk As Object, ByRef xlsapp As Object, ByVal CaminhoModelo As String) As Object
Dim fso             As Object
Dim contador        As Integer
Set fso = VBA.CreateObject("Scripting.FileSystemObject")
On Error GoTo TrataErro
    If Not fso.FileExists(CaminhoModelo) Or (fso.GetExtensionName(CaminhoModelo) <> "xlsx" And fso.GetExtensionName(CaminhoModelo) <> "xls") Then
        Set abrirModeloRelatorio = Nothing
        Set xlsapp = Nothing
        Exit Function
    End If
    Set xlsapp = VBA.CreateObject("EXCEL.APPLICATION")
    Set abrirModeloRelatorio = xlsapp.Workbooks.Open(CaminhoModelo)
Exit Function
TrataErro:
    Stop
    Resume
    
End Function

