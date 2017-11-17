Attribute VB_Name = "AuxReport"
Option Compare Database

Function OnChange_FiltrarConsultaRelatorios()
    Dim af As Object
10  On Error GoTo OnChange_FiltrarConsulta_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
20  Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Function BLPublica.OnChange_FiltrarConsulta()"
    '----------------------------------------------------------------------------------------------------
    Dim sFiltro As String

30  On Error Resume Next
40  Set af = Screen.ActiveForm

50  If VBA.Err = 2475 Then
60      Call VBA.MsgBox("Não há formulário ativo !", VBA.vbExclamation)
70      Exit Function
80  End If

90  With af
100     af.sfLista.Report.Filter = "*"
110     sFiltro = FiltraFormConsultaReport(.sfLista.Report, .GuiaFiltros.Pages(.GuiaFiltros.value), "Text")
120     af.sfLista.Report.Filter = sFiltro
130     af.sfLista.Report.FilterOn = True
140     af.sfLista.Requery
150 End With

Fim:
160 On Error GoTo 0
170 Exit Function

OnChange_FiltrarConsulta_Error:
180 If VBA.Err <> 0 Then
190     lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
200     Call Excecoes.TratarErro(strErrorMessagem, lngErrorNumber, cstr_ProcedureName, VBA.Erl)
210 End If
    GoTo Fim:
    'Debug Mode
220 Resume
End Function

Function OnClickToBrowseItemsRelatorios(Optional actForm As Object, Optional bMultSelecao As Boolean = True)
    Dim af As Object
10  On Error GoTo OnClickToBrowseItemsRelatorios_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
20  Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Function BLPublica.OnClickToBrowseItemsRelatorios()"
    '----------------------------------------------------------------------------------------------------

30  On Error Resume Next
    If actForm Is Nothing Then
40      Set af = Screen.ActiveForm.Form
    Else
        Set af = actForm
    End If
    
50  If VBA.Err = 2475 Then
60      MessageBoxMaster "Não há formulário ativo !", VBA.vbExclamation
70      Exit Function
80  End If

90  With af
100     Call FiltraMultiSelecao(.Controls(getTheValue(.ActiveControl.tag, "ControlComboList")), , bMultSelecao)
110 End With

    Call OnChange_FiltrarConsultaRelatorios

Fim:
120 On Error GoTo 0
130 Exit Function

OnClickToBrowseItemsRelatorios_Error:
140 If VBA.Err <> 0 Then
150     lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
160     Call Excecoes.TratarErro(strErrorMessagem, lngErrorNumber, cstr_ProcedureName, VBA.Erl)
170 End If
    GoTo Fim:
    'Debug Mode
180 Resume
End Function

Private Function FiltraFormConsultaReport(frmConsulta As Report, ParentControls As Object, Optional propValue As String = "Value") As String
    Dim arrValores
10  On Error GoTo FiltraFormConsulta_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
20  Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Function BLPublica.FiltraFormConsulta()"
    '----------------------------------------------------------------------------------------------------

30  arrValores = pegarValores(ParentControls, propValue)
40  FiltraFormConsultaReport = ConstruirFiltro(arrValores, ParentControls, , frmConsulta)
Fim:
50  On Error GoTo 0
60  Exit Function

FiltraFormConsulta_Error:
70  If VBA.Err <> 0 Then
80      lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
90      Call Excecoes.TratarErro(strErrorMessagem, lngErrorNumber, cstr_ProcedureName, VBA.Erl)
100 End If
    GoTo Fim:
    'Debug Mode
110 Resume
End Function

Private Function pegarValores(ParentControls As Object, Optional propValue As String = "Value")
    Dim controle As Control
    Dim arrDados() As Variant
    Dim contador As Integer
    Dim i As Integer

10  ReDim arrDados(0 To 1, 0 To ParentControls.Controls.count)

20  For Each controle In ParentControls.Controls
30      If controle.controlType = acComboBox Or _
           controle.controlType = acTextBox Or _
           controle.controlType = acCheckBox Then
40          If Not VBA.IsArray(controle.value) Then
50              arrDados(0, contador) = controle.Name
60              On Error Resume Next
70              If Screen.ActiveControl.Name = controle.Name Then
80                  arrDados(1, contador) = Nz(controle.value)
90              Else
100                 If Not controle.controlType = acCheckBox Then
110                     If controle.InputMask <> "" Then
120                         If controle.InputMask = "!\(99"") ""!9900\-0000;;_" Then
130                             arrDados(1, contador) = VBA.Format(Nz(controle.value), "(0) 0000-0000")
140                         ElseIf controle.InputMask = "!\(99"") ""!99000\-0000;;_" Then
150                             arrDados(1, contador) = VBA.Format(Nz(controle.value), "(0) 0 0000-0000")
160                         Else
170                             arrDados(1, contador) = VBA.Format(Nz(controle.value), controle.InputMask)
180                         End If
190                     Else
200                         arrDados(1, contador) = Nz(controle.value)
210                     End If
220                 Else
230                     arrDados(1, contador) = Nz(controle.value)
240                 End If
250             End If
                'If arrDados(1, contador) <> "" Then
260             contador = contador + 1
270         End If
280     End If
290 Next controle

300 If contador <> 0 Then
310     ReDim Preserve arrDados(0 To 1, 0 To contador - 1)
320     pegarValores = arrDados
330 Else
340     ReDim arrDados(0 To 0, 0 To 0)
350 End If

360 pegarValores = arrDados
End Function
'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : BLPublica.ConstruirFiltro()
' TIPO             : Function
' DATA/HORA        : 31/05/2015 12:25
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Controi um instrução de criterios baseado em uma matriz de valores Campo=Valor
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      31/05/2015 12:25    Adelson         Criação/Atualização do procedimento
'---------------------------------------------------------------------------------------
Private Function ConstruirFiltro(arrValues As Variant, Optional objParent As Object, Optional Curinga As String = "*", Optional rsVerificaCampo As Object)
    Dim strFiltro As String
    Dim iCtr As Integer
    Dim strValue As String
    Dim vl, strCampo As String
    Dim strMuliSelecao As String
    Dim tipoDado As dao.DataTypeEnum

10  On Error GoTo ConstruirFiltro_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
20  Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Function BLPublica.ConstruirFiltro()"
    '----------------------------------------------------------------------------------------------------

30  If Not VBA.IsArray(arrValues) Then Exit Function
40  strFiltro = "1=1"

50  For iCtr = 0 To UBound(arrValues, 2)
        'Campo e Valor
60      If Not objParent Is Nothing Then
70          strCampo = VBA.Replace(arrValues(0, iCtr), objParent.Name & "_", "")
80      Else
90          strCampo = arrValues(0, iCtr)
100     End If

110     strValue = arrValues(1, iCtr)
        'Se for indicado mais de uma opção de filtro
120     If InStr(strValue, ";") > 0 Then
130         strMuliSelecao = ""
140         For Each vl In VBA.Split(strValue, ";")
150             If CStr(vl) <> "" Then
160                 strMuliSelecao = strMuliSelecao & " OR (" & strCampo & " LIKE '" & Curinga & "" & CStr(vl) & "" & Curinga & "')"
170             End If
180         Next vl
190         strMuliSelecao = Trim(strMuliSelecao)
200         strMuliSelecao = Mid(strMuliSelecao, 3)
210         If FormTemCampo(strCampo, rsVerificaCampo) Then
220             strFiltro = strFiltro & " AND (" & strMuliSelecao & ")"
230         End If
240     Else
250         If VBA.IsNumeric(strValue) Then
260             tipoDado = dbInteger
270         Else
280             tipoDado = dbText
290         End If
300         If FormTemCampo(strCampo, rsVerificaCampo) Then
310             If strValue <> "" Then
320                 strFiltro = strFiltro & " AND " & BuildCriteria(strCampo, tipoDado, "" & VBA.IIf(tipoDado = dbText, "'", "") & "" & Curinga & "" & strValue & "" & Curinga & "" & VBA.IIf(tipoDado = dbText, "'", "") & "")
330             End If
340         End If
350     End If
360 Next iCtr

370 If strFiltro = "1=1" Then strFiltro = ""
380 strFiltro = Trim(strFiltro)
390 If strFiltro <> "" Then
400     strFiltro = "(" & VBA.Replace(strFiltro, "1=1 AND ", "") & ")"
410 End If
420 ConstruirFiltro = strFiltro

Fim:
430 On Error GoTo 0
440 Exit Function

ConstruirFiltro_Error:
450 If VBA.Err <> 0 Then
460     lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
470     Call Excecoes.TratarErro(strErrorMessagem, lngErrorNumber, cstr_ProcedureName, VBA.Erl)
480 End If
    GoTo Fim:
    'Debug Mode
490 Resume
End Function

Private Function FormTemCampo(campo As String, sf As Object) As Boolean
10  On Error Resume Next
    Dim C As Object
20  Set C = sf.Controls(campo)
30  FormTemCampo = VBA.Err = 0
End Function

Function FiltraMultiSelecao(objCtr As Object, Optional bsalvaValor As Boolean = True, Optional bMultSelecao As Boolean = True)
    Dim vSelected As Variant, items_selecionados As String
    Dim strSQL As String
    Dim iColunas As Variant
    strSQL = getTheValue(objCtr.tag, "SQLBrowseItems")
    
    strSQL = strSQL & " IN '" & CurrentDb.Name & "'"
    
    iColunas = getTheValue(objCtr.tag, "QtdCols")
    vSelected = SelecionarItems(strSQL, bMultSelecao, ArraySelecionados, CurrentDb)
    items_selecionados = ""
    If VBA.IsEmpty(vSelected) Then
        'Call VBA.MsgBox("Nenhum item selecionado !",VBA.vbExclamation)
    Else
        For selecao = 0 To UBound(vSelected, 2)
            Select Case iColunas
            Case 1: items_selecionados = items_selecionados & ";" & vSelected(1, selecao)
            Case 2: items_selecionados = items_selecionados & ";" & vSelected(2, selecao)
            Case Else: items_selecionados = items_selecionados & ";" & vSelected(1, selecao) & " - " & vSelected(2, selecao)
            End Select
        Next
        If items_selecionados <> "" Then items_selecionados = Mid(items_selecionados, 2)
        
        If bsalvaValor Then
            objCtr.value = items_selecionados
            Call Access.Run(Mid(objCtr.OnChange, 6))
        End If
    End If
    FiltraMultiSelecao = items_selecionados
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : AuxReport.ExportarPara()
' TIPO             : Function
' DATA/HORA        : 19/07/2016 17:52
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Exporta o conteudo de uma consulta/tabela para um arquivo de saida (HTML, EXCEL ou PDF)
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      19/07/2016 17:52
'---------------------------------------------------------------------------------------
Function ExportarPara(strTabelaConsulta As String, Optional strArquivoSaida As String = "", Optional bAbrir As Boolean = False) As Variant
    Dim tp As AcObjectType
    Dim strExtenciaSaida As String

    '---------------------------------------------------------------------------------------
    On Error GoTo ExportarPara_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
    Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Function AuxReport.ExportarPara()"
    '---------------------------------------------------------------------------------------
    
    Access.Application.Echo False
    DoCmd.Hourglass True
    
3   If strArquivoSaida = "" Then strArquivoSaida = Environ("Temp") & "\" & VBA.Format(VBA.Now, "yyyymmdd_hhnnss") & ".html"

1   strExtenciaSaida = GetFileExtention(strArquivoSaida)

2   tp = PegarTipoDeObjeto(strTabelaConsulta)


4   Select Case tp
    Case AcObjectType.acTable
5       Call DoCmd.OutputTo(acOutputTable, strTabelaConsulta, PegarTipoSaida(strExtenciaSaida), strArquivoSaida, False, "", 65001, acExportQualityScreen)
6   Case AcObjectType.acQuery
7       Call DoCmd.OutputTo(acOutputQuery, strTabelaConsulta, PegarTipoSaida(strExtenciaSaida), strArquivoSaida, False, "", 65001, acExportQualityScreen)
8   End Select

9   ExportarPara = Array(VBA.Dir(strArquivoSaida) <> "", strArquivoSaida)

    DoCmd.Hourglass False
    Access.Application.Echo True
    
    If bAbrir Then Call VBA.Shell("explorer """ & strArquivoSaida & """")
    
Fim:
    On Error GoTo 0
    Exit Function

ExportarPara_Error:
    If VBA.Err <> 0 Then
        lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
        'Debug.Print cstr_ProcedureName, "Linha : " & VBA.Erl() & " - " & strErrorMessagem
        Call Excecoes.TratarErro(strErrorMessagem, lngErrorNumber, cstr_ProcedureName, VBA.Erl())
    End If
    GoTo Fim:
    'Debug Mode
    Resume
End Function

Function PegarTipoSaida(strTipo As String)
    Select Case LCase(strTipo)
    Case "xlsx": PegarTipoSaida = Access.Constants.acFormatXLSX
    Case "xlsb": PegarTipoSaida = Access.Constants.acFormatXLSB
    Case "xls": PegarTipoSaida = Access.Constants.acFormatXLS
    Case "html": PegarTipoSaida = Access.Constants.acFormatHTML
    Case "xps": PegarTipoSaida = Access.Constants.acFormatXPS
    Case "pdf": PegarTipoSaida = Access.Constants.acFormatPDF
    End Select
End Function

