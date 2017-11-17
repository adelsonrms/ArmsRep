Attribute VB_Name = "AuxForm"
Option Compare Database
Option Explicit
'Variáveis globais para uso em formulários

Public ctvControl As Control
Public frmCurrentForm As Form
Public currentParentControls As Object
Public dblIncremento As Double
Public primeiroControle As Object
Public cDadosOriginais As Collection

'Enum para definir algumas propriedades basicas de bloqueio de controles
Public Enum eFlagControl
    Visible_Control = 1
    Enable_Control = 2
    Locked_Control = 3
End Enum

Public Enum eRetornoItemsSelecionados
    NomeTabela = 1
    RetornaRecordSet = 2
    ArraySelecionados = 3
End Enum
'---------------------------------------------------------------------------------------
' Modulo....: AuxForm / Módulo
' Rotina....: AbrirFormulario / Sub
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 11/09/2013
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina publica para abrir os formularios atravez dos botoes da ribbon
'---------------------------------------------------------------------------------------
' + HISTÓRIO DE REVISÃO
'---------------------------------------------------------------------------------------
' DATA / DESCRIÇÃO
'-----------------------------------------------------------------------------------------------------------------------------------
' 10/02/2015 13:44 - Incluido os argumentos oFormParent e pFilter
'                    O argumento 'oFormParent' poderá ser utilizado para passar a instancia
'                    de um formulário antes do qual o form atual esta sendo chamado.
'                    O argumento 'pFilter' pode ser passado para filtrar o
'                    formulário na abertura com um criterio
' 28/07/2016 15:45 - Incluído controle de variaveis apenas recarregar, sem abrir novamente.
'-----------------------------------------------------------------------------------------------------------------------------------

Public Sub AbrirFormulario(ByVal tag As String, _
                           Optional ByVal View As Access.AcFormView = Access.AcFormView.acNormal, _
                           Optional ByVal WindowMode As Access.AcWindowMode = Access.AcWindowMode.acWindowNormal, _
                           Optional oFormParent As Form, _
                           Optional pFilter As String, _
                           Optional HideParent As Boolean)
10  On Error GoTo TratarErro
Dim arrTag As Variant
Dim frm As AccessObject

20  If HideParent Then
30      If Not oFormParent Is Nothing Then
40          oFormParent.Visible = False
50      End If
60  End If

70  arrTag = VBA.Split(tag, "|")
80  On Error Resume Next
90  Set frm = Access.CurrentProject.AllForms(arrTag(0))

100 If Not frm Is Nothing Then
110     If frm.IsLoaded Then
            'Força apenas o Reload do formulário (O formulário deve ter a função Reload() com as ações customizadas)
            If Not pegaValor("ReabrirForm") = "-1" Then GoTo Reload:
            
            'Se nao, fecha o form antes de reabri-lo
120         If frm.CurrentView = Access.AcCurrentView.acCurViewFormBrowse Then
130             Call Access.DoCmd.Close(Access.acForm, arrTag(0))
140         End If
150     End If
160 End If

170 If UBound(arrTag) > 0 Then
180     Call Access.DoCmd.openForm(FormName:=arrTag(0), View:=acNormal, OpenArgs:=arrTag(1), WindowMode:=WindowMode, WhereCondition:=pFilter)
190 Else
200     Call Access.DoCmd.openForm(FormName:=tag, View:=acNormal, WindowMode:=WindowMode, WhereCondition:=pFilter)
210     If WindowMode = acDialog Then If Not oFormParent Is Nothing Then oFormParent.Refresh
220 End If

230 If Not oFormParent Is Nothing Then
240     On Error Resume Next
        'Se o form tiver uma sub Reload()
250     Call oFormParent.Reload
260     Call oFormParent.Requery
270     If HideParent Then oFormParent.Visible = True
280 End If
  
Fim:
    Call salvaValor("ReabrirForm", 0)
  
290 On Error GoTo 0
300 Exit Sub

Reload:
    Call Access.Application.Forms(arrTag(0)).Reload
    GoTo Fim:
    
TratarErro:
310 Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxForm.AbrirFormulario", Erl)
End Sub

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : AuxForm.AbrirRelatorio()
' TIPO             : Sub
' DATA/HORA        : 08/06/2015 13:59
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Abre um relatório especificado
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      08/06/2015 13:59    Adelson         Criação/Atualização do procedimento
'---------------------------------------------------------------------------------------
Sub AbrirRelatorio(pReportName As String, Optional modoVisualizacao As AcView = acViewReport, Optional modoJanela As AcWindowMode = acDialog, Optional pFiltro As String, Optional oFormParent As Form)
10  On Error GoTo AbrirRelatorio_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
20  Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Sub AuxForm.AbrirRelatorio()"
    '----------------------------------------------------------------------------------------------------

30  Access.Application.DoCmd.OpenReport pReportName, modoVisualizacao, , pFiltro, modoJanela
40  If Not oFormParent Is Nothing Then oFormParent.Visible = True

Fim:
50  On Error GoTo 0
60  Exit Sub

AbrirRelatorio_Error:
70  If VBA.Err <> 0 Then
80      lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
90      Call Excecoes.TratarErro(strErrorMessagem, lngErrorNumber, cstr_ProcedureName, VBA.Erl)
100 End If
    GoTo Fim:
    'Debug Mode
110 Resume
End Sub

'---------------------------------------------------------------------------------------
' Modulo....: AuxForm / Módulo
' Rotina....: Atualizar_Campo_Versao / Sub
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 25/09/2013
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina que atualiza o campo de versão no formulário
'---------------------------------------------------------------------------------------
Public Sub Atualizar_Campo_Versao(ByRef lbl As Access.label)
10  On Error GoTo TratarErro

20  If versao = VBA.vbNullString Then Call Publicas.Carregar_Versao
30  lbl.Caption = VariaveisEConstantes.versao
40  lbl.TextAlign = 3

50  On Error GoTo 0
60  Exit Sub
TratarErro:
70  Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxForm.Atualizar_Campo_Versao", Erl)
End Sub
'---------------------------------------------------------------------------------------
' Modulo....: AuxForm / Módulo
' Rotina....: AutoFitColumns / Sub
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 15/01/2014
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina para fazer o autofit das colunas de um subformulario
'---------------------------------------------------------------------------------------
Public Sub AutoFitColumns(ByRef sFrm As SubForm)
On Error GoTo TratarErro
Dim intContador     As Integer
    With sFrm
        For intContador = 0 To .Form.Controls.count - 1 Step 1
            If VBA.TypeName(.Form(intContador)) = "TextBox" Or VBA.TypeName(.Form(intContador)) = "ComboBox" Then
                .Form(intContador).ColumnWidth = 32767 'MAX do integer
                .Form(intContador).ColumnWidth = -2
            End If
        Next intContador
    End With

On Error GoTo 0
Exit Sub
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxForm.AutoFitColumns", Erl)
End Sub
'---------------------------------------------------------------------------------------
' Modulo....: AuxForm / Módulo
' Rotina....: CarregaDataReferencia / Function
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 05/09/2013
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina para carregar a data desejada.
'---------------------------------------------------------------------------------------
Public Function CarregaDataReferencia(Optional ByVal dtRef As Date = 0, Optional ByVal TextoInformativo As String = "Informar Data", _
                                      Optional ByVal Mascara As String = "!00/00/0000;0;_") As Date
On Error GoTo TrataErro
Dim dtAux           As Variant
Dim objFormData     As Form

    Call salvaValor("DataSelecionada", "")
    Call Access.DoCmd.openForm(FormName:="frmEscolherData", View:=acNormal, OpenArgs:=Publicas.GerarChaves(VBA.CStr(dtRef), TextoInformativo, Mascara), WindowMode:=acDialog)
    
    dtAux = pegaValor("DataSelecionada")
    
    If Not VBA.IsDate(dtAux) Then
        Call AuxMensagens.MessageBoxMaster("F009")
    Else
        CarregaDataReferencia = dtAux
    End If
    
Exit Function
TrataErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxForm.CarregaDataReferencia()", Erl)
    Exit Function
    Resume
End Function


Function PegaDataSelecionada() As Date
    Dim dt
    dt = AuxAplicacao.pegaValor("DataSelecionada")
    If VBA.IsDate(dt) Then
        PegaDataSelecionada = CDate("01-" & VBA.Format(dt, "mm-yyyy"))
    Else
        PegaDataSelecionada = VBA.Date
    End If
End Function

Sub CancelarEdicao(ByVal formulario As Form)
    Dim ctlTextbox    As Control
    For Each ctlTextbox In formulario.Controls
        If ctlTextbox.controlType = acTextBox Then
            ctlTextbox.value = ctlTextbox.OldValue
        End If
    Next ctlTextbox
End Sub
'---------------------------------------------------------------------------------------
' Modulo....: Form_frmProdutosGrupo / Documento VBA
' Rotina....: Carregar_Controles / Sub
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 28/08/2013
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina para carregar as posições iniciais de cada controle
'---------------------------------------------------------------------------------------
Public Sub Carregar_Controles(ByRef frm As Form, ByRef Controles As Object)
On Error GoTo TratarErro
Dim cnt         As Access.Control
Dim img         As Access.image

    With frm
        For Each cnt In .Controls
            If VBA.TypeName(cnt) = "Image" Then
                If Not Controles.Exists(cnt.Name) Then
                    Set img = cnt
                    With img
                        Call Controles.Add(cnt.Name, Publicas.GerarChaves(.Top, .Left, .Height, .Width))
                    End With
                    Call Publicas.RemoverObjetosMemoria(img)
                End If
            End If
            Call Publicas.RemoverObjetosMemoria(cnt)
        Next cnt
    End With

On Error GoTo 0
Exit Sub
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxForm.Carregar_Controles", Erl)
End Sub

'---------------------------------------------------------------------------------------
' Modulo....: AuxForm / Módulo
' Rotina....: CentralizarControleVerticalmente / Function
' Autor.....: Jonathan Dantas
' Contato...: jonathan@tecnun.com.br
' Data......: 09/09/2014
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina para centralização das margens de um controle
'---------------------------------------------------------------------------------------
Public Function CentralizarControleVerticalmente()
On Error GoTo TratarErro
    Dim ctr As Control
    Dim altura As Double
    Dim tamFonte As Double
    Dim pixelEmCentimetro As Double
    Dim TwipEmCentimetro As Double

    Set ctr = Screen.ActiveControl

    pixelEmCentimetro = 0.02645833333333
    TwipEmCentimetro = 0.001763888888889
    tamFonte = ctr.FontSize * pixelEmCentimetro
    altura = ctr.Height
    
    If (((altura / 15) / 2) * pixelEmCentimetro - tamFonte) / TwipEmCentimetro > 0 Then
        ctr.TopMargin = (((altura / 15) / 2) * pixelEmCentimetro - tamFonte) / TwipEmCentimetro
    End If

On Error GoTo 0
Exit Function
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxForm.CentralizarControleVerticalmente", Erl)
End Function

'---------------------------------------------------------------------------------------
' Modulo....: AuxForm / Módulo
' Rotina....: Criar_Form / Sub
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 14/01/2014
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina que tem como objetivo criar um form a partir de uma consulta
'---------------------------------------------------------------------------------------
Public Sub Criar_Form(ByVal SubForm As String, ByVal Consulta As String, _
                      Optional ByVal Alinhamento As Byte = 0)
On Error GoTo TratarErro
Dim rsRel               As Object 'ADODB.Recordset
Dim strNomeRel          As String
Dim strCampos           As String
Dim strFormName         As String
Dim sfrmTemp            As Access.Form
Dim txt                 As Access.TextBox
Dim contador            As Integer

    Set rsRel = Conexao.PegarRS(Consulta)
    If rsRel Is Nothing Then GoTo Fim

    If Conexao.ObjetoExiste(Access.CodeDb, Access.AcObjectType.acForm, SubForm) Then
        Call Access.DoCmd.DeleteObject(Access.AcObjectType.acForm, SubForm)
    End If

    Set sfrmTemp = Access.CreateForm()

    With sfrmTemp
        .AllowAdditions = False
        .AllowDeletions = False
        .AllowEdits = False
        .DefaultView = 2
        With rsRel
            For contador = 0 To .Fields.count - 1 Step 1
                Set txt = Access.CreateControl(sfrmTemp.Name, Access.AcControlType.acTextBox, _
                                               Access.AcSection.acDetail)
                txt.ControlSource = .Fields(contador).Name
                txt.Name = .Fields(contador).Name
                txt.TextAlign = Alinhamento
            Next contador
        End With

        .RecordSource = Consulta
        Call Access.DoCmd.Save(Access.AcObjectType.acForm, sfrmTemp.Name)
        Call Access.DoCmd.CopyObject(, SubForm, acForm, sfrmTemp.Name)
        strFormName = .Name
        Call Access.DoCmd.Close(Access.AcObjectType.acForm, .Name, acSaveNo)
        Call Access.DoCmd.DeleteObject(Access.AcObjectType.acForm, strFormName)
    End With

Fim:
    Call Publicas.RemoverObjetosMemoria(rsRel, sfrmTemp)

On Error GoTo 0
Exit Sub
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxForm.CriaForm", Erl)
End Sub

Sub LimparProgressoEmForm(Optional formInstance As Form)
10  On Error Resume Next
20  If formInstance Is Nothing Then Set formInstance = Screen.ActiveForm
30  If formInstance Is Nothing Then Exit Sub
40  formInstance.lblDe.Visible = False
50  formInstance.lblDe.Caption = ""

60  formInstance.lblCaption.Visible = False
70  formInstance.lblCaption.Caption = ""

80  formInstance.Contorno.Visible = False
90  formInstance.Progresso.Visible = False
100 formInstance.Progresso.Width = 0
110 formInstance.Progresso.tag = ""

120 Call Access.DoCmd.RepaintObject(Access.AcObjectType.acForm, formInstance.Name)
    'SysCmd acSysCmdClearStatus
End Sub
Public Sub IncrementaBarraProgresso(ByVal dblIncremento As Double, Optional ByVal strStatus As String = VBA.vbNullString)
On Error Resume Next ' Em caso de erro seguir para o proximo item
Dim dblTamTotal         As Double
Dim dblValorAtual       As Double
Dim dblUnidade          As Double
Dim frmProg             As Form: Set frmProg = Forms("frmProgresso")
    
    If frmProg.IsLoaded Then
        dblTamTotal = frmProg.Contorno.Width
        dblUnidade = dblTamTotal / 100
        dblValorAtual = frmProg.Progresso.Width
        If dblIncremento <= 0 Then
            frmProg.Progresso.Width = 0
        ElseIf (dblIncremento * dblUnidade) + dblValorAtual >= dblTamTotal Then
            frmProg.Progresso.Width = dblTamTotal
        Else
            frmProg.Progresso.Width = dblValorAtual + (dblIncremento * dblUnidade)
        End If
        dblValorAtual = frmProg.Progresso.Width / dblUnidade
        If strStatus = VBA.vbNullString Then
            frmProg.Caption = "Progresso Importação - " & VBA.Format(dblValorAtual, "00.00") & "%"
        Else
            frmProg.Caption = strStatus & " - " & VBA.Format(dblValorAtual, "00.00") & "%"
        End If
        VBA.DoEvents
        Call Access.DoCmd.RepaintObject(Access.AcObjectType.acForm, "AuxForm.IncrementaBarraProgresso")
    End If
End Sub

Public Function ExibirProgresso(lngItem As Integer, Optional lntTotal As Long, Optional pMensagem As String)
Dim dblTamTotal         As Double
Dim dblValorAtual       As Double
Dim dblUnidade          As Double
Dim frmProg             As Form_frmProgresso
On Error GoTo Erro

    Set frmProg = Forms!frmProgresso
    
    If Access.CurrentProject.AllForms("frmProgresso").IsLoaded Then
        dblTamTotal = frmProg.Contorno.Width
        If lngItem > 0 Then frmProg.Progresso.Width = frmProg.Progresso.Width + (dblTamTotal / lntTotal)
        frmProg.lblStatus.Caption = VBA.Format(lngItem / lntTotal, "0%") & " / (" & lngItem & " de " & lntTotal & ")"
        frmProg.Caption = "Processando..." & pMensagem
        VBA.DoEvents
        Call Access.DoCmd.RepaintObject(Access.AcObjectType.acForm, "frmProgresso")
    End If
Fim:
    Exit Function
Erro:
If VBA.Err <> 0 Then
    VBA.MsgBox VBA.Error, VBA.vbCritical
End If
GoTo Fim

Resume

End Function


Sub LimparProgressoEmFormStatusBar()
    Call SysCmd(acSysCmdClearStatus)
    Call SysCmd(acSysCmdRemoveMeter)
End Sub

'Fecha um objeto (Seja relatório, formulário, tabela)
Public Sub FecharObjeto(pObjectName As String, Optional bSave As AcCloseSave)
    Dim pObjectType As AcObjectType
    If pObjectName = "" Then
        pObjectName = Access.Application.CurrentObjectName
        pObjectType = Access.Application.CurrentObjectType
    Else
        pObjectType = AuxForm.PegarTipoDeObjeto(pObjectName)
    End If
    Call DoCmd.Close(pObjectType, pObjectName, bSave)
    On Error Resume Next 'Nem sempre o objeto esta ativo para ser ocultado.
    If Not frmCurrentForm Is Nothing Then frmCurrentForm.Visible = True
End Sub

'Determina, de acordo com o nome do objeto, qual o seu tipo : Form, Query, Table, Report
Function PegarTipoDeObjeto(objName As String) As AcObjectType
    Dim objType As String
1   objType = CLng(Nz(DLookup("Type", "MSysObjects", "Name='" & objName & "'")))
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

'Recupera a instancia de um objeto de acordo com o tipo
Function PegarObjetoDBPorNome(objName As String) As Object
    Dim objType As AcObjectType
    Dim accObj As AccessObject
1   objType = PegarTipoDeObjeto(objName)
    'Dim objAcc As AccessObject
2   Select Case objType
    Case AcObjectType.acForm: Set accObj = Access.Application.CurrentProject.AllForms(objName):
    Case AcObjectType.acQuery: Set accObj = Access.Application.CurrentDb.QueryDefs(objName):
    Case AcObjectType.acTable: Set accObj = Access.Application.CurrentData.AllTables(objName):
    Case AcObjectType.acReport: Set accObj = Access.Application.CurrentProject.AllReports(objName):
    Case AcObjectType.acModule: Set accObj = Access.Application.CurrentProject.AllModules(objName)
12  End Select
13  Set PegarObjetoDBPorNome = accObj
End Function

Sub FecharTelaProgresso(Optional pProgressType As String = "List")
10  If pProgressType = "List" Then
20      Call Access.DoCmd.Close(Access.AcObjectType.acForm, "frmProgressoLista")
30  Else
40      Call Access.DoCmd.Close(Access.AcObjectType.acForm, "frmProgresso")
50  End If
    Call LimparProgressoEmFormStatusBar
End Sub

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : AuxForm.ExibirSubForm()
' TIPO             : Sub
' DATA/HORA        : 18/11/2014 18:05
' CONSULTOR        : (ADELSON)/ TECNUN - Adelson Rosendo Marques da Silva
' CONTATO          : adelson@tecnun.com.br
' DESCRIÇÃO        : Configura um objeto em um subformulário genericamente
'---------------------------------------------------------------------------------------
Sub ExibirSubForm(sfContainer As SubForm, objName As String)
    Dim objType As String
    '----------------------------------------------------------------------------------------------------
10  On Error GoTo ExibirSubForm_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
20  Dim dtSartRunProc As Date: dtSartRunProc = Time
    Const cstr_ProcedureName As String = "Sub prjDRE.AuxForm.ExibirSubForm()"
    '----------------------------------------------------------------------------------------------------

30  objType = Nz(DLookup("Type", "MSysObjects", "Name='" & objName & "'"))

40  Select Case objType
        Case -32768
50          sfContainer.SourceObject = "Form." & objName
60      Case 5
70          sfContainer.SourceObject = "Query." & objName
80      Case 1, 6
90          sfContainer.SourceObject = "Table." & objName
100     Case -32764
110         sfContainer.SourceObject = "Report." & objName
120 End Select

130 On Error GoTo 0
140 Exit Sub

ExibirSubForm_Error:
150 If VBA.Err <> 0 Then
160     lngErrorNumber = VBA.Err.Number: strErrorMessagem = "Ocorreu um erro na configuração do subformulçário " & VBA.vbNewLine & VBA.Err.Description
        'Tratamento de erro personalizado
170     Call Excecoes.TratarErro(strErrorMessagem, lngErrorNumber, cstr_ProcedureName, VBA.Erl)
180 End If
190 Exit Sub
    'Debug Mode
200 Resume
End Sub

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : AuxForm.ExibirSubFormSelecionarItems()
' TIPO             : Sub
' DATA/HORA        : 08/06/2015 14:16
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Exibe o conteudo no sub formulário da tela de seleção de varios itens
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      08/06/2015 14:16    Adelson         Criação/Atualização do procedimento
'---------------------------------------------------------------------------------------
Sub ExibirSubFormSelecionarItems(objName As String, pFilter As String)
    Dim args As String
    Dim objType As String, pTable As String

10  On Error GoTo ExibirSubFormSelecionarItems_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
20  Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Sub AuxForm.ExibirSubFormSelecionarItems()"
    '----------------------------------------------------------------------------------------------------

30  objType = DLookup("Type", "MSysObjects", "Name='" & objName & "'")
40  Select Case objType
    Case -32768
50      If Access.Application.LanguageSettings.LanguageID(msoLanguageIDUI) = 1046 Then
60          pTable = "Formulário." & objName
70      Else
80          pTable = "Form." & objName
90      End If
100 Case 5
110     If Access.Application.LanguageSettings.LanguageID(msoLanguageIDUI) = 1046 Then
120         pTable = "Consulta." & objName
130     Else
140         pTable = "Query." & objName
150     End If
160 Case 6
170     If Access.Application.LanguageSettings.LanguageID(msoLanguageIDUI) = 1046 Then
180         pTable = "Tabela." & objName
190     Else
200         pTable = "Table." & objName
210     End If
220 Case -32764
230     If Access.Application.LanguageSettings.LanguageID(msoLanguageIDUI) = 1046 Then
240         pTable = "Relatório." & objName
250     Else
260         pTable = "Report." & objName
270     End If

280 End Select
290 args = pTable & ";" & pFilter
300 Call AbrirFormulario("frmSelecionarItemsTable|" & args, _
       , acDialog, _
                         frmCurrentForm)

Fim:
310 On Error GoTo 0
320 Exit Sub

ExibirSubFormSelecionarItems_Error:
330 If VBA.Err <> 0 Then
340     lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
350     Call Excecoes.TratarErro(strErrorMessagem, lngErrorNumber, cstr_ProcedureName, VBA.Erl)
360 End If
    GoTo Fim:
    'Debug Mode
370 Resume
End Sub

'Variáveis publicas necessárias
'----------------------------------------------------
'Public ctvControl As control
'Public frmCurrentForm As Form
'----------------------------------------------------

Sub AtualizaStatusBarAccess(counter As Integer)
    On Error Resume Next
    Call SysCmd(acSysCmdUpdateMeter, counter)
End Sub



'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : AuxForm.DefineInstanciaControleAtual()
' TIPO             : Function
' DATA/HORA        : 28/08/2014 09:39
' CONSULTOR        : (ADELSON)/ TECNUN - Adelson Rosendo Marques da Silva
' CONTATO          : adelson@tecnun.com.br
' DESCRIÇÃO        : Define a instancia atual do formulário e controle ativo
'---------------------------------------------------------------------------------------
Function DefineInstanciaControleAtual() As Variant
10  On Error Resume Next
20  Set frmCurrentForm = Screen.ActiveForm
30  Set ctvControl = frmCurrentForm.ActiveControl
End Function
'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : AuxForm.DefiniPropriedadeControles()
' TIPO             : Sub
' DATA/HORA        : 07/04/2014 18:50
' CONSULTOR        : (ADELSON)/ TECNUN - Adelson Rosendo Marques da Silva
' CONTATO          : adelson@tecnun.com.br
' DESCRIÇÃO        : Define o valor de um propriedade em varios controles de uma so vez em um container
'---------------------------------------------------------------------------------------
Sub DefiniPropriedadeControles(objContainer As Object, _
                             pPropertyName As String, _
                             Optional pValue As Variant, _
                             Optional typeControl As String = "*", _
                             Optional hasName As String = "*", _
                             Optional FilterByTag As String = "*", _
                             Optional Ignorar As String = "*")
10  On Error Resume Next
    Dim ctr As Control

20  If objContainer Is Nothing Then Exit Sub

30  For Each ctr In objContainer.Controls
40      If VBA.TypeName(ctr) Like "*" & typeControl & "*" Then
            With ctr
                If InStr(Ignorar, ctr.Name) > 0 And Ignorar <> "*" Then GoTo Proximo
50              If .Name Like "*" & hasName & "*" Then
60                  If .tag Like "*" & FilterByTag & "*" Then
                        If ctr.Visible Then
                            If LCase(pPropertyName) = "value" Then
                                ctr.value = pValue
                            Else
                                ctr.Properties(pPropertyName).value = pValue
                            End If
                        End If
80                  End If
90              End If
            End With
100     End If
Proximo:
110 Next

End Sub

Public Sub IncrementaBarraProgressoCompactar(ByVal QTD As Long, ByVal TOTQTD As Long, Optional ByVal strStatus As String = VBA.vbNullString)
On Error Resume Next
Dim dblTamTotal         As Double
Dim dblValorAtual       As Double
Dim dblIncremento       As Double
Dim frmProg             As Form: Set frmProg = Forms("frmProgresso")

    If Access.CurrentProject.AllForms("frmProgressBar").IsLoaded Then
        dblTamTotal = frmProg.Contorno.Width
        dblValorAtual = frmProg.Progresso.Width
        dblIncremento = (dblTamTotal / TOTQTD) * QTD
        frmProg.Progresso.Width = dblIncremento
        dblValorAtual = (100 / TOTQTD) * QTD
        If strStatus = VBA.vbNullString Then
            frmProg.Caption = "Progresso Importação - " & VBA.Format(dblValorAtual, "00.00") & "%"
        Else
            frmProg.Caption = strStatus & " - " & VBA.Format(dblValorAtual, "00.00") & "%"
        End If
        VBA.DoEvents
        Call Access.DoCmd.RepaintObject(Access.AcObjectType.acForm, "AuxForm.IncrementaBarraProgressoCompactar")
    End If
End Sub

'********************SUB-ROTINAS********************
'---------------------------------------------------------------------------------------
' Modulo....: AuxForm / Módulo
' Rotina....: IncrementaBarraProgresso / Sub
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 22/08/2013
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina para incrementar a barra de progresso, alterando o width da imagem(progresso)
'---------------------------------------------------------------------------------------
Public Sub IncrementaBarraProgressoLista(ByVal dblIncremento As Double, ByVal Titulo As String, _
                                         ByVal ItemLista As String, ByVal StatusAnterior As String)
On Error Resume Next
Dim frm                 As Access.Form
Dim dblTamTotal         As Double
Dim dblValorAtual       As Double
Dim dblUnidade          As Double
Dim ItemAnterior        As String
Dim frmProg             As Form: Set frmProg = Forms("frmProgressoLista")

    If Access.CurrentProject.AllForms("frmProgressoLista").IsLoaded Then
        Set frm = Forms!frmProgressoLista
        With frmProg
            dblTamTotal = .Contorno.Width
            dblUnidade = dblTamTotal / 100
            dblValorAtual = .Progresso.Width
            If dblIncremento <= 0 Then
                .Progresso.Width = 0
            ElseIf (dblIncremento * dblUnidade) + dblValorAtual >= dblTamTotal Then
                .Progresso.Width = dblTamTotal
            Else
                .Progresso.Width = dblValorAtual + (dblIncremento * dblUnidade)
            End If
            dblValorAtual = .Progresso.Width / dblUnidade
            If Titulo = VBA.vbNullString Then
                .lblTitulo.Caption = "Progresso Importação - " & VBA.Format(dblValorAtual, "00.00") & "%"
            Else
                .lblTitulo.Caption = Titulo & " - " & VBA.Format(dblValorAtual, "00.00") & "%"
            End If
            
            If Not ItemLista = VBA.vbNullString Then
                With .lstItems
                    If .ListCount > 0 Then
                        ItemAnterior = .ItemData(.ListCount - 1)
                        Call .RemoveItem(.ListCount - 1)
                        .Selected(.ListCount - 1) = True
                        If StatusAnterior = VBA.vbNullString Then
                            Call .addItem("'" & ItemAnterior & "';'OK'")
                        Else
                            Call .addItem("'" & ItemAnterior & "';'" & StatusAnterior & "'")
                        End If
                    End If
                    .Selected(.ListCount - 1) = True
                    Call .addItem("'" & ItemLista & " " & VBA.String$(100, ".") & "';''")
                    '.ListIndex = .ListCount
                    .Selected(.ListCount - 1) = True
                End With
            End If
        End With
        VBA.DoEvents
        Call Access.DoCmd.RepaintObject(Access.AcObjectType.acForm, "frmProgressoLista")
    End If
End Sub

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : AuxForm.inserirLogAlteracoes()
' TIPO             : Sub
' DATA/HORA        : 02/02/2015 17:50
' CONSULTOR        : (Jonathan)/ TECNUN - Jonathan Rocha Figueroa Dantas
' CONTATO          :  jonathan@tecun.com.br
' DESCRIÇÃO        : Esta rotina tem como objetivo inserir em uma tabela (tblLogAlteracoes) todos os campos alterados,
'                    incluindo o valor antigo, valor atual, hora e usuario que alterou o
'                    registro de um formulário com os campos acoplados.
'---------------------------------------------------------------------------------------


Sub inserirLogAlteracoes(ByVal formulario As Form, Optional ByVal rowIDTabela As Long, Optional registro As Variant)

    Dim rs          As Object ' Recordset
    Dim arrDados    As Variant
    Dim contador    As Integer

   On Error GoTo inserirLogAlteracoes_Error

    Set rs = CodeDb.OpenRecordset("tblLogAlteracoes")
    arrDados = AuxForm.verificarCamposAlterados(formulario)
    If VBA.IsArray(arrDados) And UBound(arrDados) > 0 Then
    With rs
        For contador = 0 To UBound(arrDados, 2)
            .addNew
            .Fields("NomeTabela") = formulario.RecordSource
            .Fields("IDRegistro") = registro
            .Fields("rowIDTabela") = rowIDTabela
            .Fields("NomeCampo") = arrDados(0, contador)
            .Fields("ValorAntigo") = arrDados(1, contador)
            .Fields("ValorNovo") = arrDados(2, contador)
            .Fields("UsuarioAlteracao") = VBA.Environ("ComputerName") & "\" & VBA.Environ("Username")
            .Update
        Next contador
    End With
    End If

   On Error GoTo 0
   Exit Sub

inserirLogAlteracoes_Error:
   If VBA.Err <> 0 Then Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxForm.inserirLogAlteracoes()", VBA.Erl)
   Exit Sub

End Sub
'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : AuxForm.SelecionarItems()
' TIPO             : Function
' DATA/HORA        : 11/05/2015 16:02
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Exibe um formulário com estilo de seleção de listas permitindo multipla seleção. Recupera os items selecionados para um Array
'                    source - Deve ser uma tabela ou consulta que deve ter pelo menos 2 campos obrigatorios : Codigo e Descrição. Os demais campos que houverem nao serão exibidos
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      11/05/2015 16:02    Adelson         Criação/Atualização do procedimento
'---------------------------------------------------------------------------------------
Function SelecionarItems(source As String, _
                     Optional bMultiSelect As Boolean = True, _
                     Optional eRetorno As eRetornoItemsSelecionados = eRetornoItemsSelecionados.ArraySelecionados, _
                     Optional dbInst As Object, _
                     Optional bSelecionarTodos As Boolean = True) As Variant
    Dim fsi    As Form
    Dim db   As Object
    Dim sTempTable As String
    Dim sqlSource As String
    Dim rs     As Object

10  On Error GoTo SelecionarItems_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
20  Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Function AuxForm.SelecionarItems()"
    '----------------------------------------------------------------------------------------------------
    
    If dbInst Is Nothing Then
        Set db = CodeDb
    Else
        Set db = dbInst
    End If

    sTempTable = AuxTabela.CriarTabelaTemporariaDDL(source, db)
    
50  If FormularioEstaAberto("frmSelecionarItems") Then Call FecharObjeto("frmSelecionarItems", acSaveYes)
    'Abre o formulário
100 Call AuxForm.AbrirFormulario("frmSelecionarItems|" & sTempTable & ";" & eRetorno & ";" & CInt(bMultiSelect) & ";" & CInt(bSelecionarTodos), acNormal, acDialog)

    'Se for tabela, recupera apenas o nome da tabela temporária. Nao á exclui
110 If eRetorno = NomeTabela Then
120     SelecionarItems = CStr(sTempTable)
        'Se for recordset, retorna a instancia do recordset somente os selecionados
130 ElseIf eRetorno = RetornaRecordSet Then
140     Set SelecionarItems = db.OpenRecordset("SELECT tb.* FROM " & sTempTable & " tb WHERE tb.selecionar = -1")
150     If Conexao.ObjetoExiste(db, acTable, sTempTable) Then db.Execute ("DROP TABLE " & sTempTable)
        'Se for Array, recupera da propriedade ItemsSelecionados do formulario ainda em memoria, porem oculto.
        'Fecha o formulario na sequencia
160 ElseIf eRetorno = ArraySelecionados Then
170     Set rs = db.OpenRecordset("SELECT tb.* FROM [" & sTempTable & "] tb WHERE tb.selecionar = -1")
180     If Not rs.EOF Then
190         rs.MoveLast: rs.MoveFirst
200         SelecionarItems = rs.GetRows(rs.RecordCount)
210     End If
220     rs.Close
230     Set rs = Nothing
240     Call AuxTabela.ExcluiTabela(sTempTable, db)
250 End If

Fim:
260 On Error GoTo 0
270 Exit Function

SelecionarItems_Error:
280 If VBA.Err <> 0 Then
290     lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
300     Call Excecoes.TratarErro(strErrorMessagem, lngErrorNumber, cstr_ProcedureName, VBA.Erl)
310 End If
    GoTo Fim:
    'Debug Mode
320 Resume

End Function


Function SelecionarEstadoCidade() As Variant
    Dim fsi As Form
    Dim db As Object
    Dim rs As Object

1   On Error GoTo SelecionarItems_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
2   Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Function AuxForm.SelecionarItems()"
    '----------------------------------------------------------------------------------------------------

3   Set db = CurrentDb

4   If FormularioEstaAberto("frmEscolherEstadoCidade") Then Call FecharObjeto("frmEscolherEstadoCidade", acSaveYes)
    Call db.Execute("UPDATE tblCidades SET selecionar = 0")
    'Abre o formulário
5   Call AuxForm.AbrirFormulario("frmEscolherEstadoCidade", acNormal, acDialog)

6   Set rs = db.OpenRecordset("SELECT tb.* FROM [tblCidades] tb WHERE tb.selecionar = -1")
7   If Not rs.EOF Then
8       rs.MoveLast: rs.MoveFirst
9       SelecionarEstadoCidade = rs.GetRows(rs.RecordCount)
10  End If
11  rs.Close
12  Set rs = Nothing

Fim:
13  On Error GoTo 0
14  Exit Function

SelecionarItems_Error:
15  If VBA.Err <> 0 Then
16      lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
17      Call Excecoes.TratarErro(strErrorMessagem, lngErrorNumber, cstr_ProcedureName, VBA.Erl)
18  End If
    GoTo Fim:
    'Debug Mode
19  Resume

End Function

'Abre a lista de seleção de varios items em uma tela
Function PegarSQLItemsSelecionados(strOrigem As String, strConsultaModeloSelecao As String, strCampoFiltroSelecao As String, Optional bSelecionarTodos As Boolean = True)
    Dim vListaSelecionados As Variant
    Dim strDinamicSQL As String
    Dim vPegarSelecao As Variant
    Dim strFiltroWhere As String
    Dim intIndex As Integer
    '---------------------------------------------------------------------------------------
    On Error GoTo PegarSQLItemsSelecionados_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
    Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Function BLRotina.PegarSQLItemsSelecionados()"
    '---------------------------------------------------------------------------------------
    
    Dim vSelecao As Variant
    vSelecao = Array(False, VBA.vbNullString)
    If Conexao Is Nothing Then Call Publicas.Inicializar_Globais(False)
        
    'Abre o formulario para seleção
    '-----------------------------------------------------------------------------------------------
    strDinamicSQL = VBA.Replace(Conexao.PegarComandoSQLModelo(strOrigem, CurrentProject.Connection), ";", "")
    vListaSelecionados = AuxForm.SelecionarItems(source:=strDinamicSQL, bSelecionarTodos:=bSelecionarTodos)
    If VBA.IsEmpty(vListaSelecionados) Then
        vSelecao = Array(-2, VBA.vbNullString)
        GoTo Fim
    End If
       
    'Monta o SQL com os items filtrados
    '-----------------------------------------------------------------------------------------------
    For intIndex = 0 To UBound(vListaSelecionados, 2)
        strFiltroWhere = strFiltroWhere & ",'" & vListaSelecionados(1, intIndex) & "'"
    Next intIndex
    strFiltroWhere = Trim(Mid(strFiltroWhere, 2))
    strDinamicSQL = Conexao.PegarComandoSQLModelo(strConsultaModeloSelecao, CurrentProject.Connection, "[@WhereIn]", VBA.IIf(strFiltroWhere <> "", " " & strCampoFiltroSelecao & " IN(" & strFiltroWhere & ")", ""))
    
    'Retorno do SQL com os items a serem filtrados
    vSelecao = Array(True, strDinamicSQL)

Fim:
    PegarSQLItemsSelecionados = vSelecao
90  On Error GoTo 0
100 Exit Function

PegarSQLItemsSelecionados_Error:
110 If VBA.Err <> 0 Then
120     lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
        'Debug.Print cstr_ProcedureName, "Linha : " & VBA.Erl() & " - " & strErrorMessagem
130     Call Excecoes.TratarErro(strErrorMessagem, lngErrorNumber, cstr_ProcedureName, VBA.Erl())
        vSelecao = Array(True, VBA.vbNullString)
140 End If
    
    GoTo Fim:
    'Debug Mode
150 Resume

End Function

Function FiltraMultiSelecao(objCtr As Object, Optional bsalvaValor As Boolean = True)
    Dim vSelected As Variant, items_selecionados As String
    Dim strSQL As String
    Dim iColunas As Variant
    Dim selecao As Integer
    strSQL = AuxRibbon.getTheValue(objCtr.tag, "SQLSelecionarItems")
    iColunas = AuxRibbon.getTheValue(objCtr.tag, "QtdCols")
    vSelected = SelecionarItems(strSQL, True, ArraySelecionados)
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
' Modulo....: Form_frmProdutosGrupo / Documento VBA
' Rotina....: MoverControle / Sub
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 28/08/2013
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina para simular o evento do click do mouse como se estivesse apertando
'             um botão. as posições serão: 0 = Cima, 1 = Left, 2 = Height e 3 = Width
'---------------------------------------------------------------------------------------
Public Sub MoverControle_Setas(ByRef img As Access.image, ByRef Controles As Object, _
                               ByVal QtdDeslocamento As Integer, ByVal Posicao As Byte)
On Error GoTo TratarErro
Dim arrAux          As Variant
Dim ValorOriginal   As Double

    With img
        If Controles.Exists(.Name) Then
            arrAux = VBA.Split(Controles.item(.Name), "|")
            ValorOriginal = VBA.CDbl(arrAux(Posicao))
            Select Case Posicao
                Case 0
                    .Top = ValorOriginal + QtdDeslocamento
                Case 1
                    .Left = ValorOriginal + QtdDeslocamento
                Case 2
                    .Height = ValorOriginal + QtdDeslocamento
                Case 3
                    .Width = ValorOriginal + QtdDeslocamento
            End Select
        End If
    End With
On Error GoTo 0
Exit Sub
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxForm.MoverControle", Erl)
Resume
End Sub


'---------------------------------------------------------------------------------------
' Modulo....: AuxForm / Módulo
' Rotina....: OrganizaArgs() / Function
' Autor.....: Victor Félix
' Contato...: victor.santos@mondial.com.br
' Data......: 03/04/2012
' Empresa...: Mondial Tecnologia em Informática LTDA.
' Descrição.: Rotina que organiza os parâmetros inseridos na função "MessageBoxMaster"
' que serão passados para a abertura do formulario MessageBox
'---------------------------------------------------------------------------------------
Private Function OrganizaArgs(ByVal strTitulo As String, ByVal strMensagem As String, ByVal TipoMsg As VbMsgBoxStyle) As String
On Error GoTo TrataErro
    OrganizaArgs = strTitulo & "|" & strMensagem & "|" & TipoMsg
Exit Function
TrataErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxForm.OrganizaArgs()", Erl)
End Function

'---------------------------------------------------------------------------------------
' Modulo....: Form_frmProdutosGrupo / Documento VBA
' Rotina....: Carregar_Controles / Sub
' Autor.....: Jonathan Dantas
' Contato...: Jonathan@tecnun.com.br
' Data......:
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.:
'             Rotina utilizada para retornar um array com todos os registros que são selecionados dentro de um subform(em modo Folha de Dados), o array retorna com todas as colunas do subform.
'             DICA Para pegar o primeiro registro Selecionado em um subform ou Pegar a quantidade de registro selecionados utilizem a seguinte função:
'             primeiroRegistroSelecionado = subformulario.SelTop
'             QuantidadeRegistroSelecionado = subformulario.SelHeight
'---------------------------------------------------------------------------------------
    
Function PegarRegistrosSelecionadosSubform(ByVal formulario As Form, ByVal primeiroRegistro As Long, ByVal QuantidadeRegistros As Long) As Variant

    Dim arrSelecionados()     As Variant
    Dim rs                  As Object 'Recordset
    Dim contadorLinhas      As Long
    Dim quantidadeCampos    As Long
    Dim contadorColunas     As Long
    

   On Error GoTo PegarRegistrosSelecionadosSubform_Error

    Set rs = formulario.RecordsetClone
    quantidadeCampos = rs.Fields.count
    If rs.Fields.count = 0 Then
        PegarRegistrosSelecionadosSubform = "Nenhuma tabela selecionada "
        Exit Function
    ElseIf rs.RecordCount = 0 Then
        PegarRegistrosSelecionadosSubform = "Nenhum registro na tabela ou nenhuma empresa selecionada"
        Exit Function
    ElseIf primeiroRegistro = 0 Or QuantidadeRegistros = 0 Then
        PegarRegistrosSelecionadosSubform = "Nenhum registro selecionado"
        Exit Function
    End If
    rs.MoveFirst
    rs.Move (primeiroRegistro - 1)
    ReDim arrSelecionados(0 To QuantidadeRegistros - 1, 0 To quantidadeCampos - 1)
    
    For contadorLinhas = 0 To QuantidadeRegistros - 1
        For contadorColunas = 0 To quantidadeCampos - 1
            arrSelecionados(contadorLinhas, contadorColunas) = rs.Fields(contadorColunas)
        Next contadorColunas
        rs.MoveNext
    Next contadorLinhas

PegarRegistrosSelecionadosSubform = arrSelecionados

   On Error GoTo 0
   Exit Function

PegarRegistrosSelecionadosSubform_Error:
   If VBA.Err <> 0 Then Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxForm.PegarRegistrosSelecionadosSubform()", VBA.Erl)
   Exit Function
   

End Function

'---------------------------------------------------------------------------------------
' Modulo....: AuxForm / Modulo
' Rotina....: PreencherListBoxComMatriz / Sub
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 22/08/2013
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina para preencher o listBox com o conteúdo de uma matriz
'---------------------------------------------------------------------------------------
Private Sub PreencherListBoxComMatriz(ByRef lst As Access.ListBox, ByVal arrDados As Variant)
On Error GoTo TrataErro
Dim valores         As String
Dim ContLinhas      As Long
Dim ContColunas     As Integer

    If VBA.IsArray(arrDados) Then
        With lst
            .ColumnCount = UBound(arrDados, 2) + 1
            .RowSourceType = "Value List"
            For ContLinhas = 0 To UBound(arrDados, 1) Step 1
                valores = VBA.vbNullString
                For ContColunas = 0 To UBound(arrDados, 2) Step 1
                    valores = valores & VBA.IIf(VBA.IsNumeric(arrDados(ContLinhas, ContColunas)), _
                                            VBA.Format(arrDados(ContLinhas, ContColunas), "hh:MM:ss"), _
                                            arrDados(ContLinhas, ContColunas)) & ";"
                Next ContColunas
                valores = VBA.Left(valores, VBA.Len(valores) - 1)
                .addItem valores
            Next ContLinhas
        End With
    End If
Exit Sub
TrataErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxForm.PreencherListBoxComMatriz()", Erl)
End Sub

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : AuxForm.verificarCamposAlterados()
' TIPO             : Function
' DATA/HORA        : 02/02/2015 17:49
' CONSULTOR        : (Jonathan)/ TECNUN - Jonathan Rocha Figueroa Dantas
' CONTATO          :  jonathan@tecun.com.br
' DESCRIÇÃO        : Esta rotina tem como objetivo devolver um array com os
'                    campos que tiveram seus valores alterados dentro de um formulario acoplado....
'---------------------------------------------------------------------------------------

Function verificarCamposAlterados(ByVal formulario As Form, ParamArray NomeControles() As Variant) As Variant
    
    Dim controle    As Control
    Dim arrDados()    As Variant
    Dim contador    As Integer

   On Error GoTo verificarCamposAlterados_Error

    ReDim arrDados(0 To 2, 0 To 100)
    
    If VBA.IsArray(NomeControles) And UBound(NomeControles) > 0 Then
        For contador = 0 To UBound(NomeControles)
            If formulario.Controls(NomeControles(contador)).OldValue <> formulario.Controls(NomeControles(contador)).value Then
                arrDados(0, contador) = NomeControles(contador)
                arrDados(1, contador) = formulario.Controls(NomeControles(contador)).OldValue
                arrDados(2, contador) = formulario.Controls(NomeControles(contador)).value
            End If
        Next
        ReDim Preserve arrDados(0 To 2, 0 To contador)
        verificarCamposAlterados = arrDados
        Exit Function
    End If
    For Each controle In formulario.Controls
        If controle.controlType = acComboBox Or controle.controlType = acTextBox Or controle.controlType = acCheckBox Then
            If controle.OldValue <> controle.value Then
                arrDados(0, contador) = controle.Name
                arrDados(1, contador) = controle.OldValue
                arrDados(2, contador) = controle.value
                contador = contador + 1
            End If
        End If
    Next controle
    
    If contador <> 0 Then
    ReDim Preserve arrDados(0 To 2, 0 To contador - 1)
    verificarCamposAlterados = arrDados

    Else
        ReDim arrDados(0 To 0, 0 To 0)
        verificarCamposAlterados = arrDados
        Exit Function
    End If


   On Error GoTo 0
   Exit Function

verificarCamposAlterados_Error:
   If VBA.Err <> 0 Then Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxForm.verificarCamposAlterados()", VBA.Erl)
   Exit Function


End Function

'Verifica se um controle existe em determinado formulário
'06/01/2015 - Adelson Silva
Function ControleExiste(objParent As Object, ctrName As String) As Boolean
    Dim C As Object
    On Error Resume Next
    Set C = objParent.Controls(ctrName)
    ControleExiste = VBA.Err.Number = 0 And Not C Is Nothing
    Set C = Nothing
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : AuxForm.FormularioEstaAberto()
' TIPO             : Function
' DATA/HORA        : 18/11/2014 18:27
' CONSULTOR        : (ADELSON)/ TECNUN - Adelson Rosendo Marques da Silva
' CONTATO          : adelson@tecnun.com.br
' DESCRIÇÃO        : Verifica se um formulário esta aberto
'---------------------------------------------------------------------------------------
Function FormularioEstaAberto(objName As String) As Boolean
    On Error Resume Next
    Dim frm    As AccessObject
    Dim cp     As Object
60  Set frm = Access.CurrentProject.AllForms(objName)
70  If Not frm Is Nothing Then
80      FormularioEstaAberto = frm.IsLoaded And Access.Application.Forms(frm.Name).Visible
110 End If
End Function

'Atualiza os icones dos arquivos (caso tenha no formulario)
'Deve haver as imagens com os nomes imgErro para Erro e imgOK para Aviso
Public Sub AtualizaFlagStatusArquivo(frm As Object, Optional bOK As Boolean = False)
    If ControleExiste(frm, "imgErro") And ControleExiste(frm, "imgOK") Then
40      frm.imgErro.Visible = Not bOK
50      frm.imgOK.Visible = Not frm.imgErro.Visible
    End If
End Sub

Function ReloadAllForms()
    On Error Resume Next 'Pode haver formulários que não tenham o função personalizada Reload()
    Dim frm    As Object
    Dim cp     As Object
60  For Each frm In Access.Application.Forms
70      If frm.IsLoaded Then
            Call frm.Reload: Call frm.Reqyery
110     End If
    Next frm
End Function

Public Sub preencheCampoData(ByVal controleTexto As Control, ByVal KeyAsC As Integer, Optional ByVal dataMinima As Date = "01/01/2000", Optional ByVal dataMaxima As Date = "01/01/2050")
On Error GoTo TratarErro
Dim caracter As String
Dim DataAux As String

    caracter = Chr$(KeyAsC)
    
    If Len(controleTexto.Text) >= 10 And KeyAsC <> 8 Then Exit Sub
    If (caracter < "0" Or caracter > "9") And KeyAsC <> 8 Then Exit Sub

    If Not KeyAsC = 8 Then
        controleTexto.Text = controleTexto.Text & caracter
    End If
    
    If Len(controleTexto.Text) = 2 And KeyAsC <> 8 And InStr(1, controleTexto.Text, "/") = 0 Then
            If CInt(controleTexto.Text) > 31 Then
                Call AuxMensagens.MessageBoxMaster("F011")
                controleTexto.Text = ""
                Exit Sub
            End If
            controleTexto.Text = controleTexto.Text & "/"
    ElseIf Len(controleTexto.Text) = 5 And KeyAsC <> 8 Then
        DataAux = VBA.Left(controleTexto.Text, 3)
        If Mid(controleTexto.Text, 4, 2) > 12 Then
            Call AuxMensagens.MessageBoxMaster("F011")
            
            controleTexto.Text = DataAux
            controleTexto.SelStart = Len(controleTexto.Text)
            Exit Sub
        End If
            controleTexto.Text = controleTexto.Text & "/"
    
    ElseIf Len(controleTexto.Text) = 10 Then
        DataAux = VBA.Left(controleTexto.Text, 6)
        If Not VBA.IsDate(controleTexto.Text) Then
            controleTexto.Text = ""
            controleTexto.SelStart = 0
            Call AuxMensagens.MessageBoxMaster("F011")
            Exit Sub
        End If
        If Not VBA.IsNull(dataMinima) Or dataMinima <> 0 Then
            If dataMinima > CDate(controleTexto.Text) Then
                controleTexto.Text = ""
                controleTexto.SelStart = 0
                Call AuxMensagens.MessageBoxMaster("S008", "Data informada é menor do que a permitida" & VBA.vbNewLine & "Data minima: " & dataMinima)
                Exit Sub
            End If
        End If
        If Not VBA.IsNull(dataMaxima) Or dataMaxima <> 0 Then
            If dataMaxima < CDate(controleTexto.Text) Then
                controleTexto.Text = ""
                controleTexto.SelStart = 0
                Call AuxMensagens.MessageBoxMaster("S008", "Data informada é maior do que a permitida" & VBA.vbNewLine & "Data máxima: " & dataMaxima)
                Exit Sub
            End If
       End If
    End If
    controleTexto.SelStart = Len(controleTexto.Text)

Fim:

On Error GoTo 0
Exit Sub
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxForm.preencheCampoData", Erl)
    GoTo Fim
End Sub


'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : Form_frmEditCadastroCorrespondente.validarCamposRequeridos()
' TIPO             : Function
' DATA/HORA        : 14/06/2015 20:36
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Analisa os campos e verifica se há informações de preenchimento  obrigatório em branco
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      14/06/2015 20:36    Adelson         Criação/Atualização do procedimento
'---------------------------------------------------------------------------------------
Function validarCamposRequeridos(frm As Object, Optional ByRef primeiroControle As Object) As Variant
    Dim ctr As Control, sCampos As String
    Dim contarErros As Integer
10  On Error GoTo validarCamposRequeridos_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
20  Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Function Form_frmEditCadastroCorrespondente.validarCamposRequeridos()"
    '----------------------------------------------------------------------------------------------------
30  For Each ctr In frm.Controls
40      If ctr.controlType = acComboBox Or ctr.controlType = acTextBox Or ctr.controlType = acCheckBox Then
50          If ctr.tag Like "*CampoMonitorado=True*" Then
60              On Error Resume Next
100             ctr.BorderColor = 0
120             If (ctr.tag Like "*Obrigatorio=True*") And Nz(ctr.value) = "" Then
130                 ctr.BorderColor = VBA.vbRed
140                 If contarErros = 0 Then Set primeiroControle = ctr
150                 contarErros = contarErros + 1
                    sCampos = sCampos & "Campo : " & ctr.Name & VBA.vbNewLine
160             ElseIf (ctr.tag Like "*Obrigatorio=True*") And Nz(ctr.value) <> "" Then
170                 ctr.BorderColor = 5026082
180             Else
190                 ctr.BorderColor = 0
210             End If
230         End If

240     End If
250 Next
260 validarCamposRequeridos = Array(contarErros, sCampos, primeiroControle)
Fim:
270 On Error GoTo 0
280 Exit Function

validarCamposRequeridos_Error:
290 If VBA.Err <> 0 Then
300     lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
310     Call Excecoes.TratarErro(strErrorMessagem, lngErrorNumber, cstr_ProcedureName, VBA.Erl)
320 End If
    GoTo Fim:
    'Debug Mode
330 Resume
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : Form_frmEditCadastroCorrespondente.verificaCamposAlterados()
' TIPO             : Function
' DATA/HORA        : 14/06/2015 20:37
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Analisa os dados digitados e e verifica se houve alteração em relação aos valores originais
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      14/06/2015 20:37    Adelson         Criação/Atualização do procedimento
'---------------------------------------------------------------------------------------
Function verificaCamposAlterados(f As Object, Optional ByRef cDadosOriginais As Collection) As Collection
    Dim ctr As Control
    Dim valor As Variant
    Dim cAlterados As Collection
10  On Error GoTo verificaCamposAlterados_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
20  Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Function Form_frmEditCadastroCorrespondente.verificaCamposAlterados()"
    '----------------------------------------------------------------------------------------------------
30  Set cAlterados = New Collection
    
    If cDadosOriginais Is Nothing Then
        Call MessageBoxMaster("É necessário reiniciar o cadastro para atualizar os dados Originais", VBA.vbExclamation)
        GoTo Fim
    End If
40  For Each ctr In f.Controls
        'If ctr.Name = "TelTrabalho" Then Stop
        If ctr.tag Like "*CampoMonitorado=True*" Then
50          valor = cDadosOriginais.item(ctr.Name)
60          If Nz(ctr.value, VBA.vbNullString) <> Nz(valor, VBA.vbNullString) Then
                Call cAlterados.Add("Campo : " & ctr.Name & " / Valor Original : " & Nz(valor, "(Vazio)") & " / Valor Alterado : " & Nz(ctr.value, "(Vazio)"), ctr.Name)
            End If
        End If
70  Next

Fim:
80  Set verificaCamposAlterados = cAlterados
90  On Error GoTo 0
100 Exit Function

verificaCamposAlterados_Error:
110 If VBA.Err <> 0 Then
120     lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
130     Call Excecoes.TratarErro(strErrorMessagem, lngErrorNumber, cstr_ProcedureName, VBA.Erl)
140 End If
    GoTo Fim:
    'Debug Mode
150 Resume
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : AuxFormCliente.ContarCamposPreenchidos()
' TIPO             : Function
' DATA/HORA        : 12/09/2015 00:30
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Conta quantos controles estão preenchidos em um formulário com campos de preenchimentos
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      12/09/2015 00:30    Adelson         Criação/Atualização do procedimento
'---------------------------------------------------------------------------------------
Function ContarCamposPreenchidos(f As Object) As Collection
    Dim ctr As Control
    Dim valor As Variant
    Dim cPreenchidos As Collection
10  On Error GoTo ContarCamposPreenchidos_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
20  Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Function AuxFormCliente.ContarCamposPreenchidos()"
    '----------------------------------------------------------------------------------------------------

30  Set cPreenchidos = New Collection

40  For Each ctr In f.Controls
50      If ctr.tag Like "*CampoMonitorado=True*" And (ctr.controlType = acComboBox Or ctr.controlType = acTextBox) Then
60          valor = Nz(ctr.value, VBA.vbNullString)
70          If valor <> "" Then Call cPreenchidos.Add("Campo : " & ctr.Name & " / Valor : " & valor)
80      End If
90  Next

Fim:
100 Set ContarCamposPreenchidos = cPreenchidos

110 On Error GoTo 0
120 Exit Function

ContarCamposPreenchidos_Error:
130 If VBA.Err <> 0 Then
140     lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
150     Call Excecoes.TratarErro(strErrorMessagem, lngErrorNumber, cstr_ProcedureName, VBA.Erl)
160 End If
    GoTo Fim:
    'Debug Mode
170 Resume
End Function

Sub ResetaCores(f As Object)
    Dim ctr As Control
40  For Each ctr In f.Controls
50      If ctr.tag Like "*CampoMonitorado=True*" And (ctr.controlType = acComboBox Or ctr.controlType = acTextBox) Then
70          ctr.BorderColor = 0
80      End If
90  Next
End Sub
'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : AuxFormCliente.OnClick()
' TIPO             : Function
' DATA/HORA        : 31/08/2014 13:40
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Executa uma ação personalizada generica para qualquer botão clicado
'                    em um formulário que seja chamado da macro OnClick
'                    Essa rotina ja contempla todas as ações necessárias para modificar o comportamento do formulario,
'                    colocar os icones de concluído
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      31/08/2014 13:40    ARMS            Criação
' * 1.00      09/10/2015 16:42    ARMS            Atualização e documentação
'---------------------------------------------------------------------------------------
Public Function OnClick(Optional objAcvtiveControl As Control, Optional objForm As Form)
'---------------------------------------------------------------------------------------
1   On Error GoTo OnClick_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
2   Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Function AuxFormCliente.OnClick()"
    '---------------------------------------------------------------------------------------
    Dim strObject As String
    Dim bResult As Boolean
    Dim vResult As Variant

    'Atualiza a instancia do Objeto e Formulário ativos
3   If Not objAcvtiveControl Is Nothing And Not objForm Is Nothing Then
4       Set ctvControl = objAcvtiveControl
5       Set frmCurrentForm = objForm
6   Else
7       Set frmCurrentForm = Nothing
8       Set ctvControl = Nothing
9       Call setActiveFormControl
10  End If

11  If ctvControl Is Nothing Then Exit Function
12  strObject = frmCurrentForm.Name & "." & ctvControl.Name

    'Prepara a tela
13  Call IniciaAçãoClick(frmCurrentForm)
14  bResult = True

    '************************************************************************************************************************
    'Manipula cada objeto individualmente
    'Incluir as chamadas de regra de negócio aqui
    'A variável bResult deve receber o retorno True ou False que será usado no tratamento das imagens
    '************************************************************************************************************************

15  If ctvControl.tag Like "ExportarParaExcel_TabelaConsultaOrigem:=*" Then
        Dim xExcel As New cTFW_Excel
        Call xExcel.ExportarParaPlanilha(pOrigem:=getTheValue(ctvControl.tag, "ExportarParaExcel_TabelaConsultaOrigem"), _
                                     pDestino:="Plan Teste", _
                                     pIncluirCabecalho:=True, _
                                     pOpcoes:=Array("NomeLista:=" & getTheValue(ctvControl.tag, "ConfigNomeLista"), "Layout=TableStyleLight1", "="), _
                                     pExibirArquivo:=True)
17      Call RemoverObjetosMemoria(xExcel)
18  Else
19      Select Case strObject
        Case frmCurrentForm.Name & ".btnAlterarPeriodo"
            Call CarregaDataReferencia(, "Informar data")
        
        Case frmCurrentForm.Name & ".btnProcessarAlgo"
            '**** INSTRUÇÕES *********************'**************************************************************************************************************************
            'Faça a chamada à uma função de processamento com retorno.
            'Tratar o retorno da função e salvar na variável 'bResult' com valor de True ou False para que os icones sejao exibidos adequadamente no formulário de processos
            '------------------------------------------------------------
            'bResult = Call Modulo.RotinaProcessamento()
            '------------------------------------------------------------
            '****************************************************************************************************************************************************************

            'Exibi uma consulta
20      Case frmCurrentForm.Name & ".btnExibirSubForm_consulta"
21          Call AtivarSubFormVisualizacao(frmCurrentForm, "DBReport_Tabelas")

            'Exibi um Sub Formulário
22      Case frmCurrentForm.Name & ".btnExibirSubForm_subformulario"
23          Call AtivarSubFormVisualizacao(frmCurrentForm, "sfSubFormularioTeste")

            'Exibi uma tabela
24      Case frmCurrentForm.Name & ".btnExibirSubForm_tabela"
25          Call AtivarSubFormVisualizacao(frmCurrentForm, "MSysObjects")

            'Exibi o que estiver na Tag do Botão
26      Case frmCurrentForm.Name & ".btnExibirSubForm_porTag"

27          Call AtivarSubFormVisualizacao(frmCurrentForm, ctvControl.tag)

28      Case frmCurrentForm.Name & ".btnExportarParaExcel"
            '**** INSTRUÇÕES *********************'**************************************************************************************************************************
            'Faça a chamada à uma função que exporte algo para o Excel
            'Da mesma forma, se quiser exibir a imagem "Tick" indicando que a ação foi executada. Tratar o retorno da função e salvar na variável 'bResult' com valor de True
            'ou False para que os icones sejao exibidos adequadamente no formulário de processos
            '------------------------------------------------------------
            'Call Modulo.ExportarAlgoParaExcel()
            '------------------------------------------------------------
            '****************************************************************************************************************************************************************

            'Demais ações pode ser incluidas no case
29      Case frmCurrentForm.Name & ".btnBotaoDeAcao1"

30      Case frmCurrentForm.Name & ".btnBotaoDeAcao2"

31      Case frmCurrentForm.Name & ".btnBotaoDeAcao3"

            'Caso deseja que a mesma ação seja executada para mais de um botão, basta chamar assim...
32      Case frmCurrentForm.Name & ".btnBotaoDeAcao4", _
             frmCurrentForm.Name & ".btnBotaoDeAcao5", _
             frmCurrentForm.Name & ".btnBotaoDeAcao6"

            'Dica : Lembrando que é possivel ler qualque informação do botão Clicado pela instancia do Controle na variável : ctvControl
            'ctvControl.Name
            'ctvControl.Tag
            'etc...
33      Case Else
            Debug.Print "OnLick() " & strObject & " - Ação não configurada !"
34          VBA.MsgBox strObject & " - Ação não configurada !", VBA.vbExclamation
35      End Select
36  End If

Fim:
37  Call FinalizaAcaoClick(frmCurrentForm, CInt(bResult), strObject)
38  On Error Resume Next

39  Resume
40  Call frmCurrentForm.Reload

41  On Error GoTo 0
42  Exit Function

OnClick_Error:
43  If VBA.Err <> 0 Then
44      lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
45      Call Excecoes.TratarErro(strErrorMessagem, lngErrorNumber, cstr_ProcedureName, VBA.Erl())
46  End If
    GoTo Fim:
    'Debug Mode
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : AuxForm.setActiveFormControl()
' TIPO             : Function
' DATA/HORA        : 28/08/2014 09:39
' CONSULTOR        : (ADELSON)/ TECNUN - Adelson Rosendo Marques da Silva
' CONTATO          : adelson@tecnun.com.br
' DESCRIÇÃO        : Define a instancia atual do formulário e controle ativo
'---------------------------------------------------------------------------------------
Public Function setActiveFormControl() As Variant
10  On Error Resume Next
20  Set frmCurrentForm = Screen.ActiveForm
    If frmCurrentForm.Name = "frmConfiguracoes" Then Set frmCurrentForm = frmCurrentForm.sfContainer.Form
30  Set ctvControl = frmCurrentForm.ActiveControl
End Function

Sub Atributos()
    Dim dbsNorthwind As Object
    Set dbsNorthwind = CurrentDb
    With dbsNorthwind
        Dim f As Object 'dao.Field
        Dim p As Object 'dao.Property
        Dim tdfloop As Object 'dao.TableDef
        Dim sPath As String
        sPath = CurrentProject.Path & "\Log\Tables"
        Call MkFullDirectory(sPath)

        For Each tdfloop In .TableDefs
            If Not tdfloop.Name Like "MSys*" Then
                Call RegistraLog("Table : " & tdfloop.Name & " / Atributo : " & tdfloop.Attributes & String(50, "-"), sPath & "\Table_" & tdfloop.Name & ".txt", True)
                
                Call RegistraLog(String(60, "-"), sPath & "\Table_" & tdfloop.Name & ".txt")
                Call RegistraLog("### Fields ### (" & tdfloop.Fields.count & ") Total", sPath & "\Table_" & tdfloop.Name & ".txt")
                Call RegistraLog(String(60, "-"), sPath & "\Table_" & tdfloop.Name & ".txt")
                
                For Each f In tdfloop.Fields
                    Call RegistraLog(Space(2) & f.Name, sPath & "\Table_" & tdfloop.Name & ".txt")
                Next f
                
                Call RegistraLog(VBA.vbNewLine, sPath & "\Table_" & tdfloop.Name & ".txt")
                Call RegistraLog(String(60, "-"), sPath & "\Table_" & tdfloop.Name & ".txt")
                Call RegistraLog("### Properties ### (" & tdfloop.Properties.count & ") Total", sPath & "\Table_" & tdfloop.Name & ".txt")
                Call RegistraLog(String(60, "-"), sPath & "\Table_" & tdfloop.Name & ".txt")
                
                For Each p In tdfloop.Properties
                    Call RegistraLog(Space(2) & p.Name & " = " & p.value, sPath & "\Table_" & tdfloop.Name & ".txt")
                Next
            End If
        Next tdfloop
        OpenFileInNotepad CurrentProject.Path & "\Log\Table_" & tdfloop.Name & ".txt"
        .Close
    End With
End Sub

Sub AtivarSubFormVisualizacao(frmProcesso As Object, sFonteDados As String)
    Set frmProcesso.ActivePage = frmProcesso.mpProcesso.Pages(frmProcesso.mpProcesso.value)
    frmCurrentForm.parent.Form.Fechar.Caption = "&Inicio"
    'frmProcesso.Fechar.Caption = "&Inicio"
    
250 Call AuxForm.ExibirSubForm(frmProcesso.sfContainer, sFonteDados)
    'Esse metodo irá ativar a guia de visualização de dados que deverá ser a ultima no Controle Guia
260 frmProcesso.mpProcesso.value = frmProcesso.mpProcesso.Pages(frmProcesso.mpProcesso.Pages.count - 1).PageIndex
End Sub

Function FormsExists(pFormName As String) As Boolean
    Dim bRet As String
    On Error Resume Next
    bRet = Access.Application.CurrentProject.AllForms(pFormName).Name
    FormsExists = bRet <> ""
End Function



'--------------------------------------------------------------------------------------------------------------------------------------------------------
' FUNÇÕES ADICIONAIS QUE MANIPULAM AS IMAGEMS QUE REPRESENTAM OS PROCESSO
'--------------------------------------------------------------------------------------------------------------------------------------------------------
'Variáveis publicas necessárias
'----------------------------------------------------
'Public ctvControl As control
'Public frmCurrentForm As Form
'----------------------------------------------------
Public Sub IniciaAçãoClick(objActiveForm As Form)
    On Error Resume Next
    DoCmd.Echo False
10  If objActiveForm Is Nothing Then Set objActiveForm = Screen.ActiveForm
20  Call AuxForm.DefiniPropriedadeControles(objActiveForm, "Enabled", False, "CommandButton", "*")
30  Call AuxForm.DefiniPropriedadeControles(objActiveForm.Detalhe, "Visible", False, "Image", "img")
'40  Call AuxForm.ExibirIcone_Executando(objActiveForm, ctvControl.Name)
    DoCmd.Echo True
End Sub
'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : AuxForm.FinalizaAcaoClick()
' TIPO             : Sub
' DATA/HORA        : 28/08/2014 09:40
' CONSULTOR        : (ADELSON)/ TECNUN - Adelson Rosendo Marques da Silva
' CONTATO          : adelson@tecnun.com.br
' DESCRIÇÃO        : Finaliza uma ação configurando os icones de um formulário de processo
'---------------------------------------------------------------------------------------
Public Sub FinalizaAcaoClick(objActiveForm As Form, Optional bUpdateStatus As Boolean, Optional strObjStatus As String, Optional pConexto As String)
10  On Error GoTo FinalizaAcaoClick_Error
20  If strObjStatus = "" Then strObjStatus = objActiveForm.Name & "." & ctvControl.Name
40  Call AuxForm.DefiniPropriedadeControles(objActiveForm, "Enabled", True, "CommandButton", "*")
50  Call AuxForm.LimparProgressoEmForm(objActiveForm)
    'Aplicavel a outro processo
    Call SalvarStatusProcesso(strObjStatus, CInt(bUpdateStatus), pConexto)
    Call CarregarImagensStatus(objActiveForm.mpProcesso, True, objActiveForm)
70  On Error GoTo 0
80  Exit Sub
FinalizaAcaoClick_Error:
90  If VBA.Err <> 0 Then Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxForm.FinalizaAcaoClick()", VBA.Erl, , False)
100 Exit Sub
110 Resume
End Sub

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : AuxForm.CarregarImagensStatus()
' TIPO             : Sub
' DATA/HORA        : 04/05/2014 10:22
' CONSULTOR        : (ADELSON)/ TECNUN - Adelson Rosendo Marques da Silva
' CONTATO          : adelson@tecnun.com.br
' DESCRIÇÃO        : Visualiza as imagems Ticks dos status de acordo com o botão
'---------------------------------------------------------------------------------------
Public Sub CarregarImagensStatus(objParent As Object, _
                                 Optional filterType As String = "*", _
                                 Optional frmParent As Form, Optional contexto As String)

    Dim ctr As Control
    Dim intStatus As Integer
    Dim img As Object
    Dim vResultStatus As Variant

    '10  On Error Resume Next
10  On Error GoTo CarregarImagensStatus_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
20  Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Sub CarregarImagensStatus()"
    '----------------------------------------------------------------------------------------------------

30  Set currentParentControls = frmParent

40  For Each ctr In currentParentControls.Controls
50      VBA.DoEvents
60      If VBA.TypeName(ctr) Like "*" & filterType & "*" Or filterType = "*" Then
70          vResultStatus = PegarStatusContextoProcessos(frmParent.Name & "." & ctr.Name, contexto)
80          On Error Resume Next
90          Set img = objParent.Controls("img_" & ctr.Name)
100         If Not img Is Nothing Then
110             img.Visible = vResultStatus(0) = -1 And ctr.Enabled
120         End If
130         Set img = Nothing
140         On Error GoTo 0
150     End If
160 Next

Fim:
170 On Error GoTo 0
180 Exit Sub

CarregarImagensStatus_Error:
190 If VBA.Err <> 0 Then
200     lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
210     Call Excecoes.TratarErro(strErrorMessagem, lngErrorNumber, cstr_ProcedureName, VBA.Erl, , False)
220 End If
    GoTo Fim:
    'Debug Mode
230 Resume
End Sub
'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : mdlProcesso.PegarStatusContextoProcessos()
' TIPO             : Function
' DATA/HORA        : 04/03/2014 13:51
' CONSULTOR        : (ADELSON)/ TECNUN - Adelson Rosendo Marques da Silva
' CONTATO          : adelson@tecnun.com.br
' DESCRIÇÃO        : Recupera informações sobre status dos processos
'---------------------------------------------------------------------------------------
'                   ....................................................................................
'                    Consultas necessárias no CLIENTE
'                   ....................................................................................
'                     > Pegar_StatusEtapaProcesso (@cenario, @contexto)
'                        SELECT tb.cenario, tb.contexto, tb.Status, tb.data_hora, tb.usuario, tb.info_adicional
'                        FROM tblStatusProcesso AS tb
'                        WHERE tb.cenario=[@cenario] AND (tb.contexto=[@contexto] OR @cenario] Is Null);

Public Function PegarStatusContextoProcessos(contexto As String, _
                                      Optional strCenario As String) As Variant
    Dim rsTableStatus As Object
    Dim sSelect As String
    Dim pWhere As String
    Dim sVisao As String
10  On Error GoTo PegarStatusContextoProcessos_Error

    Dim vStatus As Variant
    If Conexao Is Nothing Then Call Inicializar_Globais(False)
    
    'Atualiza a conexão com o cliente, pois o PegarArray fecha a conexão
    Set Conexao.ConexaoBanco = CurrentProject.Connection
20  vStatus = Conexao.PegarArray("Pegar_StatusEtapaProcesso", strCenario, contexto)
    If Not VBA.IsEmpty(vStatus) Then
        PegarStatusContextoProcessos = Array(vStatus(2, 0), vStatus(1, 0))
    Else
        PegarStatusContextoProcessos = Array(0, contexto)
    End If
160 On Error GoTo 0
170 Exit Function

PegarStatusContextoProcessos_Error:
180 If VBA.Err <> 0 Then Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "mdlProcesso.PegarStatusContextoProcessos()", VBA.Erl)
190 Exit Function
200 Resume
End Function
'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : mdlProcesso.SalvarStatusProcesso()
' TIPO             : Sub
' DATA/HORA        : 04/03/2014 13:51
' CONSULTOR        : (ADELSON)/ TECNUN - Adelson Rosendo Marques da Silva
' CONTATO          : adelson@tecnun.com.br
' DESCRIÇÃO        : Salva os informações sobre status do processo
'                   ....................................................................................
'                    Consultas necessárias no Cliente
'                   ....................................................................................
'                     > Limpar_StatusEtapaProcesso (@cenario, @contexto)
'                     > Insere_StatusEtapaProcesso (@cenario, @contexto, @status, @usuario)
'---------------------------------------------------------------------------------------
Public Sub SalvarStatusProcesso(contexto As String, pValorStaus As String, strCenario As String, Optional infoAdicionais As Variant)
    Dim rsTableStatus As Object
    Dim sSelect As String
    Dim pWhere As String

10  On Error GoTo SalvarStatusProcesso_Error
20  If VBA.IsNumeric(pValorStaus) Then pValorStaus = CInt(pValorStaus)
    If Conexao Is Nothing Then Inicializar_Globais
30  Call Conexao.InserirRegistros("Limpar_StatusEtapaProcesso", strCenario, contexto)
40  Call Conexao.InserirRegistros("Insere_StatusEtapaProcesso", strCenario, contexto, pValorStaus, ChaveUsuario)
50  On Error GoTo 0
60  Exit Sub

SalvarStatusProcesso_Error:
70  If VBA.Err <> 0 Then Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "mdlProcesso.SalvarStatusProcesso()", VBA.Erl)
80  Exit Sub
90  Resume
End Sub

Sub RefreshImageState(objContainer As Object, Optional bClearState As Boolean = True, Optional frmParent As Object)
10  If (frmParent Is Nothing) Or (objContainer Is Nothing) Then Exit Sub
20  If bClearState Then Call AuxForm.DefiniPropriedadeControles(objContainer, "Visible", False, "Image", "img")
30  Call AuxForm.CarregarImagensStatus(objContainer, "CommandButton", frmParent)
End Sub

Sub RequeryForms()
10  On Error GoTo TratarErro
    Dim frm    As AccessObject
20  For Each frm In Access.CurrentProject.AllForms
30      If frm.IsLoaded Then Access.Application.Forms(frm.Name).Refresh
40  Next
50  Exit Sub
TratarErro:
60  Exit Sub
End Sub

Sub LimparCombo(objCombo As Object)
    Do While objCombo.ListCount <> 0
        objCombo.RemoveItem objCombo.ListCount - 1
    Loop
    objCombo = ""
End Sub

Sub AbrirSubFormulario(objFonte As String, Optional Legenda As String, Optional fParent As String = "frmConfiguracoes", Optional ModoJanelaMaximizada As Boolean = False, Optional bReabrirForm As Boolean = True)
    If fParent = "frmPrincipal" Then
        Form_frmPrincipal.Visible = True
        Call salvaValor("TrocarForm", "-1")
        Call salvaValor("subform", objFonte)
        Call salvaValor("CaptionForm", Legenda)
        Call AbrirFormulario(fParent, acNormal, AcWindowMode.acWindowNormal)
        Form_frmPrincipal.Reload
    Else
        Call salvaValor("CaptionForm", Legenda)
        Call salvaValor("subform", objFonte)
        Call salvaValor("ReabrirForm", CInt(bReabrirForm))
        Call AbrirFormulario(fParent, acNormal, VBA.IIf(ModoJanelaMaximizada, AcWindowMode.acWindowNormal, AcWindowMode.acDialog))
    End If
    If fParent = "frmPrincipal" Then Call salvaValor("TrocarForm", "")
End Sub
