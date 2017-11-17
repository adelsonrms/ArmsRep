Attribute VB_Name = "AuxRibbon"
'---------------------------------------------------------------------------------------
' PROJETO/MODULO   : TFWCliente.AuxRibbon
' TIPO             : Module
' DATA/HORA        : 19/01/2015 14:49
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Contem assinatura de todos os callbacks disponiveis para cada elemento
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão do Módulo
' **************************************************************************************
'   Versão    Data/Hora             Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      12/07/2016 23:47
'---------------------------------------------------------------------------------------
Option Explicit

Dim strXML      As String
Dim Guia_Ribbon As Boolean
Dim Quant_Objetos As Integer
Dim bValidaVisible As Boolean

Public Ribbon_Tela As String

Public objRibbonInstance As Object ' IRibbonUI

Public Sub OnRibbonLoadMain(ribbon As Object)
10  Set objRibbonInstance = ribbon
End Sub

'Callback generico iniciado por todos os botões da Ribbon
Public Sub OnActionButton(Control As Object) ' Object)
    'Direciona para a rotina generica que recebe o ID e a Tag do botão clicado
    Call ExecutaClickBotaoRibbon(Control.ID, Control.tag)
End Sub
'--------------------------------------------------------------------
' PROCEDIMENTO     : Ribbon.ExecutaClickBotaoRibbon()
' TIPO             : Sub
' DATA/HORA        : 17/11/2014 17:51
' CONSULTOR        : (ADELSON)/ TECNUN - Adelson Rosendo Marques da Silva
' CONTATO          : adelson@tecnun.com.br
' DESCRIÇÃO        : Executa o clique dos botões customizável da Ribbon com funções pre-definidas
'                    De acordo com a configuração na Tag do botão, executa ações especificas e dinamicas
'                    1 - Nome de Formulário : Abre o formulário atraves da função AbrirFormulario()
'                    2 - Parametro 'RunSubRotina:=NomeDaRotina', singifica que deverá ser executado uma rotina especificada (NomeDaRotina).
'                        Dessa forma não precisamos criar um Case especifico para exectar essa função.
'                    3 - Parametro 'FormName:=' : Indica que queremos abrir um formulário indicando o nome (mesma logica que a opção 1, porem aqui podemos incluir mais parametros)
'                    4 - Parametros 'FormPrincial:=NomeFormPrincial;SubForm_Origem:=NomeOrigemSubForm;FormLegenda:=Legenda do form'
'                        Essa configuração permite abrirmos qualquer formulário (como subformario) dentro de um formulário princial ja padronizado. Ex. O formulario de Opçoes
'
'                    Caso nenhuma configuração for informana na Tag, será analisado cada botão especifico atraves de um SELECT CASE idBotao
'---------------------------------------------------------------------------------------

Private Sub ExecutaClickBotaoRibbon(pID As String, Optional pTag As String)
1   On Error GoTo OnAction_DefaultForm_Error
2   On Error Resume Next
4   On Error GoTo OnAction_DefaultForm_Error
    Dim FormNameFromTag As String

6   If pTag <> "" Then
        'Busca informações da Tag atualizadas (Caso atualize nao tenha reiniciado
        pTag = BuscaInfoTag(pID)
        
        'Abre o formulário indicado na tag
7       If FormsExists(pTag) Then
8           Call AbrirFormulario(pTag)
            'Caso seja especificado o nome de uma rotina a qual deseja ser executada.
            'A tag tem que ter somente um item
9       ElseIf pTag Like "RunSubRotina:=*" Then
10          Call Access.Application.Run(getTheValue(pTag, "RunSubRotina"))

11      Else
12          If pTag Like "rpt*" Then
13              Call salvaValor("RPT_NAME", pTag)
14              Call AbrirFormulario("frm_ReportHorasMes")
15          Else
16              FormNameFromTag = getTheValue(pTag, "FormName")
17              If FormsExists(FormNameFromTag) Then
18                  Call AbrirFormulario(FormNameFromTag)
19              Else
20                  If getTheValue(pTag, "FormPrincial") <> "" Then
21                      Call AbrirSubFormulario(getTheValue(pTag, "SubForm_Origem"), _
                                                getTheValue(pTag, "FormLegenda"), _
                                                getTheValue(pTag, "FormPrincial"), _
                                                getTheValue(pTag, "WindowNormal") = "Sim")
22                  End If
23              End If
24          End If
25      End If
26  Else
27      Select Case pID
        Case "btnAbout"
28          Call AbrirFormulario("frmSobre")

29      Case "btnMensagens"
30          Call AbrirFormulario("frmMensagens")
            '------------------------------------------------------------------------
            'Funções do Aplicativo
            '------------------------------------------------------------------------
31      Case "btnNovaAplicacao"
32          Call AbrirFormulario("frmNovaApp")

33      Case "btnAplicacao"
34          Call AbrirFormulario("frmAppConfig", , acDialog, Form_frmPrincipal)

35      Case "btnRibbonEditor"
36          Call AbrirRibbonEditor

        Case "btnConfiguracao"
            Call AbrirFormulario("frmConfiguracoes")

'37      Case "btnCompactarBE_user"
'38          If AuxMensagens.MessageBoxMaster("F017") = VBA.vbYes Then
'39              Call AuxTabela.Compactar_BackEnd
'40          End If
41      Case Else
42          Debug.Print pID
43          Call VBA.MsgBox(pID & " - Botão ou comando sem ação Atribuida", VBA.vbExclamation, "Ribbon")
44      End Select
45  End If

Fim:
46  On Error GoTo 0
47  Exit Sub

48  On Error GoTo 0
49  Exit Sub
OnAction_DefaultForm_Error:
50  If VBA.Err = 2517 Then
51      VBA.MsgBox "FUNÇÃO : " & getTheValue(pTag, "RunSubRotina") & VBA.vbNewLine & VBA.vbNewLine & "A função configurada para ser executada no botão ainda não foi implementada !", VBA.vbExclamation, "RIBBON - Função Necessário"
52  Else
53      If VBA.Err <> 0 Then Call TratarErro(VBA.Err.Description, VBA.Err.Number, "Ribbon.OnAction_DefaultForm()", VBA.Erl)
54  End If
55  Exit Sub
56  Resume
End Sub

Private Function BuscaInfoTag(pID As String)
    With CurrentDb.OpenRecordset("SELECT tag FROM tblRibbon_Controls WHERE id = '" & pID & "'")
        If Not .EOF Then BuscaInfoTag = Nz(!tag.value, "")
    End With
End Function


Sub AbrirRibbonEditor()
    Call AbrirSubFormulario("frmRibbonEditor", "Configurações da Ribbon")
End Sub

'GetEnabled
Public Sub GetEnabled(Control As Object, ByRef Enabled)
10  Select Case Control.ID
        Case Else
20          Enabled = True
30  End Select
    If Control.tag <> "" Then
        If getTheValue(Control.tag, "Enabled") = "0" Then Enabled = False
    End If
End Sub

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : Ribbon.resetRibbon()
' TIPO             : Sub
' DATA/HORA        : 17/11/2014 17:51
' CONSULTOR        : (ADELSON)/ TECNUN - Adelson Rosendo Marques da Silva
' CONTATO          : adelson@tecnun.com.br
' DESCRIÇÃO        : Reinicia a instancia da Ribbon
'---------------------------------------------------------------------------------------
Sub resetRibbon(Optional oRibbon As Object, Optional pControlID As String)
10  On Error GoTo resetRibbon_Error
20  If oRibbon Is Nothing Then
30      Set oRibbon = objRibbonInstance
40  End If

50  If oRibbon Is Nothing Then
60      VBA.MsgBox "Não é possivel redefinir a Faixa de Opções (Ribbon). É necessário reiniciar o arquivo.", VBA.vbExclamation, "Instancia Inválida"
70  Else
80      If pControlID <> "" Then
90          Call oRibbon.InvalidateControl(pControlID)
100     Else
110         oRibbon.Invalidate
120     End If
130     If VBA.Err = 0 Then VBA.MsgBox "A Faixa de Comandos foi redefinida com sucesso !", VBA.vbInformation
140 End If
150 On Error GoTo 0
160 Exit Sub
resetRibbon_Error:
170 If VBA.Err <> 0 Then Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "Ribbon.resetRibbon()", VBA.Erl, , False)
180 Exit Sub
190 Resume
End Sub

' *************************************************************
' Created from     : Avenius
' Parameter        : Input String, SuchValue String
' Date created     : 05.01.2008
'
' Sample:
' getTheValue("DefaultValue:=Test;Enabled:=0;Visible:=1", "DefaultValue")
' Return           : "Test"
' *************************************************************
Public Function getTheValue(strTag As String, strValue As String) As String
10  On Error Resume Next

    Dim workTb() As String
    Dim Ele()  As String
    Dim myVariabs() As String
    Dim i      As Integer

20  workTb = VBA.Split(strTag, ";")

30  ReDim myVariabs(LBound(workTb) To UBound(workTb), 0 To 1)
40  For i = LBound(workTb) To UBound(workTb)
50      Ele = VBA.Split(workTb(i), ":=")
60      myVariabs(i, 0) = Ele(0)
70      If UBound(Ele) = 1 Then
80          myVariabs(i, 1) = Ele(1)
90      End If
100 Next

110 For i = LBound(myVariabs) To UBound(myVariabs)
120     If strValue = myVariabs(i, 0) Then
130         getTheValue = myVariabs(i, 1)
140     End If
150 Next

End Function

Public Function getAppPath() As String
    Dim strDummy As String
10  strDummy = CurrentProject.Path
20  If Right(strDummy, 1) <> "\" Then strDummy = strDummy & "\"
30  getAppPath = strDummy
End Function

Sub setEnabled(controlID As String, ByRef Enabled)
    Select Case controlID
        Case ""
            Enabled = False
        Case Else
            Enabled = True
    End Select
End Sub

Function RibbonObjectExists(pID As String) As Boolean
    RibbonObjectExists = Not VBA.IsNull(Nz(DLookup("ID", "tblRibbon_Controls", "ID='" & pID & "'"), Null))
End Function

Function getVisibleObject(pID As String) As Boolean
    getVisibleObject = Nz(DLookup("visible", "tblRibbon_Controls", "ID='" & pID & "'"), False)
End Function

Sub setVisible(controlID As String, ByRef Visible)
    'Default é Visivel
    Visible = True
    If RibbonObjectExists(controlID) Then Visible = getVisibleObject(controlID)
    'Oculta os controles de acordo com condições espeficicas
    Select Case controlID
        Case "tabDev", "tabFunc"    ', "TabHomeAccess", "TabCreate"
            'O grupo 'Tools - Desenv' deve ficar visivel somente para os desenvolvedores.
            'Valida o acesso pelo nome do usuáro ou pelo nome da maquina
            'Adicionar mais nomes aqui caso haja novos desenvolvedores
            Visible = InStr("JEFFERSON|ADELSON|JEFFERSON_PC|ADELSON-PC", UCase(Environ("ComputerName"))) > 0
    End Select
End Sub

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : Ribbon.GetVisible()
' TIPO             : Sub
' DATA/HORA        : 05/03/2014 14:40
' CONSULTOR        : (ADELSON)/ TECNUN - Adelson Rosendo Marques da Silva
' CONTATO          : adelson@tecnun.com.br
' DESCRIÇÃO        : Define a visibilidade dos controles da Ribbon
'                    O argumento 'visible' define o valor True ou False para tornar
'                    visivel ou nao o elemento. É um argumento de saida
'---------------------------------------------------------------------------------------
Sub getVisible(Control As Object, ByRef Visible)
    'Default é Visivel
    Visible = True
    'Verifica se contem a marca que será validada a versão
    If Control.tag Like "*CheckIsDevMode=True*" Then
        'Verifica se a versão atual é esta em Homologação/Desenvolvimento
        Visible = DevMode()
    ElseIf Control.tag Like "*CheckProfileAccess=True*" Then
        
    End If
End Sub

Sub gerarRibbonXML(strRibbonName As String)
    Dim sPathToSave As String
    Dim intFile As Integer
10  strXML = ""
20  Call loadRibbon(strRibbonName)

    Dim dlg As Object
    
40  Set dlg = Access.Application.FileDialog(msoFileDialogSaveAs)
50  With dlg
60      .Title = ""
70      .AllowMultiSelect = False
        .InitialFileName = strRibbonName & ".xml"
        .show
80      sPathToSave = .SelectedItems(1)
90  End With

100 If VBA.Dir(sPathToSave) <> "" Then Call VBA.Kill(sPathToSave)
110 intFile = VBA.FreeFile()
120 Open sPathToSave For Output As #intFile: Print #intFile, strXML: Close #intFile

    If VBA.MsgBox("Deseja visualizar o arquivo gerado ?", VBA.vbQuestion + VBA.vbYesNo, "Ribbon XML") = VBA.vbNo Then Exit Sub
    Call AuxFileSystem.OpenFileInNotepad(sPathToSave)
End Sub
'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : mdlRibbon.InitializeRibbon()
' TIPO             : Function
' DATA/HORA        : 19/01/2015 13:53
' CONSULTOR        : (ADELSON)/ TECNUN - Adelson Rosendo Marques da Silva
' CONTATO          : adelson@tecnun.com.br
' DESCRIÇÃO        : Inicializa a ribbon da aplicação
'---------------------------------------------------------------------------------------
Public Function InitializeRibbon()
    'carrega as faixas de opção para o usuário
    Dim rsRibbon As Object
    Dim intFile As Integer

    '----------------------------------------------------------------------------------------------------
10  On Error GoTo InitializeRibbon_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
20  Dim dtSartRunProc As Date: dtSartRunProc = Time
    Dim varivaeis As Variant
    Const cstr_ProcedureName As String = "Function RibbonTable.mdlRibbon.InitializeRibbon()"
    '----------------------------------------------------------------------------------------------------

40  Set rsRibbon = CurrentDb.OpenRecordset("tblRibbon") ' " &VBA.vba.iif(ribbonName <> "", " WHERE nome = '" & ribbonName & "'", "") & " ORDER BY Ordem")

50  If rsRibbon.EOF = False Then
60      rsRibbon.MoveFirst
70      Do
80          strXML = ""
90          Call loadRibbon(rsRibbon!nome.value)
            If rsRibbon!carregarXMLAccess.value = -1 Then
100             Call Access.Application.LoadCustomUI(rsRibbon!nome.value, strXML)
            End If
110         If rsRibbon!salvarXML.value = True Then
120             If VBA.Dir(VBA.Environ("temp") & "\" & CurrentProject.Name & "_" & rsRibbon!nome.value & ".xml") <> "" Then Call VBA.Kill(Environ("temp") & "\" & CurrentProject.Name & "_" & rsRibbon!nome.value & ".xml")
130             intFile = VBA.FreeFile()
140             Open VBA.Environ("temp") & "\" & CurrentProject.Name & "_" & rsRibbon!nome.value & ".xml" For Output As #intFile: Print #intFile, strXML: Close #intFile
150         End If
160         rsRibbon.MoveNext
170     Loop Until rsRibbon.EOF = True
180 End If

190 rsRibbon.Close

200 On Error GoTo 0
210 Exit Function

InitializeRibbon_Error:
220 If VBA.Err <> 0 Then
230     lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
240     varivaeis = Array(vbNullString)
        RegistraLog "Ribbon ja esta carregada..."
260 End If
270 Exit Function
    'Debug Mode
280 Resume
End Function

Private Function loadRibbon(ribbon As String) As String
    Dim rs     As Object
    Dim strVersao As String
20  Set rs = CurrentDb.OpenRecordset("select * from tblribbon where nome='" & ribbon & "'")
    
    strVersao = pegarVariavelAplicacao("appVersion")
    
     If strVersao <> "" Then
        strVersao = VBA.Split(VBA.Split(strVersao, "-")(0), ".")(2)
     End If
    
30  If rs.EOF = False Then
40      strXML = "<customUI xmlns=""http://schemas.microsoft.com/office/2006/01/customui"" onLoad=""" & rs!OnRibbonLoad.value & """>" & VBA.vbCrLf
        'strXML = strXML & VBA.Chr(9) & "<ribbon " & getRibbonAtributo(rs, "startFromScratch") & ">" & VBA.vbCrLf
50      If rs.Fields("startFromScratch").value = -1 Or strVersao = 2 Then
60          strXML = strXML & VBA.Chr(9) & "<ribbon startFromScratch=""true"">" & VBA.vbCrLf
70      Else
80          strXML = strXML & VBA.Chr(9) & "<ribbon startFromScratch=""false"">" & VBA.vbCrLf
90      End If
        '---------------------------------------------------------------------------
100     Call loadTabs(rs!nome.value)
        '---------------------------------------------------------------------------
        strXML = strXML & VBA.Chr(9) & "</ribbon>" & VBA.vbCrLf & "</customUI>"
120 End If
130 rs.Close
End Function

Private Function getRibbonAtributo(rs As Object, strTag) As String
    Dim sXML As String
    Dim pValue As Variant
    pValue = rs.Fields(strTag).value
    getRibbonAtributo = strXML = " " & strTag & "=""" & pValue & """"
    'If rsItems.Fields("tag").Value <> "" Or VBA.IsNull(rsItems.Fields("tag").Value) = False Then strXML = strXML & " tag=""" & rsItems.Fields("tag").Value & """"
End Function
'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : AuxRibbonXMLDesinger.loadTabs()
' TIPO             : Function
' DATA/HORA        : 04/02/2015 10:26
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Monta o XML com as tabs configuradas e seus grupos
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      04/02/2015 10:26    Adelson         Criação/Atualização do procedimento
'---------------------------------------------------------------------------------------
Private Function loadTabs(ribbon As String)
    Dim SQL    As String
    Dim rs As Object, rsGrupos As Object

10  On Error GoTo loadTabs_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
20  Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Function AuxRibbonXMLDesinger.loadTabs()"
    '----------------------------------------------------------------------------------------------------

30  Guia_Ribbon = False
40  SQL = "select *, 1 AS valor_bloqueio from tblribbon_Tabs where parent='" & ribbon & "' order by [order] "

50  Set rs = CurrentDb.OpenRecordset(SQL)

60  If rs.EOF = False Then
70      rs.MoveFirst
80      Do
90          If rs.Fields("OfficeMenu").value = -1 Then
100             strXML = strXML & VBA.Chr(9) & VBA.Chr(9) & "<officeMenu>" & VBA.vbCrLf
                'Localiza o grupo para carregar o menu
110             SQL = "SELECT tblRibbon_Groups.Nome FROM tblRibbon_Groups as grp WHERE (((grp.parent)='" & rs!ID.value & "'));"

120             Set rsGrupos = CurrentDb.OpenRecordset(SQL)

130             If rsGrupos.EOF = False Then
140                 loadElements (rsGrupos!ID.value)
150             End If
160             rsGrupos.Close
170             strXML = strXML & VBA.Chr(9) & VBA.Chr(9) & "</officeMenu>" & VBA.vbCrLf
180         Else

190             If Guia_Ribbon = False Then
200                 strXML = strXML & VBA.Chr(9) & VBA.Chr(9) & "<tabs>" & VBA.vbCrLf
210                 Guia_Ribbon = True
220             End If

230             If rs!mso.value = -1 Then
240                 strXML = strXML & VBA.Chr(9) & VBA.Chr(9) & "<tab idMso"
250             Else
260                 strXML = strXML & VBA.Chr(9) & VBA.Chr(9) & "<tab id"
270             End If

280             strXML = strXML & "=""" & rs!ID.value & """"

290             If VBA.IsNull(rs!label.value) = False Or rs!label.value <> "" Then strXML = strXML & " label=""" & rs!label.value & """"
300             If VBA.IsNull(rs!tag.value) = False Or rs!tag.value <> "" Then strXML = strXML & " tag=""" & rs!tag.value & """"
                If VBA.IsNull(rs!insertBeforeMso.value) = False Or rs!insertBeforeMso.value <> "" Then strXML = strXML & " insertBeforeMso=""" & rs!insertBeforeMso.value & """"
                '--------------------------------------------------------------------------------------------------------------------------
                '### Verifica aqui se a tab tem definição de permissão configurada
                'Se recuperar um 0, nega a visibilidade (oculta a tab)
                '--------------------------------------------------------------------------------------------------------------------------
310             If rs.Fields("getVisible").value <> "" Then
320                 strXML = strXML & " getVisible=""" & rs.Fields("getVisible").value & """"
330             Else
340                 If Nz(rs.Fields("valor_bloqueio").value, 1) = 0 Then
350                     strXML = strXML & " visible=""false""" & VBA.vbCrLf & VBA.vbCrLf
360                 Else
                        'Se é visível ou não
370                     If rs.Fields("visible").value = -1 Then
380                         strXML = strXML & " visible=""true""" & VBA.vbCrLf & VBA.vbCrLf
390                     Else
400                         strXML = strXML & " visible=""false""" & VBA.vbCrLf & VBA.vbCrLf
410                     End If

420                 End If
430             End If
440             strXML = strXML & " >"
                '---------------------------------------------------------------------------
450             Call loadGroup(rs!ID.value)
                '---------------------------------------------------------------------------
460             strXML = strXML & VBA.Chr(9) & VBA.Chr(9) & "</tab>" & VBA.vbCrLf & VBA.vbCrLf
470         End If
480         rs.MoveNext
490     Loop Until rs.EOF = True
500 End If
510 rs.Close
520 strXML = strXML & VBA.Chr(9) & "</tabs>"
Fim:
530 On Error GoTo 0
540 Exit Function

loadTabs_Error:
550 If VBA.Err <> 0 Then
560     lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
570     Debug.Print "Erro na criação das tabs : " & lngErrorNumber & " - " & strErrorMessagem & "(Line " & VBA.Erl() & " in " & cstr_ProcedureName & ")"
580 End If
    GoTo Fim:
590 Resume
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : mdlRibbon.loadGroup()
' TIPO             : Function
' DATA/HORA        : 19/01/2015 14:21
' CONSULTOR        : (ADELSON)/ TECNUN - Adelson Rosendo Marques da Silva
' CONTATO          : adelson@tecnun.com.br
' DESCRIÇÃO        : Cria os elementos de grupos associados a uma tab
'---------------------------------------------------------------------------------------
Private Function loadGroup(Tabs As String)
    Dim rsItems As Object
    Dim SQL As String
    

1   On Error GoTo loadGroup_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
2   Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Function AuxRibbonXMLDesinger.loadGroup()"
    '----------------------------------------------------------------------------------------------------

3   SQL = "select * from tblRibbon_Groups as grp where parent='" & Tabs & "' Order by [order]"

4   Set rsItems = CurrentDb.OpenRecordset(SQL)

5   If rsItems.EOF = False Then
6       rsItems.MoveFirst
7       Do
            'Nome do Grupo
8           If rsItems.Fields("Mso").value = -1 Then
9               strXML = strXML & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & "<group idMso"
10          Else
11              strXML = strXML & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & "<group id"
12          End If
13          strXML = strXML & "=""" & rsItems.Fields("id").value & """ "

14          If rsItems.Fields("tag").value <> "" Or VBA.IsNull(rsItems.Fields("tag").value) = False Then strXML = strXML & " tag=""" & rsItems.Fields("tag").value & """"
15          If rsItems.Fields("label").value <> "" Or VBA.IsNull(rsItems.Fields("label").value) = False Then strXML = strXML & " label=""" & rsItems.Fields("label").value & """"
16          If rsItems.Fields("getVisible").value <> "" Then strXML = strXML & " getVisible=""" & rsItems.Fields("getVisible").value & """"
            
            'Permite customizar a visibilidade do elemento por uma ação personalizada
            If rsItems.Fields("tag").value Like "*ValidarVisibilidadePorFuncao:=True*" Then
                bValidaVisible = ValidarVisibilidadePorFuncao(rsItems.Fields("tag").value)
            Else
                bValidaVisible = rsItems.Fields("visible").value
            End If

            'Se é visível ou não
17          If bValidaVisible Then
                strXML = strXML & " visible=""true""" & VBA.vbCrLf & VBA.vbCrLf
            Else
                strXML = strXML & " visible=""false""" & VBA.vbCrLf & VBA.vbCrLf
            End If


22          strXML = strXML & " >"

23          loadElements (rsItems.Fields("id").value)
24          strXML = strXML & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & "</group>" & VBA.vbCrLf & VBA.vbCrLf
25          rsItems.MoveNext
26      Loop Until rsItems.EOF = True
27  End If

28  rsItems.Close

Fim:
29  On Error GoTo 0
30  Exit Function

loadGroup_Error:
31  If VBA.Err <> 0 Then
32      lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
33      Debug.Print "Erro na criação das tabs : " & lngErrorNumber & " - " & strErrorMessagem & "(Line " & VBA.Erl() & " in " & cstr_ProcedureName & ")"
34  End If
    GoTo Fim:
35  Resume
End Function

Private Function loadElements(Grp As String)
    Dim SQL    As String
    Dim rs     As Object
    Dim linha  As String
    Dim Fim_Bloco As Boolean

10  On Error GoTo loadElements_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
20  Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Function AuxRibbonXMLDesinger.loadElements()"
    '----------------------------------------------------------------------------------------------------

30  Quant_Objetos = 0

50  SQL = ""
60  SQL = SQL & VBA.vbNewLine & ""
70  SQL = SQL & VBA.vbNewLine & " SELECT"
80  SQL = SQL & VBA.vbNewLine & "    items.*,"
90  SQL = SQL & VBA.vbNewLine & "    tools.groupControl,"
100 SQL = SQL & VBA.vbNewLine & "    tab.OfficeMenu"
110 SQL = SQL & VBA.vbNewLine & " FROM tblRibbon_Tabs AS tab"
120 SQL = SQL & VBA.vbNewLine & "    INNER JOIN (tblRibbon_Groups AS [group]"
130 SQL = SQL & VBA.vbNewLine & "    INNER JOIN (tblRibbon_ControlType AS tools"
140 SQL = SQL & VBA.vbNewLine & "    INNER JOIN  tblRibbon_Controls AS items ON tools.controlType = items.controlType)"
150 SQL = SQL & VBA.vbNewLine & "    ON group.id = items.parent)"
160 SQL = SQL & VBA.vbNewLine & "    ON tab.id = group.parent"
170 SQL = SQL & VBA.vbNewLine & " WHERE"
180 SQL = SQL & VBA.vbNewLine & "    items.parent='" & Grp & "'"
190 SQL = SQL & VBA.vbNewLine & " ORDER BY"
200 SQL = SQL & VBA.vbNewLine & "    items.ORDER;"


210 Set rs = CurrentDb.OpenRecordset(SQL)

220 If rs.EOF = False Then
230     rs.MoveFirst
240     Do
250         Quant_Objetos = Quant_Objetos + 1
            'inicia a linha
260         If rs.Fields("controlType").value = "dialogBoxLauncher" Then
270             linha = VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & "<dialogBoxLauncher>" & VBA.vbCrLf
280             linha = linha & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & "<button"
290         Else
300             linha = VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & "<" & rs.Fields("controlType").value
310         End If

            'insere o nome do objeto
320         If rs.Fields("Mso").value = 0 Then
330             linha = linha & " id=""" & rs.Fields("id").value & """"
340         Else
350             linha = linha & " idMso=""" & rs.Fields("id").value & """"
360         End If
            'Rótulo
370         If rs.Fields("label").value <> "" And VBA.IsNull(rs.Fields("label").value) = False Then
380             linha = linha & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & " label=""" & rs.Fields("label").value & """" & VBA.vbCrLf
390         End If
            'boxStyle
400         If rs.Fields("boxStyle").value <> "" And VBA.IsNull(rs.Fields("boxStyle").value) = False Then
410             linha = linha & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & " boxStyle=""" & rs.Fields("boxStyle").value & """" & VBA.vbCrLf
420         End If

            'keytip Tecla de atalho
430         If VBA.IsNull(rs.Fields("keytip").value) = False Or rs.Fields("keyTip").value <> "" Then
440             linha = linha & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & " keytip=""" & rs.Fields("keytip").value & """" & VBA.vbCrLf
450         End If

            'Tamanho da imagem do botão
460         If VBA.IsNull(rs.Fields("size").value) = False Or rs.Fields("size").value <> "" Then
470             linha = linha & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & " size=""" & rs.Fields("size").value & """" & VBA.vbCrLf
480         End If

            'tamanho da imagem dos itens do menu
490         If VBA.IsNull(rs.Fields("itemSize").value) = False Or rs.Fields("itemSize").value <> "" Then
500             linha = linha & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & " itemSize=""" & rs.Fields("itemSize").value & """" & VBA.vbCrLf
510         End If
            'Imagem
520         If rs.Fields("imageMso").value <> "" And VBA.IsNull(rs.Fields("imageMso").value) = False Then
530             linha = linha & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & " imageMso=""" & rs.Fields("imageMso").value & """" & VBA.vbCrLf
540         End If

            'screentip
550         If rs.Fields("ScreenTip").value <> "" And VBA.IsNull(rs.Fields("ScreenTip").value) = False Then
560             linha = linha & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & " screentip=""" & rs.Fields("ScreenTip").value & """" & VBA.vbCrLf
570         End If
            'SuperTip
580         If rs.Fields("SuperTip").value <> "" And VBA.IsNull(rs.Fields("SuperTip").value) = False Then
590             linha = linha & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & " supertip=""" & rs.Fields("SuperTip").value & """" & VBA.vbCrLf
600         End If
            'getPressed
610         If rs.Fields("getPressed").value <> "" And VBA.IsNull(rs.Fields("getPressed").value) = False Then
620             linha = linha & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & " getPressed=""" & rs.Fields("getPressed").value & """" & VBA.vbCrLf
630         End If
            'getText
640         If rs.Fields("getText").value <> "" And VBA.IsNull(rs.Fields("getText").value) = False Then
650             linha = linha & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & " getText=""" & rs.Fields("getText").value & """" & VBA.vbCrLf
660         End If
            'getLabel
670         If rs.Fields("getLabel").value <> "" And VBA.IsNull(rs.Fields("getLabel").value) = False Then
680             linha = linha & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & " getLabel=""" & rs.Fields("getLabel").value & """" & VBA.vbCrLf
690         End If
            'getContent
700         If rs.Fields("getContent").value <> "" And VBA.IsNull(rs.Fields("getContent").value) = False Then
710             linha = linha & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & " getContent=""" & rs.Fields("getContent").value & """" & VBA.vbCrLf
720         End If
            'getEnabled
730         If rs.Fields("getEnabled").value <> "" And VBA.IsNull(rs.Fields("getEnabled").value) = False Then
740             linha = linha & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & " getEnabled=""" & rs.Fields("getEnabled").value & """" & VBA.vbCrLf
750         End If
            'onChange
760         If rs.Fields("onChange").value <> "" And VBA.IsNull(rs.Fields("onChange").value) = False Then
770             linha = linha & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & " onChange=""" & rs.Fields("onChange").value & """" & VBA.vbCrLf
780         End If
            'onAction
790         If rs.Fields("onAction").value <> "" And VBA.IsNull(rs.Fields("onAction").value) = False Then
800             linha = linha & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & " onAction=""" & rs.Fields("onAction").value & """" & VBA.vbCrLf
810         End If
            'tag (adelson, em 02/03/2011)
820         If rs.Fields("tag").value <> "" And VBA.IsNull(rs.Fields("tag").value) = False Then
830             linha = linha & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & " tag=""" & rs.Fields("tag").value & """" & VBA.vbCrLf
840         End If
            'sizeString (adelson, em 12/06/2011) - Tamanho de um EditBox
850         If rs.Fields("sizeString").value <> "" And VBA.IsNull(rs.Fields("sizeString").value) = False Then
860             linha = linha & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & " sizeString=""" & rs.Fields("sizeString").value & """" & VBA.vbCrLf
870         End If
            'getSelectedItemIndex (adelson, em 02/03/2011)
880         If rs.Fields("getSelectedItemIndex").value <> "" And VBA.IsNull(rs.Fields("getSelectedItemIndex").value) = False Then
890             linha = linha & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & " getSelectedItemIndex=""" & rs.Fields("getSelectedItemIndex").value & """" & VBA.vbCrLf
900         End If
            'getItemCount (adelson, em 02/03/2011)
910         If rs.Fields("getItemCount").value <> "" And VBA.IsNull(rs.Fields("getItemCount").value) = False Then
920             linha = linha & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & " getItemCount=""" & rs.Fields("getItemCount").value & """" & VBA.vbCrLf
930         End If
            'getItemLabel (adelson, em 02/03/2011)
940         If rs.Fields("getItemLabel").value <> "" And VBA.IsNull(rs.Fields("getItemLabel").value) = False Then
950             linha = linha & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & " getItemLabel=""" & rs.Fields("getItemLabel").value & """" & VBA.vbCrLf
960         End If
            'getItemScreentip (adelson, em 02/03/2011)
970         If rs.Fields("getItemScreentip").value <> "" And VBA.IsNull(rs.Fields("getItemScreentip").value) = False Then
980             linha = linha & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & " getItemScreentip=""" & rs.Fields("getItemScreentip").value & """" & VBA.vbCrLf
990         End If
            'getItemSupertip (adelson, em 02/03/2011)
1000        If rs.Fields("getItemSupertip").value <> "" And VBA.IsNull(rs.Fields("getItemSupertip").value) = False Then
1010            linha = linha & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & " getItemSupertip=""" & rs.Fields("getItemSupertip").value & """" & VBA.vbCrLf
1020        End If
            'getVisible (adelson, em 10/08/2012)
1030        If rs.Fields("getVisible").value <> "" And VBA.IsNull(rs.Fields("getVisible").value) = False Then
1040            linha = linha & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & " getVisible=""" & rs.Fields("getVisible").value & """" & VBA.vbCrLf
1050        Else
                'Se é visível ou não
1060            If rs.Fields("OfficeMenu") = -1 And rs.Fields("controlType").value = "Menu" And rs.Fields("Mso").value = -1 Then
1070                linha = linha & " visible=""true""" & VBA.vbCrLf
1080            End If

1090            If rs.Fields("OfficeMenu").value = 0 Then
1100                If rs.Fields("Visible").value = -1 Then
1110                    linha = linha & " visible=""true""" & VBA.vbCrLf
1120                Else
1130                    linha = linha & " visible=""false""" & VBA.vbCrLf
1140                End If
1150            End If
1160        End If
            'remove último fim de linha
            'Linha = VBA.Left(Linha, Len(Linha) - 2)
            'Verifica se é um comando de bloco
1170        Fim_Bloco = True
1180        If rs.Fields("groupControl").value = 0 Then
1190            linha = linha & " />" & VBA.vbCrLf & VBA.vbCrLf
1200            Fim_Bloco = False
1210        Else
1220            Select Case rs.Fields("comando").value
                    Case "dialogBoxLauncher"
1230                    linha = linha & " />" & VBA.vbCrLf & VBA.vbCrLf
1240                    Fim_Bloco = False
1250                Case "splitButton"
1260                    If rs.Fields("Office") = -1 Then
1270                        linha = linha & " />" & VBA.vbCrLf & VBA.vbCrLf
1280                        Fim_Bloco = False
1290                    Else
1300                        linha = linha & " insertAfterMso=""FileOpenDatabase"" >" & VBA.vbCrLf & VBA.vbCrLf
1310                    End If
1320                Case "Menu"
1330                    If rs.Fields("OfficeMenu").value = -1 Then
1340                        linha = linha & " insertAfterMso=""FileOpenDatabase"" >" & VBA.vbCrLf & VBA.vbCrLf
1350                    Else
1360                        linha = linha & " >" & VBA.vbCrLf & VBA.vbCrLf
1370                    End If
1380                Case Else
1390                    linha = linha & " >" & VBA.vbCrLf & VBA.vbCrLf
1400            End Select
1410        End If
1420        strXML = strXML & linha
1430        linha = ""
1440        If rs.Fields("groupControl").value = -1 Then
1450            Carregar_Item (rs.Fields("id").value)
1460            If Fim_Bloco = True Then
1470                strXML = strXML & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & "</" & rs.Fields("controlType").value & ">" & VBA.vbCrLf & VBA.vbCrLf
1480            End If
1490        End If
1500        rs.MoveNext
1510    Loop Until rs.EOF = True
1520 End If
1530 rs.Close
Fim:
1540 On Error GoTo 0
1550 Exit Function
loadElements_Error:
1560 If VBA.Err <> 0 Then
1570    lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
1580    Debug.Print "Erro na criação das tabs : " & lngErrorNumber & " - " & strErrorMessagem & "(Line " & VBA.Erl() & " in " & cstr_ProcedureName & ")"
1590 End If
    GoTo Fim:
    'Debug Mode
1600 Resume
End Function

Function ValidarVisibilidadePorFuncao(pTag As String)
    Dim strEmpresas As String
    strEmpresas = getTheValue(pTag, "Empresas")
    'ValidarVisibilidadePorFuncao = InStr(strEmpresas, getEmpresa()) > 0 'DCount("*", "tblEmpresa", "HABILITAR=-1 AND COD_EMPRESA IN (" & strEmpresas & ")") > 0
End Function

Function Carregar_Item(obj As String)
    Dim rs     As Object
    Dim SQL    As String

    Dim linha  As String

10  SQL = "SELECT tblRibbon_Item.*, tblRibbon_Item_Tools.Comando, tblRibbon_Item_Tools.Bloco, tblRibbon_Tab.OfficeMenu" & _
          " FROM tblRibbon_Tab INNER JOIN (tblRibbon_Group INNER JOIN (tblRibbon_Object INNER JOIN (tblRibbon_Item_Tools INNER JOIN tblRibbon_Item ON tblRibbon_Item_Tools.Nome = tblRibbon_Item.Tipo) ON tblRibbon_Object.Nome = tblRibbon_Item.Objeto) ON tblRibbon_Group.Nome = tblRibbon_Object.Grupo) ON tblRibbon_Tab.Nome = tblRibbon_Group.Tab" & _
          " WHERE (((tblRibbon_Item.Objeto)='" & obj & "')) ORDER BY tblRibbon_Item.Ordem;"

20  Set rs = CurrentDb.OpenRecordset(SQL)

30  If rs.EOF = False Then
40      rs.MoveFirst
50      Do
60          linha = linha & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & "<" & rs.Fields("comando").value
70          If rs.Fields("Office").value = -1 Then
80              linha = linha & " idMso=""" & rs.Fields("Nome").value & """"
90          Else
100             linha = linha & " id=""" & rs.Fields("Nome").value & """"
110         End If
120         If rs.Fields("comando").value = "menuSeparator" Then
130             If VBA.IsNull(rs.Fields("Descricao").value) = False Or rs.Fields("Descricao").value <> "" Then
140                 linha = linha & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & " title=""" & rs.Fields("Descricao").value & """" & VBA.vbCrLf
150             End If
160         Else
170             If VBA.IsNull(rs.Fields("Descricao").value) = False Or rs.Fields("Descricao").value <> "" Then
180                 linha = linha & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & " label=""" & rs.Fields("Descricao").value & """" & VBA.vbCrLf
190             End If
200         End If
            'boxStyle
210         If rs.Fields("boxStyle").value <> "" And VBA.IsNull(rs.Fields("boxStyle").value) = False Then
220             linha = linha & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & " boxStyle=""" & rs.Fields("boxStyle").value & """" & VBA.vbCrLf
230         End If
            'keytip Tecla de atalho
240         If VBA.IsNull(rs.Fields("keytip").value) = False Or rs.Fields("keyTip").value <> "" Then
250             linha = linha & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & " keytip=""" & rs.Fields("keytip").value & """" & VBA.vbCrLf
260         End If
            'Tamanho da imagem do botão
270         If VBA.IsNull(rs.Fields("Tamanho").value) = False Or rs.Fields("Tamanho").value <> "" Then
280             linha = linha & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & " size=""" & rs.Fields("Tamanho").value & """" & VBA.vbCrLf
290         End If
            'tamanho da imagem dos itens do menu
300         If VBA.IsNull(rs.Fields("TamanhoItem").value) = False Or rs.Fields("TamanhoItem").value <> "" Then
310             linha = linha & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & " itemSize=""" & rs.Fields("TamanhoItem").value & """" & VBA.vbCrLf
320         End If
            'Imagem
330         If rs.Fields("Imagem").value <> "" And VBA.IsNull(rs.Fields("Imagem").value) = False Then
340             linha = linha & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & " imageMso=""" & rs.Fields("Imagem").value & """" & VBA.vbCrLf
350         End If
            'screentip
360         If rs.Fields("ScreenTip").value <> "" And VBA.IsNull(rs.Fields("ScreenTip").value) = False Then
370             linha = linha & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & " screentip=""" & rs.Fields("ScreenTip").value & """" & VBA.vbCrLf
380         End If
            'SuperTip
390         If rs.Fields("SuperTip").value <> "" And VBA.IsNull(rs.Fields("SuperTip").value) = False Then
400             If rs.Fields("OfficeMenu").value = -1 Then
410                 linha = linha & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & " description=""" & rs.Fields("SuperTip").value & """" & VBA.vbCrLf
420             Else
430                 linha = linha & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & " supertip=""" & rs.Fields("SuperTip").value & """" & VBA.vbCrLf
440             End If
450         End If
            'getPressed
460         If rs.Fields("getPressed").value <> "" And VBA.IsNull(rs.Fields("getPressed").value) = False Then
470             linha = linha & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & " getPressed=""" & rs.Fields("getPressed").value & """" & VBA.vbCrLf
480         End If
            'getText
490         If rs.Fields("getText").value <> "" And VBA.IsNull(rs.Fields("getText").value) = False Then
500             linha = linha & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & " getText=""" & rs.Fields("getText").value & """" & VBA.vbCrLf
510         End If
            'getLabel
520         If rs.Fields("getLabel").value <> "" And VBA.IsNull(rs.Fields("getLabel").value) = False Then
530             linha = linha & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & " getLabel=""" & rs.Fields("getLabel").value & """" & VBA.vbCrLf
540         End If
            'getContent
550         If rs.Fields("getContent").value <> "" And VBA.IsNull(rs.Fields("getContent").value) = False Then
560             linha = linha & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & " getContent=""" & rs.Fields("getContent").value & """" & VBA.vbCrLf
570         End If
            'onChange
580         If rs.Fields("onChange").value <> "" And VBA.IsNull(rs.Fields("onChange").value) = False Then
590             linha = linha & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & " onChange=""" & rs.Fields("onChange").value & """" & VBA.vbCrLf
600         End If
            'onAction
610         If rs.Fields("onAction").value <> "" And VBA.IsNull(rs.Fields("onAction").value) = False Then
620             linha = linha & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & " onAction=""" & rs.Fields("onAction").value & """" & VBA.vbCrLf
630         End If
640         strXML = strXML & linha
650         linha = ""
660         If rs.Fields("Bloco").value = -1 Then
670             strXML = strXML & " >" & VBA.vbCrLf & VBA.vbCrLf
680             Carregar_Item_Sub1 (rs.Fields("Nome").value)
690             strXML = strXML & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & "</" & rs.Fields("comando").value & ">" & VBA.vbCrLf & VBA.vbCrLf
700         Else
710             strXML = strXML & " />" & VBA.vbCrLf & VBA.vbCrLf
720         End If
730         rs.MoveNext
740     Loop Until rs.EOF = True
750 End If
760 rs.Close
End Function

Function Carregar_Item_Sub1(obj As String)
    Dim rs     As Object
    Dim SQL    As String

10  SQL = "SELECT tblRibbon_Item_Sub1.*, tblRibbon_Item_Tools.Comando, tblRibbon_Item_Tools.Bloco, tblRibbon_Tab.OfficeMenu" & _
          " FROM tblRibbon_Tab INNER JOIN (tblRibbon_Group INNER JOIN (tblRibbon_Object INNER JOIN (tblRibbon_Item INNER JOIN (tblRibbon_Item_Tools INNER JOIN tblRibbon_Item_Sub1 ON tblRibbon_Item_Tools.Nome = tblRibbon_Item_Sub1.Tipo) ON tblRibbon_Item.Nome = tblRibbon_Item_Sub1.Objeto) ON tblRibbon_Object.Nome = tblRibbon_Item.Objeto) ON tblRibbon_Group.Nome = tblRibbon_Object.Grupo) ON tblRibbon_Tab.Nome = tblRibbon_Group.Tab" & _
          " WHERE (((tblRibbon_Item_Sub1.Objeto)='" & obj & "')) ORDER BY tblRibbon_Item_Sub1.Ordem;"

20  Set rs = CurrentDb.OpenRecordset(SQL)

30  If rs.EOF = False Then
40      rs.MoveFirst
50      Do
60          strXML = strXML & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & "<" & rs.Fields("comando").value & " id=""" & rs.Fields("Nome").value & """" & VBA.vbCrLf
70          If rs.Fields("comando").value = "menuSeparator" Then
80              If VBA.IsNull(rs.Fields("Descricao").value) = False Or rs.Fields("Descricao").value <> "" Then
90                  strXML = strXML & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & " title=""" & rs.Fields("Descricao").value & """" & VBA.vbCrLf
100             End If
110         Else
120             If VBA.IsNull(rs.Fields("Descricao").value) = False Or rs.Fields("Descricao").value <> "" Then
130                 strXML = strXML & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & " label=""" & rs.Fields("Descricao").value & """" & VBA.vbCrLf
140             End If
150         End If

            'keytip Tecla de atalho
160         If VBA.IsNull(rs.Fields("keytip").value) = False Or rs.Fields("keyTip").value <> "" Then
170             strXML = strXML & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & " keytip=""" & rs.Fields("keytip").value & """" & VBA.vbCrLf
180         End If

190         If rs.Fields("Imagem").value <> "" And VBA.IsNull(rs.Fields("imagem").value) = False Then
200             strXML = strXML & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & " imageMso=""" & rs.Fields("Imagem").value & """" & VBA.vbCrLf
210         End If

            'screentip
220         If rs.Fields("ScreenTip").value <> "" And VBA.IsNull(rs.Fields("ScreenTip").value) = False Then
230             strXML = strXML & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & " screentip=""" & rs.Fields("ScreenTip").value & """" & VBA.vbCrLf
240         End If

            'SuperTip
250         If rs.Fields("SuperTip").value <> "" And VBA.IsNull(rs.Fields("SuperTip").value) = False Then
260             If rs.Fields("OfficeMenu").value = -1 Then
270                 strXML = strXML & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & " description=""" & rs.Fields("SuperTip").value & """" & VBA.vbCrLf
280             Else
290                 strXML = strXML & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & " supertip=""" & rs.Fields("SuperTip").value & """" & VBA.vbCrLf
300             End If
310         End If
320         If rs.Fields("OnAction").value <> "" And VBA.IsNull(rs.Fields("OnAction").value) = False Then
330             strXML = strXML & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & VBA.Chr(9) & " onAction=""" & rs.Fields("OnAction").value & """" & VBA.vbCrLf
340         End If
350         strXML = strXML & " />" & VBA.vbCrLf & VBA.vbCrLf
360         rs.MoveNext
370     Loop Until rs.EOF = True
380 End If
390 rs.Close
End Function


