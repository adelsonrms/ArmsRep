VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmSelecionarItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim intCol

Private Sub btnSalvar_Click()
    Call salvaValor("continua_exportacao", 1)
    DoCmd.Close acForm, Me.Name, acSaveYes
End Sub

Private Sub chkSelecionarTodo_Click()
    If Me.sfLista.Form.RecordSource = "" Then Exit Sub
    Call CodeDb.Execute("UPDATE [" & Me.sfLista.Form.RecordSource & "] SET selecionar =  " & CInt(chkSelecionarTodo.value) & "")
    If Me.sfLista.Form.Filter <> "" Then
        Call CodeDb.Execute("UPDATE [" & Me.sfLista.Form.RecordSource & "] SET selecionar = " & CInt(chkSelecionarTodo.value) & " WHERE " & Me.sfLista.Form.Filter)
    End If
    sfLista.Requery
    Me.Requery
End Sub

Private Sub Comando11_Click()
    txtFiltro.value = ""
    Call UpdateSourceSubForm(getFields(txtFiltro.value))
End Sub

Private Sub Fechar_Click()
    chkSelecionarTodo.value = 0
    chkSelecionarTodo_Click
    DoCmd.Close acForm, Me.Name, acSaveYes
End Sub

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : Form_frmSelecionarItems.LoadForm()
' TIPO             : Sub
' DATA/HORA        : 30/09/2015 13:53
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Inicializa as variáveis do formulário
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      30/09/2015 13:53
'---------------------------------------------------------------------------------------
Sub LoadForm()
    Dim tbl As Object 'Recordset
    Dim colDef As String
    Dim sf As Form
    Dim vArs As Variant
    Dim i As Integer
    Dim sArgs As String
    '----------------------------------------------------------------------------------------------------
10  On Error GoTo LoadForm_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
20  Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Sub Form_frmSelecionarItems.LoadForm()"
    '----------------------------------------------------------------------------------------------------

30  Set sf = Me.sfLista.Form
    sArgs = Nz(Me.OpenArgs)
    
40  If Not VBA.IsNull(Me.OpenArgs) Then vArs = VBA.Split(sArgs, ";")

    If VBA.IsEmpty(vArs) Then Exit Sub

    '11/02/2016 - Valida os argumentos enviados. Deve ter pelo menos os 5 previstos
    For i = UBound(vArs) + 1 To 5
        sArgs = sArgs & ";"
    Next i
    vArs = VBA.Split(sArgs, ";")

50  If Not VBA.IsNull(Me.OpenArgs) Then
60      With sf
70          .RecordSource = vArs(0)
            .rowID.ControlSource = .Recordset.Fields(0).Name
80          .chkSelecionar.ControlSource = .Recordset.Fields(1).Name
90          .txtCodigo.ControlSource = .Recordset.Fields(2).Name
100         If .Recordset.Fields.count = 4 Then
110             .txtDescricao.ControlSource = .Recordset.Fields(3).Name
120         Else
130             .txtCodigo.Visible = False
140             .txtCodigo.Width = 0
150             .txtDescricao.ControlSource = .Recordset.Fields(2).Name
160         End If
            .foco.SetFocus
            If vArs(2) Then
                .chkSelecionar.Visible = True
                .optSelecionar.Visible = False
                .optSelecionar.Width = 0
            Else
                .optSelecionar.Visible = True
                .chkSelecionar.Visible = False
                .chkSelecionar.Width = 0
            End If
            'If Not .chkSelecionar.Visible Then .chkSelecionar.Width = 0
170     End With
180     Me.chkSelecionarTodo.Visible = vArs(2)
190 End If
200 sfLista.Requery
210 Call salvaValor("continua_exportacao", 0)
220 Me.lblSelecionarTudo.Visible = Me.chkSelecionarTodo.Visible
    '11/02/2016 - Opção para selecionar todos automaticamente
    If VBA.IsNumeric(vArs(3)) Then
        Me.chkSelecionarTodo.value = CBool(vArs(3))
        Call chkSelecionarTodo_Click
    End If
    '11/02/2016 - Argumento para modificar o titulo do botão e da barra de titulo
    If vArs(4) <> "" Then txtCaption.Caption = vArs(4) Else txtCaption.Caption = "&Selecionar Items"
    If vArs(5) <> "" Then btnSalvar.Caption = vArs(5) Else btnSalvar.Caption = "&Continuar"
    
    Call AuxForm.AutoFitColumns(Me.sfLista)
Fim:
230 On Error GoTo 0
240 Exit Sub

LoadForm_Error:
250 If VBA.Err <> 0 Then
260     lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
270     Call Excecoes.TratarErro(strErrorMessagem, lngErrorNumber, cstr_ProcedureName, VBA.Erl)
280 End If
    GoTo Fim:
    'Debug Mode
290 Resume
End Sub

Private Sub Form_Load()
    Call LoadForm
End Sub

Private Sub txtFiltro_Change()
    Call UpdateSourceSubForm(getFields(txtFiltro.Text))
End Sub

'Atualiza o filtro do formulário
Sub UpdateSourceSubForm(pSQLSource As String)
    Dim tbl As Object 'Recordset
30  Me.sfLista.Form.Filter = pSQLSource
    Me.sfLista.Form.FilterOn = True
End Sub

'Monta o criterio de filtro do formulário
Function getFields(Optional pValue As String) As String
    Dim tbl    As Object ' Recordset
    Dim pFilter As String
    Set tbl = sfLista.Form.Recordset
    If tbl Is Nothing Then Exit Function
    pFilter = ""
    pValue = VBA.Replace(pValue, "'", "")
    For Each vl In VBA.Split(pValue, ";")
        If vl <> "" Then
            For i = 0 To tbl.Fields.count - 1
                If InStr(CStr(vl), "*") = 0 Then
                    pFilter = pFilter & " OR ([" & tbl.Fields(i).Name & "] LIKE '*" & CStr(vl) & "*')"
                Else
                    pFilter = pFilter & " OR ([" & tbl.Fields(i).Name & "] LIKE '" & CStr(vl) & "')"
                End If
            Next i
        End If
    Next
    pFilter = Mid(pFilter, 4)
    getFields = pFilter
End Function


