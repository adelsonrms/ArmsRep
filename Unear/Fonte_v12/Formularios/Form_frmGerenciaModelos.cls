VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmGerenciaModelos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnEditarMSG_Click()
    If Not FileExists(txtArquivoModelo.value) Then
        MsgBox "Modelo da mensagem não encontrada !", vbExclamation
        Exit Sub
    End If
    Call VBA.Shell("explorer """ & txtArquivoModelo.value & """")
End Sub

Private Sub btnEscolher_Click()
   Call CadastrarModeloDeEmail
   Me.sfGrupo.Requery
End Sub

Private Sub Comando85_Click()
    If validaCampos Then
        Call salvaValor("ContinuarEnvio", "Continuar")
        Call salvaValor("ModeloSelecionado", cmbIDModelo.value)
        Call FecharObjeto(Me.Name, acSaveNo)
    End If
End Sub

Private Sub Comando86_Click()
    Call FecharObjeto(Me.Name, acSaveNo)
End Sub

Private Sub Form_Load()
    Call salvaValor("ListaSelecionada", strAquivo)
    Call salvaValor("ContinuarEnvio", "")
End Sub

Function validaCampos() As Boolean
    If Nz(txtListaEmails.value) = "" Then MsgBox "Escolha uma lista de Clientes !", vbExclamation, "Validaçao": Exit Function
    If Nz(cmbIDModelo.value) = "" Then MsgBox "Selecione um modelo antes de continuar!", vbExclamation, "Validaçao": cmbIDModelo.SetFocus: Exit Function
    validaCampos = True
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : mBL_MailMarketing.CadastrarModeloDeEmail()
' TIPO             : Sub
' DATA/HORA        : 11/07/2017 10:42
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Cadastra um modelo de mensagem no banco de dados a partir de um arquivo selecionado
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      11/07/2017 10:42
'---------------------------------------------------------------------------------------
Private Sub CadastrarModeloDeEmail()
    '--------------------------------------------------------------------------------------------------------------
    On Error GoTo Error_Handle
    Dim lngErrorNumber As Long, strErrorMessagem As String: Const cstr_ProcedureName As String = "CadastrarModeloDeEmail()"
    '--------------------------------------------------------------------------------------------------------------
    Dim strFiltro As String, strAquivo As String, rsLog As Object
    Dim vArquivo
    Dim cArq
140 strFiltro = "Mensagem do Outlook (msg) ;*.msg"

160 vArquivo = CaixaDeDialogo(msoFileDialogOpen, _
                "Selecionar uma mensagem de email para modelo", _
                 True, _
                "Escolher", _
                "Outros Arquivos;*.*|" & strFiltro, _
                PegaEndereco_Templates())
    
    If Not VBA.IsArray(vArquivo) Then vArquivo = Array(VBA.CStr(vArquivo))
    
    For Each cArq In vArquivo
        strAquivo = VBA.CStr(cArq)
170     If strAquivo <> "" Then
            strAquivo = VBA.Dir(strAquivo)
            Set rsLog = CurrentDb.OpenRecordset("Pegar_ModeloPorNome", VBA.Left(strAquivo, VBA.Len(strAquivo) - 4))
            If rsLog.EOF Then rsLog.addNew Else rsLog.edit
            rsLog.Fields("NomeModelo").value = VBA.Left(strAquivo, VBA.Len(strAquivo) - 4)
            rsLog.Update
        End If
    Next cArq
Fim:
Exit Sub
'-------------------------------------------------------
'Tratamento de erros não previstos
'-------------------------------------------------------
Error_Handle:
    If VBA.Err <> 0 Then
          lngErrorNumber = VBA.Err.Number: strErrorMessagem = "Erro ao localizar a mensagem no Outlook - " & VBA.Err.Description
          Debug.Print cstr_ProcedureName, lngErrorNumber & "-" & strErrorMessagem
          Call VBA.MsgBox(VBA.vbNewLine & cstr_ProcedureName & " Em" & VBA.IIf(VBA.Erl() <> 0, " >(" & VBA.Erl() & ")", "") & " > " & strErrorMessagem, VBA.vbCritical)
          Call VBA.Err.Raise(VBA.Err.Number, cstr_ProcedureName, VBA.vbNewLine & cstr_ProcedureName & " Em" & VBA.IIf(VBA.Erl() <> 0, " >(" & VBA.Erl() & ")", "") & " > " & strErrorMessagem)
          GoTo Fim
    End If
    'Debug Mode
    Resume
End Sub

