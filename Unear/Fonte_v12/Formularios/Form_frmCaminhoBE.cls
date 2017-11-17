VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmCaminhoBE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnTabelasVinculadas_Click()
    'DoCmd.openForm "frmTabelasVinculadas", , , , , acDialog
    Call AbrirSubFormulario(objFonte:="frmTabelasVinculadas", Legenda:="Configuração Back-End / Tabelas Vinculadas", bReabrirForm:=False)
End Sub

Private Sub Form_Load()
    grpOpcaoDB_AfterUpdate
End Sub
'---------------------------------------------------------------------------------------
' Modulo....: Form_frmCaminhoBE/ Documento VBA
' Rotina....: Form_Open / Sub
' Autor.....: Victor Félix
' Contato...: victor.santos@mondial.com.br
' Data......: 23/04/2013
' Empresa...: Mondial Tecnologia em Informática LTDA.
' Descrição.: Evento de abertura do formulário. Carrega, da tabela de configurações, o
' caminho atual do BackEnd da aplicação
'---------------------------------------------------------------------------------------
Private Sub Form_Open(Cancel As Integer)
    On Error GoTo TratarErro
    Me.ctLocal.value = getDBPath
    'Call AuxForm.Atualizar_Campo_Versao(Me.lblVersao)
    Exit Sub
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "frmCaminhoBE.Form_Open()", Erl)
End Sub

Private Sub btnCancelar_Click()
    On Error GoTo TratarErro
    Call Access.DoCmd.Close(acForm, Me.Name)
    On Error GoTo 0
    Exit Sub
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "Form_frmCaminhoBE.btnCancelar_Click", Erl)
End Sub

'---------------------------------------------------------------------------------------
' Modulo....: Form_frmCaminhoBE/ Documento VBA
' Rotina....: btnOK_Click / Sub
' Autor.....: Victor Félix
' Contato...: victor.santos@mondial.com.br
' Data......: 23/04/2013
' Empresa...: Mondial Tecnologia em Informática LTDA.
' Descrição.: Evento de clique do botão "OK". Realiza a operação de atualização do vínculo
' das tabelas vinculadas. Se a tentativa não for bem sucedida emite um aviso ao usuário
'---------------------------------------------------------------------------------------
Private Sub btnOK_Click()
10  On Error GoTo TratarErro
    Dim sPath As String
20  sPath = Me.ctLocal.value

    If Not AuxDataBase.RecriarVinculoTabelasCadastradas(sPath) Then
        Call AuxMensagens.MessageBoxMaster("A atualização do Vínculo da Aplicação com o banco de dados falhou. " & _
                                           "Cheque o caminho configurado para o BackEnd da aplicação", _
                                          VBA.vbCritical + VBA.vbOKOnly, "Atualização de Vínculo")
    Else
        sPath = Me.ctLocal.value
        Call AuxMensagens.MessageBoxMaster("F005")
        Call Access.DoCmd.Close(acForm, Me.Name)
    End If

    Call salvaVariavelAplicacao("DB_PATH", sPath)

170 Exit Sub
TratarErro:
180 Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "frmCaminhoBE.btnOK_Click", Erl)
190 Exit Sub
200 Resume
End Sub

'---------------------------------------------------------------------------------------
' Modulo....: Form_frmCaminhoBE/ Documento VBA
' Rotina....: btnLocalizarArquivo_Click / Sub
' Autor.....: Victor Félix
' Contato...: victor.santos@mondial.com.br
' Data......: 23/04/2013
' Empresa...: Mondial Tecnologia em Informática LTDA.
' Descrição.: Evento do botão de procurar arquivos. Exibe a opção do usuário setar o
' arquivo Access que contém as tabelas de BackEnd da aplicação
'---------------------------------------------------------------------------------------
Private Sub btnLocalizarArquivo_Click()
10    On Error GoTo TratarErro
    Dim strEndereco As String

20    On Error GoTo TratarErro
    Dim fso As Object
    Dim endereco As String
    Dim arrCaminho As Variant
    
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    
    If VBA.MsgBox("Atenção !" & VBA.vbNewLine & VBA.vbNewLine & "O local escolhido deve ser uma pasta raiz onde as demais subpastas (Dados e Configurações) deverão existir", VBA.vbExclamation + VBA.vbYesNo, "Local do Banco de dados") = VBA.vbNo Then Exit Sub
    

30    endereco = CaixaDeDialogo(msoFileDialogFolderPicker, "Selecione o Back-End de Dados")

40    If Not endereco = "Vazio" And Not endereco = "" Then
50      Me.ctLocal.value = endereco
60      Me.Requery
70      Call Publicas.Inicializar_Globais

80      arrCaminho = Conexao.PegarArray("Pegar_CaminhoBE", Publicas.ChaveUsuario)
90      If VBA.IsArray(arrCaminho) Then
100         Call Conexao.AlterarRegistros("Atualizar_CaminhoBE", Publicas.ChaveUsuario, endereco)
110     Else
120         Call Conexao.InserirRegistros("Inserir_CaminhoBE", Publicas.ChaveUsuario, endereco)
130     End If
140   Else
150     Call AuxMensagens.MessageBoxMaster("F006")
160   End If
Fim:
170   Me.ctLocal.value = endereco
180   Call salvaValor("DBConfig_CustomDBRootPath", ctLocal.value)

190   Exit Sub
TratarErro:
200   Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "frmCaminhoBE.btnLocalizarArquivo_Click()", Erl)
End Sub

Private Sub grpOpcaoDB_AfterUpdate()
    Dim arrCaminho, endereco As String
10  If grpOpcaoDB.value = 1 Then
20      grpOpcaoDB.tag = PegaEnderecoBDs
        Inicializar_Globais
30      arrCaminho = Conexao.PegarArray("Pegar_CaminhoBE", Publicas.ChaveUsuario)
40      If VBA.IsArray(arrCaminho) Then
50          Call Conexao.AlterarRegistros("Atualizar_CaminhoBE", Publicas.ChaveUsuario, grpOpcaoDB.tag)
60      Else
70          Call Conexao.InserirRegistros("Inserir_CaminhoBE", Publicas.ChaveUsuario, grpOpcaoDB.tag)
80      End If
90  Else
100     grpOpcaoDB.tag = pegaValor("DBConfig_CustomDBRootPath")
110 End If
120 Call salvaValor("DBConfig_UseCurrentDBRootPath", CInt(grpOpcaoDB.value = 1))
130 ctLocal.Enabled = grpOpcaoDB.value <> 1
140 btnLocalizarArquivo.Enabled = ctLocal.Enabled
150 ctLocal.value = grpOpcaoDB.tag

End Sub
