VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmIniciar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Modulo....: Form_frmIniciar/ Documento VBA
' Rotina....: Form_Load / Sub
' Autor.....: Victor Félix
' Contato...: victor.santos@mondial.com.br
' Data......: 24/04/2013
' Empresa...: Mondial Tecnologia em Informática LTDA.
' Descrição.: Evento de inicialização do formulário de inicialização da aplicação. Realiza
' a atualização do vínculo das tabelas vinculadas.
'---------------------------------------------------------------------------------------

Private Sub Form_Load()
10  On Error GoTo TratarErro
    'Stop
    
    Dim strAux As String

    'DoCmd.openForm "frmStart", , , , , acHidden
30  cxTarja.BackColor = pegarVariavelAplicacao("color_bar")
40  lblAppName.Caption = VBA.Replace(pegarVariavelAplicacao("appName"), "&", "&&")

50  Call Inicializar_Globais

60  strAux = AuxAplicacao.PegaEnderecoBDs()

70  FalhaNoLink = AuxTabela.AtualizarVinculos(strAux)

    If Not FalhaNoLink Then
        If AuxMensagens.MessageBoxMaster("F003") = VBA.vbYes Then
110         Call Access.DoCmd.openForm("frmCaminhoBE")
120     Else
130         Call AuxMensagens.MessageBoxMaster("F004")
140         FalhaNoLink = True
150     End If
    End If

260 Call salvaValor("DB_PATH", strAux)

270 'Call InicializaSessaoUsuario
    Call AuxAplicacao.salvaValor("DataHoraLogon", VBA.Now)

    Call closeForm("frmSplah")

280 Exit Sub
TratarErro:
290 Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "frmIniciar.Form_Load()")
300 Exit Sub
310 Resume

End Sub


'---------------------------------------------------------------------------------------
' Modulo....: Form_frmIniciar/ Documento VBA
' Rotina....: Form_Timer / Sub
' Autor.....: Victor Félix
' Contato...: victor.santos@mondial.com.br
' Data......: 24/04/2013
' Empresa...: Mondial Tecnologia em Informática LTDA.
' Descrição.: Evento timer do formulário. Fecha o formulário e habilita a aplicação ao
' usuário
'---------------------------------------------------------------------------------------
Private Sub Form_Timer()
    On Error GoTo TratarErro
    'Call Access.DoCmd.Close(acForm, Me.Name)
'    Stop
    Me.Visible = False
    Call Access.DoCmd.openForm("frmPrincipal")
    Me.TimerInterval = 0
    Exit Sub
TratarErro:
    'Call TratarErro(VBA.Err.Description, VBA.Err.Number, "frmIniciar.Timer")
End Sub

