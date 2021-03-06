VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMsgBoxMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Modulo....: Form_frmVBA.MsgBoxMaster / Formul�rio
' Rotina....: btn1_Click() / Sub
' Autor.....: Victor F�lix
' Contato...: victor.santos@mondial.com.br
' Data......: 03/04/2012
' Empresa...: Mondial Tecnologia em Inform�tica LTDA.
' Descri��o.: "Seta" um valor do Enum VbMsgBoxResult em uma vari�vel global para o retorno
' do valor do bot�o clicado pelo usu�rio
'---------------------------------------------------------------------------------------
Private Sub btn1_Click()
On Error GoTo TrataErro
    VariaveisEConstantes.Resultado = Me.btn1.tag
    Me.Visible = False
Exit Sub
TrataErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "frmVBA.MsgBoxMaster.btn1_Click()", Erl)
End Sub

'---------------------------------------------------------------------------------------
' Modulo....: Form_frmVBA.MsgBoxMaster / Formul�rio
' Rotina....: btn2_Click() / Sub
' Autor.....: Victor F�lix
' Contato...: victor.santos@mondial.com.br
' Data......: 03/04/2012
' Empresa...: Mondial Tecnologia em Inform�tica LTDA.
' Descri��o.: "Seta" um valor do Enum VbMsgBoxResult em uma vari�vel global para o retorno
' do valor do bot�o clicado pelo usu�rio
'---------------------------------------------------------------------------------------
Private Sub btn2_Click()
On Error GoTo TrataErro
    VariaveisEConstantes.Resultado = Me.btn2.tag
    Me.Visible = False
Exit Sub
TrataErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "frmVBA.MsgBoxMaster.btn2_Click()", Erl)
End Sub

'---------------------------------------------------------------------------------------
' Modulo....: Form_frmVBA.MsgBoxMaster / Formul�rio
' Rotina....: btn3_Click() / Sub
' Autor.....: Victor F�lix
' Contato...: victor.santos@mondial.com.br
' Data......: 03/04/2012
' Empresa...: Mondial Tecnologia em Inform�tica LTDA.
' Descri��o.: "Seta" um valor do Enum VbMsgBoxResult em uma vari�vel global para o retorno
' do valor do bot�o clicado pelo usu�rio
'---------------------------------------------------------------------------------------
Private Sub btn3_Click()
On Error GoTo TrataErro
    VariaveisEConstantes.Resultado = Me.btn3.tag
    Me.Visible = False
Exit Sub
TrataErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "frmVBA.MsgBoxMaster.btn3_Click()", Erl)
    Resume Next
End Sub
'---------------------------------------------------------------------------------------
' Modulo....: Form_frmVBA.MsgBoxMaster / Formul�rio
' Rotina....: Form_Open / Sub
' Autor.....: Victor F�lix
' Contato...: victor.santos@mondial.com.br
' Data......: 03/04/2012
' Empresa...: Mondial Tecnologia em Inform�tica LTDA.
' Descri��o.: Rotina de Abertura do formul�rio. Trata a propriedade OpenArgs do Formul�rio
' para a format��o de acordo com a MessageBox desejada
'---------------------------------------------------------------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo TrataErro
Dim arrAux      As Variant
Dim TipoAux     As VbMsgBoxStyle
    With Me
        .txtMensagem.ForeColor = pegarVariavelAplicacao("color_bar")
        .lblTitulo.ForeColor = .txtMensagem.ForeColor
        
        If Not VBA.IsNull(.OpenArgs) Then
            arrAux = VBA.Split(.OpenArgs, "|")
            TipoAux = arrAux(2)
            Select Case TipoAux
                Case VBA.vbYesNo
                    Call AtivaBtnSim(.btn2)
                    Call AtivaBtnNao(.btn3)
                Case VBA.vbCritical
                    .imgCritical.Visible = True
                    Call AtivaBtnOK(.btn3)
                    Call FormataComImagens
                Case VBA.vbOKOnly
                    Call AtivaBtnOK(.btn3)
                Case VBA.vbInformation
                    .imgInfo.Visible = True
                    Call AtivaBtnOK(.btn3)
                    Call FormataComImagens
                Case VBA.vbYesNoCancel
                    Call AtivaBtnSim(.btn1)
                    Call AtivaBtnNao(.btn2)
                    Call AtivaBtnCancel(.btn3)
                Case VBA.vbOKCancel
                    Call AtivaBtnOK(.btn2)
                    Call AtivaBtnCancel(.btn3)
                Case VBA.vbAbortRetryIgnore
                    Call AtivaBtnAbort(.btn1)
                    Call AtivaBtnRetry(.btn2)
                    Call AtivaBtnIgnore(.btn3)
                Case VBA.vbExclamation
                    Call AtivaBtnOK(.btn3)
                    .imgQuestion.Visible = False
                    .imgCritical.Visible = False
                    .imgInfo.Visible = False
                    .imgExclamation.Visible = True
                    Call FormataComImagens
                Case VBA.vbQuestion
                    Call AtivaBtnOK(.btn3)
                    .imgQuestion.Visible = True
                    Call FormataComImagens
                Case VBA.vbCritical + VBA.vbYesNo
                    Call AtivaBtnSim(.btn2)
                    Call AtivaBtnNao(.btn3)
                    Call FormataComImagens
                    .imgCritical.Visible = True
                Case VBA.vbInformation + VBA.vbYesNo
                    Call AtivaBtnSim(.btn2)
                    Call AtivaBtnNao(.btn3)
                    Call FormataComImagens
                    .imgInfo.Visible = True
                Case VBA.vbQuestion + VBA.vbYesNo
                    Call AtivaBtnSim(.btn2)
                    Call AtivaBtnNao(.btn3)
                    Call FormataComImagens
                    .imgQuestion.Visible = True
                Case VBA.vbExclamation + VBA.vbYesNo
                    Call FormataComImagens
                    .imgExclamation.Visible = True
                    Call AtivaBtnSim(.btn2)
                    Call AtivaBtnNao(.btn3)
                Case VBA.vbAbortRetryIgnore + VBA.vbInformation
                    Call FormataComImagens
                    Call AtivaBtnAbort(.btn1)
                    Call AtivaBtnRetry(.btn2)
                    Call AtivaBtnIgnore(.btn3)
                    .imgInfo.Visible = True
            End Select
            
            .lblTitulo.Caption = arrAux(0)
            .txtMensagem = arrAux(1)
            With .txtMensagem
                .SetFocus
                If arrAux(1) = .SelText Then
                    .ScrollBars = 0
                    .SelStart = 0
                    .SelLength = 0
                    .Enabled = False
                Else
                    .ScrollBars = 2
                    .SelStart = 0
                    .SelLength = 0
                End If
            End With
        End If
    End With
Exit Sub
TrataErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "CaixaDeMensagem.Form_Load", Erl)
Resume
End Sub

'---------------------------------------------------------------------------------------
' Modulo....: Form_frmVBA.MsgBoxMaster / Formul�rio
' Rotina....: Ativabtns / Sub
' Autor.....: Victor F�lix
' Contato...: victor.santos@mondial.com.br
' Data......: 03/04/2012
' Empresa...: Mondial Tecnologia em Inform�tica LTDA.
' Descri��o.: Rotinas para a configura��o dos bot�es a serem exibidos na MessageBox
' Personalizada
'---------------------------------------------------------------------------------------
Private Sub FormataComImagens()
On Error GoTo TrataErro

    With Me
        '.lblMensagem.Width = .lblMensagem.Width - 660 '400
        .txtMensagem.Width = .txtMensagem.Width - 960 '400
        .lblTitulo.Width = .lblTitulo.Width - 660 '400
        .Linha4.Width = .Linha4.Width '- 400

        '.lblMensagem.Left = .lblMensagem.Left + 660
        .txtMensagem.Left = .txtMensagem.Left + 960
        .lblTitulo.Left = .lblTitulo.Left + 660
        .Linha4.Left = .Linha4.Left '+ 400
    End With

Exit Sub
TrataErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "", Erl)
End Sub

'---------------------------------------------------------------------------------------
' Modulo....: Form_frmVBA.MsgBoxMaster/ Documento VBA
' Rotina....: AtivaBtnSim / Sub
' Autor.....: Victor F�lix
' Contato...: victor.santos@mondial.com.br
' Data......: 24/04/2013
' Empresa...: Mondial Tecnologia em Inform�tica LTDA.
' Descri��o.: Formata o controle Access especificado com o formato "Sim"
'---------------------------------------------------------------------------------------
Private Sub AtivaBtnSim(ByRef accButton As Access.CommandButton)
On Error GoTo TrataErro
    With accButton
        .Caption = "Sim"
        .Visible = True
        .tag = VBA.vbYes
    End With
Exit Sub
TrataErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "CaixaDeMensagem.AtivaBtnSim", Erl)
End Sub


'---------------------------------------------------------------------------------------
' Modulo....: Form_frmVBA.MsgBoxMaster/ Documento VBA
' Rotina....: AtivaBtnNao / Sub
' Autor.....: Victor F�lix
' Contato...: victor.santos@mondial.com.br
' Data......: 24/04/2013
' Empresa...: Mondial Tecnologia em Inform�tica LTDA.
' Descri��o.: Formata o controle Access especificado com o formato "N�o"
'---------------------------------------------------------------------------------------
Private Sub AtivaBtnNao(ByRef accButton As Access.CommandButton)
On Error GoTo TrataErro
    With accButton
        .Caption = "N�o"
        .Visible = True
        .tag = VBA.vbNo
    End With
Exit Sub
TrataErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "CaixaDeMensagem.AtivaBtnSim", Erl)
End Sub

'---------------------------------------------------------------------------------------
' Modulo....: Form_frmVBA.MsgBoxMaster/ Documento VBA
' Rotina....: AtivaBtnCancel / Sub
' Autor.....: Victor F�lix
' Contato...: victor.santos@mondial.com.br
' Data......: 24/04/2013
' Empresa...: Mondial Tecnologia em Inform�tica LTDA.
' Descri��o.: Formata o controle Access especificado com o formato "Cancelar"
'---------------------------------------------------------------------------------------
Private Sub AtivaBtnCancel(ByRef accButton As Access.CommandButton)
On Error GoTo TrataErro
    With accButton
        .Caption = "Cancelar"
        .Visible = True
        .tag = VBA.vbCancel
    End With
Exit Sub
TrataErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "CaixaDeMensagem.AtivaBtnSim", Erl)
End Sub

'---------------------------------------------------------------------------------------
' Modulo....: Form_frmVBA.MsgBoxMaster/ Documento VBA
' Rotina....: AtivaBtnOK / Sub
' Autor.....: Victor F�lix
' Contato...: victor.santos@mondial.com.br
' Data......: 24/04/2013
' Empresa...: Mondial Tecnologia em Inform�tica LTDA.
' Descri��o.: Formata o controle Access especificado com o formato "OK"
'---------------------------------------------------------------------------------------
Private Sub AtivaBtnOK(ByRef accButton As Access.CommandButton)
On Error GoTo TrataErro
    With accButton
        .Caption = "OK"
        .Visible = True
        .tag = VBA.vbOK
    End With
Exit Sub
TrataErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "CaixaDeMensagem.AtivaBtnSim", Erl)
End Sub

'---------------------------------------------------------------------------------------
' Modulo....: Form_frmVBA.MsgBoxMaster/ Documento VBA
' Rotina....: AtivaBtnAbort / Sub
' Autor.....: Victor F�lix
' Contato...: victor.santos@mondial.com.br
' Data......: 24/04/2013
' Empresa...: Mondial Tecnologia em Inform�tica LTDA.
' Descri��o.: Formata o controle Access especificado com o formato "Anular"
'---------------------------------------------------------------------------------------
Private Sub AtivaBtnAbort(ByRef accButton As Access.CommandButton)
On Error GoTo TrataErro
    With accButton
        .Caption = "Anular"
        .Visible = True
        .tag = VBA.vbAbort
    End With
Exit Sub
TrataErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "CaixaDeMensagem.AtivaBtnSim", Erl)
End Sub

'---------------------------------------------------------------------------------------
' Modulo....: Form_frmVBA.MsgBoxMaster/ Documento VBA
' Rotina....: AtivaBtnRetry / Sub
' Autor.....: Victor F�lix
' Contato...: victor.santos@mondial.com.br
' Data......: 24/04/2013
' Empresa...: Mondial Tecnologia em Inform�tica LTDA.
' Descri��o.: Formata o controle Access especificado com o formato "Repetir"
'---------------------------------------------------------------------------------------
Private Sub AtivaBtnRetry(ByRef accButton As Access.CommandButton)
On Error GoTo TrataErro
    With accButton
        .Caption = "Repetir"
        .Visible = True
        .tag = VBA.vbRetry
    End With
Exit Sub
TrataErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "CaixaDeMensagem.AtivaBtnSim", Erl)
End Sub

'---------------------------------------------------------------------------------------
' Modulo....: Form_frmVBA.MsgBoxMaster/ Documento VBA
' Rotina....: AtivaBtnIgnore / Sub
' Autor.....: Victor F�lix
' Contato...: victor.santos@mondial.com.br
' Data......: 24/04/2013
' Empresa...: Mondial Tecnologia em Inform�tica LTDA.
' Descri��o.: Formata o controle Access especificado com o formato "Ignorar"
'---------------------------------------------------------------------------------------
Private Sub AtivaBtnIgnore(ByRef accButton As Access.CommandButton)
On Error GoTo TrataErro
    With accButton
        .Caption = "Ignorar"
        .Visible = True
        .tag = VBA.vbIgnore
    End With
Exit Sub
TrataErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "CaixaDeMensagem.AtivaBtnSim", Erl)
End Sub
'---------------------------------------------------------------------------------------
