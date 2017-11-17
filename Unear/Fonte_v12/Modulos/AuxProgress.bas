Attribute VB_Name = "AuxProgress"
Option Explicit
Public sCodItemSelected As String

Public Const C_FORM_NAME_PROGRESS As String = "frmMessage"
Public Const C_FORM_NAME_PROGRESS2 As String = "frmMessage_ListBox"

'Variáveis globais que são usadas no progresso
Public strMsg As String   'Mensagem da tela
Public sArgumentos As String
Public strTitulo As String   'Mensagem da tela
Public strMacro As String   'Mensagem da tela
Public strTextBotaoEntrar As String   'Mensagem da tela
Public bolBotaoState As Boolean
Public strMsgExtendida As String
Public strErro_number As String   'Mensagem da tela
Public strRet_action_fechar As String   'Mensagem da tela
Public strText_botao_ext As String   'Mensagem da tela
Public strOnAction_botao_ext As String   'Mensagem da tela
Public intProg_TipoVisualizacao As Integer
Public intProg_IniciarContinuo As Integer
Public prog_old_icon As String
Public prog_old_atual As String
Public bShowBntCancelar As Boolean
Public bShowBntFechar As Boolean

Public strIDEvent As String
Public vTimeProgress As Date
Public CancelProgress As Boolean

Private Type Mensagem
    ID      As String
    Titulo  As String
    Estilo  As VbMsgBoxStyle
    Texto   As String
End Type

Public Enum eModoProgress
    SemProgresso = 0
    INICIO = 1
    Fim = 2
    EmEndamento = 3
    EmEndamento_MensagemFixa = 4
    EmEndamento_ProgressoFixo = 5
End Enum

Public Enum eIconeStatusMessage
    Icone_Erro = 0
    Icone_Sucesso = -1
    Icone_EmAndamento = 1
    Icone_Aviso = 2
End Enum

Private f As Form_frmMessage

'*********************************************************************************************************
' FUNÇÕES PROGRESS - FRAMEWORK
'*********************************************************************************************************
Sub ConfigurarProgresso(Optional dblIncremento As Double, Optional Message As String, Optional bShowInStatusBar As Boolean = True, Optional pProgressType As String = "List")
    Dim frm As Form
10  If pProgressType = "List" Then
20      Call AbrirFormulario("frmProgressoLista", Access.AcFormView.acNormal)
        Set frm = Forms("frmProgressoLista")
30      frm.InsideHeight = (frm.CabeçalhoDoFormulário.Height + frm.RodapéDoFormulário.Height) + 1000
40      frm.lstItems.Height = (frm.InsideHeight - frm.CabeçalhoDoFormulário.Height) - 500
50      Call MostrarProgresso(0, Message)
60  Else
70      Call AbrirFormulario("frmProgresso", Access.AcFormView.acNormal)
80  End If
    If bShowInStatusBar Then Call SysCmd(acSysCmdSetStatus, Message)
End Sub

Sub MostrarProgresso(dblIncremento As Double, _
                 Optional Message As String, _
                 Optional bShowInStatusBar As Boolean = True, _
                 Optional vStatus As Variant, Optional pProgressType As String = "List")

10  If pProgressType = "List" Then
20      Call AuxForm.IncrementaBarraProgressoLista(dblIncremento, "Aguarde...", Message, CStr(vStatus))
30  Else
40      Call AuxForm.IncrementaBarraProgresso(dblIncremento, CStr(vStatus))
50  End If
60  If bShowInStatusBar Then Call SysCmd(acSysCmdSetStatus, Message)
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------
' FUNÇÕES PARA MANIPULAÇÃO DE TELA DE MENSAGEMS (STATUS)
'----------------------------------------------------------------------------------------------------------------------------------------
Function AbrirTelaProgresso(Optional pProgressType As String = "List")
10  If pProgressType = "List" Then
20      Call AbrirFormulario("frmProgressoLista", Access.AcFormView.acNormal)
30  Else
40      Call AbrirFormulario("frmProgresso", Access.AcFormView.acNormal)
        Forms("frmProgresso").Progresso.Width = 0
        Forms("frmProgresso").Caption = "0%"
50  End If
End Function

Sub configFormMensagem(Optional intPagina As Integer = 0, _
                       Optional UpdateOnlyMensagens As Boolean = False, _
                       Optional lngMaxProgress As Long, _
                       Optional bShowCancelar As Boolean)

    Set f = New Form_frmMessage
    With f
        .StartTime = Time
        .defualtPagina = intPagina
        .MaxProgress = lngMaxProgress
        .UpdateOnlyMensagens = UpdateOnlyMensagens
    End With
End Sub

Sub MostraFormSplash(sFormName As String, Optional bModal As Boolean = False)
    If Not FormIsOpen(sFormName) Then
        Call openForm(sFormName, bModal)
    Else
        Call Forms(sFormName).LoadProgressMensagems
    End If
    If FormIsOpen(sFormName) Then Call Forms(sFormName).Repaint
End Sub

'Mostra o form de mensagens
Sub showMessage(Optional sMsg As String, _
                Optional sTitulo As String, _
                Optional ByVal modoProgress As eModoProgress = eModoProgress.SemProgresso, _
                Optional ByVal Progress_Max As Long = -1, _
                Optional bShowBntCancelar As Boolean = False, _
                Optional bShowBntFechar As Boolean = False, _
                Optional icone As eIconeStatusMessage = eIconeStatusMessage.Icone_EmAndamento, _
                Optional intTimer As Integer = 0)
    
    Dim strMsgContinua As String
    On Error Resume Next 'Copia
    strMsg = ""
    strMsg = sMsg
    strTitulo = sTitulo
    Dim arrInfo As Variant
    Dim strArgs As String
    
    
    If modoProgress = Fim Then
        sTitulo = "Finalizado !"
        Call FecharProgresso
    End If
    sMsg = VBA.Replace(sMsg, "|", "_")
    strArgs = sMsg & "|" & VBA.IIf(sTitulo <> "", sTitulo, "Operação em andamento") & "|" & bShowBntCancelar & "|" & bShowBntFechar & "|" & Progress_Max & "|" & modoProgress & "|" & icone & "|" & intTimer
    
    If Not FormIsOpen(C_FORM_NAME_PROGRESS) Then
        Call DoCmd.openForm(C_FORM_NAME_PROGRESS, acNormal, , , , VBA.IIf(modoProgress = Fim, AcWindowMode.acDialog, AcWindowMode.acWindowNormal), strArgs)
    Else
        Call Form_frmMessage.LoadProgressMensagems(strArgs)
    End If
    
    If FormIsOpen(C_FORM_NAME_PROGRESS) Then Call Forms(C_FORM_NAME_PROGRESS).Repaint
End Sub

Private Function openForm(sFormName As String, Optional bShowModal As Boolean = False) As Form
    If Not FormIsOpen(sFormName) Then
        DoCmd.openForm sFormName, acNormal, , , , VBA.IIf(bShowModal, acDialog, acWindowNormal)
    End If
    On Error Resume Next
    Set openForm = Forms(sFormName)
    VBA.Err.Clear
End Function

Public Sub FecharProgresso()
    Call FecharObjeto(C_FORM_NAME_PROGRESS, acSaveNo)
End Sub

Sub LoadProgressMensagems(fProg As Object)
    Dim bShowDet As Boolean
    fProg.lblTitulo.Caption = VBA.Space(6) & strTitulo
    fProg.lblMensagem.Caption = VBA.Space(3) & strMsg
    bShowDet = strMsgExtendida <> "*"
    fProg.lblDet.Caption = strMsgExtendida
    fProg.lblMacro.Caption = strMacro
    fProg.cmdEntrar.Caption = strTextBotaoEntrar
    fProg.cmdEntrar.Visible = bolBotaoState
    fProg.cmdFechar.Visible = fProg.cmdEntrar.Visible
    fProg.cmdDepurar.Visible = fProg.cmdEntrar.Visible
    fProg.lblMacroFechar.Caption = strRet_action_fechar
    fProg.cmdAcaoExt.Caption = strText_botao_ext
    fProg.lblMacroBotaoExt.Caption = strOnAction_botao_ext
    fProg.lblMacroDepurar.Caption = "DepurarSplash"
    fProg.icoWait.Visible = fProg.lblMensagem.Caption <> ""
    fProg.imgDetalhes.Visible = bShowDet
    fProg.lblDet.Visible = fProg.lblDet.Caption <> "*"
    fProg.cmdAcaoExt.Visible = fProg.lblMacroBotaoExt.Caption <> ""
End Sub


Sub updateProgress(Optional sMsg As String, Optional Form As Object)
    Form!lblInfo.Caption = sMsg
End Sub


Sub openFormSplash(Optional sMsg As String)
    DoCmd.openForm "frmSplash", acNormal, , , , acWindowNormal
End Sub

Sub closeSplash(Optional fName As String = "", Optional bShowInPopup As Boolean = True)
    If bShowInPopup Then
        If fName = "" Then fName = C_FORM_NAME_PROGRESS
    End If
   Call closeForm(fName)
   Call closeForm("frmProgressTimeOut")
End Sub

Sub closeForm(fName As String)
    If fName <> "" Then
        If FormIsOpen(fName) Then DoCmd.Close acForm, fName, acSaveNo
    End If
End Sub

Public Function FormIsOpen(sFormName As String) As Boolean
    On Error Resume Next
    FormIsOpen = CurrentProject.AllForms(sFormName).IsLoaded

    If FormIsOpen Then
        If CurrentProject.AllForms(sFormName).CurrentView = 0 Then
            If VBA.MsgBox("O Formulário está em Modo Estrutura, Deseja fecha-lo agora ?", VBA.vbYesNo + VBA.vbExclamation, "Modo de Design") = VBA.vbYes Then
                Call DoCmd.Close(acForm, sFormName, acSaveNo)
                FormIsOpen = False
            End If
        End If
    End If

End Function


Sub ResetWindowSize(frm As Form)
    Dim intWindowHeight As Integer
    Dim intWindowWidth As Integer
    Dim intTotalFormHeight As Integer
    Dim intTotalFormWidth As Integer
    Dim intHeightHeader As Integer
    Dim intHeightDetail As Integer
    Dim intHeightFooter As Integer

    ' Determine form's height.
    intHeightHeader = frm.Section(acHeader).Height
    intHeightDetail = frm.Section(acDetail).Height
    intHeightFooter = frm.Section(acFooter).Height
    intTotalFormHeight = intHeightHeader _
                       + intHeightDetail + intHeightFooter
    ' Determine form's width.
    intTotalFormWidth = frm.Width

    ' Determine window's height and width.
    intWindowHeight = frm.InsideHeight
    intWindowWidth = frm.InsideWidth

    If intWindowWidth <> intTotalFormWidth Then
        frm.InsideWidth = intTotalFormWidth
    End If
    If intWindowHeight <> intTotalFormHeight Then
        frm.InsideHeight = intTotalFormHeight
    End If
End Sub

Sub FocoTab(Optional sTab As String = "01")
    Call VBA.SendKeys("%Y" & sTab & "%")
End Sub

Function AddAppProperty(strName As String, _
                        varType As Variant, varValue As Variant) As Integer
    Dim dbs As Object, prp As Variant
    Const conPropNotFoundError = 3270

    Set dbs = CurrentDb
    On Error GoTo AddProp_Err
    dbs.Properties(strName) = varValue
    AddAppProperty = True

AddProp_Bye:
    Exit Function

AddProp_Err:
    If VBA.Err = conPropNotFoundError Then
        Set prp = dbs.CreateProperty(strName, varType, varValue)
        dbs.Properties.Append prp
        Resume
    Else
        AddAppProperty = False
        Resume AddProp_Bye
    End If
End Function

Sub RefresProgressBar(Optional fName As String = C_FORM_NAME_PROGRESS, Optional spbName As String, Optional bShowProg As Boolean = True)

    Dim objPB As Object
    Dim lblPB As Object
    Dim lblStatus As Object
    Dim formParent As Form

    On Error Resume Next

    Set formParent = openForm(fName)
    Set objPB = formParent.Controls("pb" & spbName)
    Set lblPB = formParent.Controls("lbl" & spbName)

    'Atualiza o progress
    If (objPB.value + 1) > objPB.Max Then objPB.Max = objPB.value + 1:
    objPB.value = objPB.value + 1
    objPB.Visible = objPB.value > 0 And bShowProg

    If bShowProg Then
        lblPB.Caption = objPB.value & " de " & objPB.Max & " ( " & VBA.Format(objPB.value / objPB.Max, "0% )")
        If Not lblPB.Visible Then lblPB.Visible = 1
        If Not formParent.Controls("lbltime").Visible Then formParent.Controls("lbltime").Visible = 1
        formParent.Controls("lbltime").Caption = VBA.Format(VBA.Now - vTimeProgress, "hh:nn:ss")
    Else
        If lblPB.Visible Then lblPB.Visible = 0
        If formParent.Controls("lbltime").Visible Then formParent.Controls("lbltime").Visible = 0
    End If
    formParent.Repaint
End Sub


Sub TesteProgresso()
Dim i
showMessage "Processando"
showMessage "Processando", , INICIO, 10

For i = 1 To 10
    showMessage "Em andamento", , EmEndamento_MensagemFixa
Next i

showMessage "Processando", , Fim

End Sub
