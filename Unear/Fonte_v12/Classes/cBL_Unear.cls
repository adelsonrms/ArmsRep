VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cBL_Unear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' PROJETO/MODULO   : TFWCliente.cBL_Unear
' TIPO             : Módulo de classe
' DATA/HORA        : 08/08/2017 11:34
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Contem a implementação das funcionalidades de interação com o navegador na pagina da Unear
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão do Módulo
' **************************************************************************************
'   Versão    Data/Hora             Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      08/08/2017 11:34
'---------------------------------------------------------------------------------------
Option Explicit

'Constantes
'------------------------------------------------------------------------------
Private m_cstrURL_Unear As String '= "https://platform.unear.net"
Private Const cstr_Sistema As String = "CRMRede"

'Eventos
'------------------------------------------------------------------------------
Public Event SaveRegistraLog(pMessage As String, pUrl As String)

Public Event Ocupado(bAtivar As Boolean)
Public Event Livre()

'Variáveis
'------------------------------------------------------------------------------
Private m_objBrowser As cIE_Tab
Private IE As cIE_Application
Private WithEvents IEDrive As cIE_Tab
Attribute IEDrive.VB_VarHelpID = -1
Private m_Navegador_StatusAtual As String
Private m_Navegador_URL As String
Private m_CRM_Conectado As Boolean
Private m_iCRM_TelaIndex As Integer
Private cLinks As Object
Private vLink As Object
Private vRetorno As Variant

Private m_sLogFile As String


Public Property Get cstrURL_Unear() As String
1   cstrURL_Unear = m_cstrURL_Unear
End Property

Public Property Let cstrURL_Unear(ByVal vNewValue As String)
1   m_cstrURL_Unear = vNewValue
End Property

'Métodos/Propriedades
'------------------------------------------------------------------------------
Private Function IEApp(Optional pTimeoutSegs As Integer = 10) As cIE_Tab
    Dim dTimeOut As Date
    dTimeOut = VBA.DateAdd("s", pTimeoutSegs, VBA.Now())
    Do While IEDrive Is Nothing
        Call AtualizaInstancia
        If VBA.CDbl(dTimeOut) > 0 Then
            If VBA.Now >= dTimeOut Then
                Debug.Print "WaitFor...Excedeu Timeout !  " & pTimeoutSegs & "Segs"
                Exit Do
            End If
        End If
    Loop
    Set IEApp = IEDrive
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cBL_Unear.AtualizaInstancia()
' TIPO             : Sub
' DATA/HORA        : 08/08/2017 11:29
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Atualiza a instancia do navegador
'---------------------------------------------------------------------------------------
Public Sub AtualizaInstancia()
'Tratamento de Erro
'---------------------------------------------------------------------------------------
1   On Error GoTo AtualizaInstancia_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String:
    Const cstr_ProcedureName As String = "cBL_Unear.AtualizaInstancia()"
    '---------------------------------------------------------------------------------------

2   Set IE = New cIE_Application
    
'    If IE.BroswerIsOpen Then
    
3   Set IEDrive = IE.NovaTab("", ByFirstIntance)
4   Set m_objBrowser = IEDrive
5   If Not m_objBrowser Is Nothing Then
6       m_Navegador_StatusAtual = "Aberto"
7       m_Navegador_URL = m_objBrowser.Url
8       RaiseEvent SaveRegistraLog(Now & " Navegador ja inicializado !", m_objBrowser.Url)
9   Else
10      m_Navegador_StatusAtual = "Fechado"
11      m_Navegador_URL = ""
12  End If
    
'    End If
Fim:
13  On Error GoTo 0
14  Exit Sub

AtualizaInstancia_Error:
15  If Err <> 0 Then
16      lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
17      Call VBA.Err.Raise(VBA.Err.Number, cstr_ProcedureName, VBA.vbNewLine & cstr_ProcedureName & " Em" & VBA.IIf(VBA.Erl() <> 0, " > Linha (" & VBA.Erl() & ")", "") & " > " & VBA.Err.Description)
18  End If
    GoTo Fim:
    'Debug Mode
19  Resume
End Sub

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cBL_Unear.AcessarSite()
' TIPO             : Function
' DATA/HORA        : 08/08/2017 11:00
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Abre o Browser e acessa o site da Unear
'---------------------------------------------------------------------------------------
Function AcessarSite()
    Dim bRetorno As Boolean
    'Se nao tiver nenhuma tab aberta, abre uma nova instancia
    'Tratamento de Erro
    '---------------------------------------------------------------------------------------
1   On Error GoTo AcessarSite_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String:
    Const cstr_ProcedureName As String = "cBL_Unear.AcessarSite()"
    '---------------------------------------------------------------------------------------

2   If IE.Tabs.count > 0 Then
3       Set IEDrive = IE.PegarTab(ModoInstancia:=ByFirstIntance).item(1)
        
        'Urls validas da plataforma
        'Nova URL : https://platform-itau.ucampaign.unear.net
4       If IEDrive.Url Like cstrURL_Unear & "*" Or _
           IEDrive.Url Like "*app.unear.net*" Or _
           IEDrive.Url Like "*platform-itau.ucampaign.unear.net*" Or _
           IEDrive.Url Like "*app.ucampaign.unear.net*" Then
           
5           bRetorno = True
6           GoTo Fim
7       Else
8           Set IEDrive = IE.NovaTab(pUrl:=cstrURL_Unear, ModoInstancia:=ByFirstIntance)
9           bRetorno = True
10      End If

11  Else
12      Set IEDrive = IE.NovaTab(pUrl:=cstrURL_Unear, ModoInstancia:=NewBrowserInstance)
13      bRetorno = True
14  End If
Fim:
15  If bRetorno Then
        IE.ShowWindow
17  Else
18      AcessarSite = Array(VBA.Err, VBA.Error)
19  End If
20  On Error GoTo 0
21  Exit Function

AcessarSite_Error:
22  If VBA.Err <> 0 Then
23      lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
24      Call VBA.Err.Raise(VBA.Err.Number, cstr_ProcedureName, VBA.vbNewLine & cstr_ProcedureName & " Em" & VBA.IIf(VBA.Erl() <> 0, " > Linha (" & VBA.Erl() & ")", "") & " > " & VBA.Err.Description)
25  End If
    GoTo Fim:
26  Resume    'Debug Mode

End Function

Sub AguardaTelaLogin()
    On Error GoTo Entrar_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String:
    Const cstr_ProcedureName As String = "cBL_Unear.Entrar()"
    '---------------------------------------------------------------------------------------
    IEApp
    
    Call EventoOcupado(False)
    If Not Me.CRM_Logado Then
        IEApp.WaitFor "username", STOP_WHEN_FIND_ID
        Do While Not IEApp.ControlExists("username") And Not IEApp.ControlExists("username")
            If IEApp.PaginaComErro() Then Exit Do
            Call EventoOcupado(True)
            If bCancelar Then Exit Sub
            If Me.CRM_Logado Then Exit Do
            VBA.DoEvents
        Loop
        Call EventoOcupado(bRodando)
    End If
Fim:
    Exit Sub
Entrar_Error:
36  If VBA.Err <> 0 Then
37      lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
38      Call VBA.Err.Raise(VBA.Err.Number, cstr_ProcedureName, VBA.vbNewLine & cstr_ProcedureName & " Em" & VBA.IIf(VBA.Erl() <> 0, " > Linha (" & VBA.Erl() & ")", "") & " > " & VBA.Err.Description)
39  End If
    GoTo Fim:
40  Resume
End Sub

Sub EventoOcupado(bAtivar As Boolean)
    bRodando = bAtivar
    RaiseEvent Ocupado(bAtivar)
End Sub

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cBL_Unear.Entrar()
' TIPO             : Function
' DATA/HORA        : 08/08/2017 11:07
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Efetua logon no site caso ainda não tenha feito
'---------------------------------------------------------------------------------------
Function Entrar(pUsuario As String, pSenha As String)
'Tratamento de Erro
'---------------------------------------------------------------------------------------
    On Error GoTo Entrar_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String:
    Const cstr_ProcedureName As String = "cBL_Unear.Entrar()"
    '---------------------------------------------------------------------------------------
    Dim strMensagem As String
    Dim div As Object

2   Call auxAuditoria.RegistraLog("Entrar() - Entrando no sistema")
3   Me.AtualizaInstancia
4   IEApp.WaitFor

5   If Not Me.CRM_Logado Then

6       If Me.CRM_TelaIndex = 1 Then
7           Call IEApp.WaitFor("username", STOP_WHEN_FIND_ID)
8           IEApp.Control("username").value = pUsuario    '"crm@userede.com.br"
9           IEApp.Control("password").value = pSenha    '"CRM*campanhas01"
10          Call auxAuditoria.RegistraLog("Entrar() - Usuario e senha preenchidos...Submit")
11          IEApp.Document.Forms("loginForm").submit
12          Call IEApp.WaitFor("topleftcontent", STOP_WHEN_FIND_ID)
13          Call auxAuditoria.RegistraLog("Entrar() - Aguardando o login no sistema")
14          strMensagem = "OK ! Logon com sucesso"
15      Else
16          strMensagem = "Ja está logado em outra tela (" & Me.CRM_TelaAtual & ")"
17      End If
18      Me.AtualizaInstancia

    'Verifica se houve erro no logon atraves da mensagem de erro reportado pelo site
23  Set div = IEApp.GetElementsByTag(IEApp.Document, "div", Array("ng-messages", "messageResult"))
24  If div.count > 0 Then
25      strMensagem = div.item(1).element.innertext
26      VBA.MsgBox "Erro durante o logon !" & div.item(1).element.innertext, vbExclamation
27  End If

19  Else
20      strMensagem = "Ja está logado em outra tela (" & Me.CRM_TelaAtual & ")"
21  End If

    'Aguarda até que o ID 'topleftcontent' esperado na tela seguinte seja encontrado
22  Call IEApp.WaitFor("topleftcontent", STOP_WHEN_FIND_ID)

Fim:
28  If IEApp.ControlExists("topleftcontent") Then
29      Entrar = Array(1, strMensagem)
30  Else
31      Entrar = Array(2, strMensagem)
32  End If
33  Call auxAuditoria.RegistraLog("Entrar() - " & strMensagem)

34  On Error GoTo 0
35  Exit Function


Entrar_Error:
36  If VBA.Err <> 0 Then
        Stop
37      lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
38      Call VBA.Err.Raise(VBA.Err.Number, cstr_ProcedureName, VBA.vbNewLine & cstr_ProcedureName & " Em" & VBA.IIf(VBA.Erl() <> 0, " > Linha (" & VBA.Erl() & ")", "") & " > " & VBA.Err.Description)
39  End If
    GoTo Fim:
40  Resume    'Debug Mode
End Function

Private Sub CarregaListaDeCampanhas()
2   Call IEApp.WaitFor("topleftcontent", STOP_WHEN_FIND_ID)
    'Localiza o link CRMRede
3   Set vLink = PegaLinkSistema()
4   Call IEApp.Click(vLink.element)
    'Aguarda o final do carregamento da pagina
5   Call IEApp.WaitFor("campaignList", STOP_WHEN_FIND_ID)
End Sub

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cBL_Unear.AcessarListaDeCampanhas()
' TIPO             : Function
' DATA/HORA        : 08/08/2017 11:08
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Acessa a tela com a lista de campanhas para o perfil CRMRede
'---------------------------------------------------------------------------------------
Public Function AcessarListaDeCampanhas()
    Dim strMensagem As String
    
    'Tratamento de Erro
    '---------------------------------------------------------------------------------------
1   On Error GoTo AcessarListaDeCampanhas_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String:
    Const cstr_ProcedureName As String = "cBL_Unear.AcessarListaDeCampanhas()"
    '---------------------------------------------------------------------------------------
    If PerfilSelecionado <> cstr_Sistema Or CRM_TelaIndex <> 2 Then
    
    Call CarregaListaDeCampanhas
    Call Me.AguardaCarregamentoTabela
    
    End If
Fim:
7   If CRM_TelaIndex = 2 Then
8       strMensagem = "Lista de Campanhas OK !"
9       AcessarListaDeCampanhas = Array(1, strMensagem)
10  Else
11      AcessarListaDeCampanhas = Array(2, strMensagem)
12  End If

13  On Error GoTo 0
14  Exit Function

AcessarListaDeCampanhas_Error:
15  If VBA.Err <> 0 Then
16      lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
17      Call VBA.Err.Raise(VBA.Err.Number, cstr_ProcedureName, VBA.vbNewLine & cstr_ProcedureName & " Em" & VBA.IIf(VBA.Erl() <> 0, " > Linha (" & VBA.Erl() & ")", "") & " > " & VBA.Err.Description)
18  End If
    GoTo Fim:
19  Resume    'Debug Mode
End Function

Function AguardaCarregamentoTabela()
    On Error GoTo Error_CarregarTela
    Dim Tabela, tBody
    
Inicia:
    'Aguarda até que o ID da tabela esteja disponivel.
    'Caso não esteja, a tabela não foi atualizada
    Call IEApp.WaitFor("bsDataTable", STOP_WHEN_FIND_ID, , 20)
    If IEApp.ControlExists("bsDataTable") Then
        Set Tabela = IEApp.Control("bsDataTable")
        Set tBody = Tabela.tBodies.item(0)
        'Aguarda o carregamento da tabela
        Do While tBody.Rows(0).Cells.Length < 2
            Set Tabela = IEApp.Control("bsDataTable")
            Set tBody = Tabela.tBodies.item(0)
            'Debug.Print tBody.innertext
            If tBody.innertext Like "Nenhum registro encontrado" Then
                If tBody.Rows(0).Cells.Length < 2 Then
                AcessarListaDeCampanhas
                End If
                GoTo Inicia
            End If
            VBA.DoEvents
        Loop
        vRetorno = VBA.Array(1, "Lista não atualizada")
    Else
        vRetorno = VBA.Array(0, "Lista não atualizada")
    End If
Fim:
    AguardaCarregamentoTabela = vRetorno
    Exit Function
    
Error_CarregarTela:
    vRetorno = VBA.Array(0, "AguardaCarregamentoTabela() Linha (" & VBA.Erl() & ") - " & VBA.Error())
    Debug.Print vRetorno(1)
    GoTo Fim
End Function

'Padrão 'CRMRede'
'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cBL_Unear.PegaLinkSistema()
' TIPO             : Function
' DATA/HORA        : 08/08/2017 11:11
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Localiza o link CRMRede na opção de seleção do perfil
'---------------------------------------------------------------------------------------
Private Function PegaLinkSistema()
'Tratamento de Erro
'---------------------------------------------------------------------------------------
1   On Error GoTo PegaLinkSistema_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String:
    Const cstr_ProcedureName As String = "cBL_Unear.PegaLinkSistema()"
    '---------------------------------------------------------------------------------------
    Dim dTimeOut As Date
2   dTimeOut = PreparaTimeOut()

PegaLink:
3   Set cLinks = IEApp.GetElementsByTag(IEApp.Document.getElementById("topleftcontent"), "a", Array("href", "#", "data-value", 50))
4   For Each vLink In cLinks
5       If VBA.Trim(vLink.value) = cstr_Sistema Then Exit For
6   Next vLink

7   If vLink Is Nothing Then
8       If DeuTimeOut(dTimeOut) Then GoTo Fim
9       GoTo PegaLink
10  End If
Fim:
11  Set PegaLinkSistema = vLink

12  On Error GoTo 0
13  Exit Function

PegaLinkSistema_Error:
14  If Err <> 0 Then
15      lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
16      Call VBA.Err.Raise(VBA.Err.Number, cstr_ProcedureName, VBA.vbNewLine & cstr_ProcedureName & " Em" & VBA.IIf(VBA.Erl() <> 0, " > Linha (" & VBA.Erl() & ")", "") & " > " & VBA.Err.Description)
17  End If
    GoTo Fim:
18  Resume    'Debug Mode
End Function

Function PegarElementoPorClass(tag As String, valor_atributo As String, Optional atributo As String = "class", Optional oParent As Object)
    Set cLinks = New VBA.Collection
    Set cLinks = IEApp.GetElementsByTag(oParent, tag, Array(atributo, valor_atributo))
16  Do While cLinks.count = 0
17      Set cLinks = IEApp.GetElementsByTag(oParent, tag, Array(atributo, valor_atributo))
18      If cLinks.count > 0 Then Exit Do
19  Loop
    Set PegarElementoPorClass = cLinks.item(1)
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cBL_Unear.Sair()
' TIPO             : Function
' DATA/HORA        : 08/08/2017 11:29
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Sai do sistema
'---------------------------------------------------------------------------------------
Function Sair()
'Tratamento de Erro
'---------------------------------------------------------------------------------------
1   On Error GoTo Sair_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String:
    Const cstr_ProcedureName As String = "cBL_Unear.Sair()"
    '---------------------------------------------------------------------------------------
    
    If CRM_Logado Then
    
    Call IEApp.WaitFor("headerContent", STOP_WHEN_FIND_ID)
    
    
    Set cLinks = IEApp.GetElementsByTag(IEApp.Document.getElementById("headerContent"), "a", Array("href", "#ucampaignLogout"))
  Do While cLinks.count = 0
      Set cLinks = IEApp.GetElementsByTag(IEApp.Document.getElementById("headerContent"), "a", Array("href", "#ucampaignLogout"))
      If cLinks.count > 0 Then Exit Do
  Loop

3   For Each vLink In cLinks
4       If VBA.Trim(vLink.value) = "Sair" Then Call IEApp.Click(vLink)
5   Next vLink

        Call AguardaTelaLogin
    End If

Fim:
6   On Error GoTo 0
7   Exit Function

Sair_Error:
8   If VBA.Err <> 0 Then
        Stop
9       lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
10      Call VBA.Err.Raise(VBA.Err.Number, cstr_ProcedureName, VBA.vbNewLine & cstr_ProcedureName & " Em" & VBA.IIf(VBA.Erl() <> 0, " > Linha (" & VBA.Erl() & ")", "") & " > " & VBA.Err.Description)
11  End If
    GoTo Fim:
12  Resume    'Debug Mode
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cBL_Unear.AcessarListaDeSelecoes()
' TIPO             : Function
' DATA/HORA        : 08/08/2017 11:31
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Seleciona a lista de Seleções acionando o clique no meno Opções > Seleções
'---------------------------------------------------------------------------------------
Function AcessarListaDeSelecoes()

'Tratamento de Erro
'---------------------------------------------------------------------------------------
1   On Error GoTo AcessarListaDeSelecoes_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String:
    Const cstr_ProcedureName As String = "cBL_Unear.AcessarListaDeSelecoes()"
    '---------------------------------------------------------------------------------------
    Dim Config_pX, Config_pY, Menu_pX, Menu_pY
    Dim vPonteiro
    
    Call showMessage("Acessa a lista de Seleções", "Unear", SemProgresso, , True)
    
2   Call Inicializar_Globais

3   IEApp.Application.Visible = 1
    
4   vPonteiro = Conexao.PegarArray("Pegar_PonteiroMouse", "MenuOpcoes", VBA.Environ("ComputerName"))
5   Config_pX = vPonteiro(0, 0)
6   Config_pY = vPonteiro(1, 0)

7   vPonteiro = Conexao.PegarArray("Pegar_PonteiroMouse", "MenuSelecoes", VBA.Environ("ComputerName"))
8   Menu_pX = vPonteiro(0, 0)
9   Menu_pY = vPonteiro(1, 0)

ClicaMenu:
10  Call ClickPosicaoMouse(Array(Config_pX, Config_pY), Array(Menu_pX, Menu_pY))
11  Do While Not IEApp.ControlExists("selectionList")
12      VBA.DoEvents
13      If CRM_TelaIndex = 2 Then GoTo ClicaMenu
14  Loop

15  If CRM_TelaIndex = 3 Then
16      AcessarListaDeSelecoes = Array(1, "OK")
17  Else
18      AcessarListaDeSelecoes = Array(2, "Erro")
19  End If

Fim:
20  On Error GoTo 0
21  Exit Function

AcessarListaDeSelecoes_Error:
22  If Err <> 0 Then
23      lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
24      Call VBA.Err.Raise(VBA.Err.Number, cstr_ProcedureName, VBA.vbNewLine & cstr_ProcedureName & " Em" & VBA.IIf(VBA.Erl() <> 0, " > Linha (" & VBA.Erl() & ")", "") & " > " & VBA.Err.Description)
25  End If
    GoTo Fim:
26  Resume    'Debug Mode
End Function

Function ClicarMouseNaPosicao(strCodPosicao As String)
    Dim arrPointer As Variant
    arrPointer = BuscaInfoPonteito(strCodPosicao):
    If arrPointer(0) <> 0 Then
        Call ClickPosicaoMouse(arrPointer)
        ClicarMouseNaPosicao = Array(True, "OK, '" & strCodPosicao & "' foi clicado")
    Else
        ClicarMouseNaPosicao = Array(False, "Não foi possivel enviar o clique na configuração '" & strCodPosicao & "'")
    End If
End Function
'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cBL_Unear.ClickPosicaoMouse()
' TIPO             : Function
' DATA/HORA        : 08/08/2017 11:37
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Envia o clique do mouse em uma posição especifca no navegador
'---------------------------------------------------------------------------------------
Private Function ClickPosicaoMouse(ParamArray Posicao() As Variant)
    Dim oWindow As cTFW_Window
    'Tratamento de Erro
    '---------------------------------------------------------------------------------------
1   On Error GoTo ClickPosicaoMouse_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String:
    Const cstr_ProcedureName As String = "cBL_Unear.ClickPosicaoMouse()"
    '---------------------------------------------------------------------------------------
    Dim iPos As Integer

2   Set oWindow = New cTFW_Window
3   Call oWindow.GetByClassName(pClassName:="IEFrame")
4   oWindow.Activate
5   Call oWindow.Wait(500)
6   For iPos = 0 To UBound(Posicao)
7       Call oWindow.MouseClick(VBA.CLng(Posicao(iPos)(0)), VBA.CLng(Posicao(iPos)(1)), 100)
        Call oWindow.Wait(2000)
8   Next iPos
9   Call oWindow.Wait(1000)

Fim:
10  On Error GoTo 0
11  Exit Function

ClickPosicaoMouse_Error:
12  If VBA.Err <> 0 Then
13      lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
14      Call VBA.Err.Raise(VBA.Err.Number, cstr_ProcedureName, VBA.vbNewLine & cstr_ProcedureName & " Em" & VBA.IIf(VBA.Erl() <> 0, " > Linha (" & VBA.Erl() & ")", "") & " > " & VBA.Err.Description)
15  End If
    GoTo Fim:
16  Resume    'Debug Mode
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cBL_Unear.UploadSelecao()
' TIPO             : Function
' DATA/HORA        : 08/08/2017 11:26
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Efetua o upload da lista de usuários na ferramenta
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      08/08/2017 11:26
'---------------------------------------------------------------------------------------
Function UploadSelecao(strLocalArquivo As String, strNome As String)
'Tratamento de Erro
'---------------------------------------------------------------------------------------
1   On Error GoTo UploadSelecao_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String:
    Const cstr_ProcedureName As String = "cBL_Unear.UploadSelecao()"
    '---------------------------------------------------------------------------------------

    Dim bSalvou As Boolean
    Dim vPonteiro
    Dim px, py
    Dim strSelecao As String

2   If Not IEApp.ControlExists("selectionList") Then VBA.MsgBox "Necessário esta na tela de Seleções !", vbExclamation
    
    'Após clicar no link de seleções, Localiza o Link de ID=11 que corresponde ao botão para Upload
5   If IEApp.ControlExists("11") Then: Call IEApp.Click(IEApp.Control("11"))

    'Aguarda o carregamento da pagina
8   Call IEApp.WaitFor("fileupload", STOP_WHEN_FIND_ID)

ClicaBotao:
    vRetorno = ClicarMouseNaPosicao("BotaoUploadSelecao")

    'Após clicar, monitora a caixa de dialogo
13  bSalvou = UploadListaEmailsTXT(strLocalArquivo)

    If Not bSalvou Then GoTo ClicaBotao:

    'Aguarda a seleção do arquivo
14  Call IEApp.WaitFor("basicInfo", STOP_WHEN_FIND_ID)
    Call win.Wait(1000)
    
15  Set vLink = IEApp.GetElementsByTag(IEApp.Document.getElementById("basicInfo"), "input", Array("name", "name"))
16  Do While vLink.count = 0
17      Set vLink = IEApp.GetElementsByTag(IEApp.Document.getElementById("basicInfo"), "input", Array("name", "name"))
18      If vLink.count > 0 Then Exit Do
19  Loop
20  If vLink.count > 0 Then vLink.item(1).value = strNome
    
    IEApp.WaitFor "htmlResume"
    Do While (Not IEApp.Control("uploadResult").innertext Like "Seleção gerada com sucesso*") And (Not IEApp.Control("uploadResult").innertext Like "Ops, ocorreu um erro, entre em contato com nossa equipe de suporte.*")
21      Call IEApp.Click(IEApp.Control("btnSave"))
22      Call IEApp.WaitFor("uploadResult")
        Debug.Print "Aguardando upload da seleção : " & IEApp.Control("uploadResult").innertext
        win.Wait
        If IEApp.Control("uploadResult").innertext Like "Ops, ocorreu um erro, entre em contato com nossa equipe de suporte.*" Then
            Exit Do
            'GoTo ClicaBotao:
        End If
        VBA.DoEvents
    Loop
    
    Call IEApp.WaitFor("Group_1", STOP_WHEN_FIND_ID)
    Call IEApp.Click(IEApp.Control("Group_1").Children(1))
    
    Call win.Wait(1000)
    UploadSelecao = PegaNumeroSelecao(strNome)

Fim:
24  On Error GoTo 0
25  Exit Function

UploadSelecao_Error:
26  If VBA.Err <> 0 Then
27      lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
        MsgBox strErrorMessagem
28      Call VBA.Err.Raise(VBA.Err.Number, cstr_ProcedureName, VBA.vbNewLine & cstr_ProcedureName & " Em" & VBA.IIf(VBA.Erl() <> 0, " > Linha (" & VBA.Erl() & ")", "") & " > " & VBA.Err.Description)
29  End If
    GoTo Fim:
    'Debug Mode
30  Resume
End Function
'Seleciona a opção de menu 'Selecoes do Usuario'
'Retorna o nomero (codigo) da uma Seleção atraves do nome
'Clica na opção '"Group_1"' para garantir que a lista esteja disponivel
Function PegaNumeroSelecao(NomeSelecao As String) As Variant
    On Error GoTo Erro
    Dim NumPaginacao As Object
    
    'Do While PegaNumeroSelecao = ""
'    Call AcessarListaDeSelecoes
'    Call IEApp.WaitFor("Group_1", STOP_WHEN_FIND_ID)
'    Call IEApp.Click(IEApp.Control("Group_1").Children(1))
    'Selecione os 25 items para forçar a atualização da tela
    Set NumPaginacao = PegarElementoPorClass("span", "btn-group dropup")
    Do While NumPaginacao Is Nothing
        If Not NumPaginacao Is Nothing Then
            NumPaginacao.element.Children(1).Children(2).Children(0).Click
            Exit Do
        End If
        Set NumPaginacao = PegarElementoPorClass("span", "btn-group dropup")
    Loop
    
    Call IEApp.WaitFor("bsDataTable", STOP_WHEN_FIND_ID)
    PegaNumeroSelecao = PegaCodigoPorNome(NomeSelecao)
    'Loop
    
    Exit Function
Erro:
    MsgBox VBA.Error
    Resume
End Function
'Recupera o numero de uma Campanha atraves do nome
Function PegaNumeroCampanha(strNomeCampanha As String) As Variant
    On Error GoTo Err_PegarNumCampanha
    Dim strSQLDinamic As String, strTableName As String
    Dim rs As Object
    If Not Me.CRM_TelaIndex = 2 Then AcessarListaDeCampanhas
    PegaNumeroCampanha = PegaCodigoPorNome(strNomeCampanha)
    Exit Function
Err_PegarNumCampanha:
    PegaNumeroCampanha = ""
End Function
'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cBL_Unear.CriarCampanha()
' TIPO             : Function
' DATA/HORA        : 08/08/2017 11:19
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Cria uma campanha no sistema
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      08/08/2017 11:19
'---------------------------------------------------------------------------------------
Function CriarCampanha(strNumeroSelecaoUsuario As String, strNome As String, Optional strAssunto As String, Optional ModeloEmail)
          'Tratamento de Erro
          '---------------------------------------------------------------------------------------
10        On Error GoTo CriarCampanha_Error
20        Dim lngErrorNumber As Long, strErrorMessagem As String:
          Const cstr_ProcedureName As String = "cBL_Unear.CriarCampanha()"
          '---------------------------------------------------------------------------------------
          Dim LinkOpcoes As Object
          Dim TipoCampanha As String
          Dim strNumeroCampanha As String, strMensagemCampanha As String, bStatus As Boolean
          
          Dim px As Long, py As Long, NumPaginacao As Object
          Dim w As New cTFW_Window
          
          Call EnviaLogStatus(VBA.String(120, ".")): Call EnviaLogStatus("INICIANDO > " & cstr_ProcedureName): Call EnviaLogStatus(VBA.String(120, "."))
          
40        Call AcessarListaDeCampanhas
          
50        TipoCampanha = pegaValor("tipoCampanha")
60        Call EnviaLogStatus("TIPO DE CAMPANHA : " & TipoCampanha)

IniciaCriacao:
          'PASSO 1 - Uma vez na tela de campanha, clica no botão "+" (ID = 7)
          '-----------------------------------------------------------------------------------------
70        Call EnviaLogStatus("Cria nova campanha (Botão +)")
80        Call IEApp.WaitFor("7", STOP_WHEN_FIND_ID)
90        Call IEApp.Click(IEApp.Control("7"))

          'Tipo de Campanha
100       If VBA.LCase(TipoCampanha) = "email" Then
110           Call EnviaLogStatus("Opção de Email já marcada")
120           IEApp.WaitFor "emailCamp", STOP_WHEN_FIND_ID
130           IEApp.Control("emailCamp").Checked = True
140           IEApp.Control("hfCampType").value = "email"
150       Else
160           Call EnviaLogStatus("Opção de SMS - Envia o clique na posição SMS pre-configurada (OpcaoSMS)")
170           win.Wait 1000
180           Call IEApp.Click(IEApp.Control("lkCampType"))
              'No caso da opção SMS, seleciona via click
190           vRetorno = ClicarMouseNaPosicao("OpcaoSMS")
200           If vRetorno(0) <> True Then
210               Call EnviaLogStatus("Erro ao selecionar a opção SMS. Clique não funcionou. Criação da campanha cancelada")
220               strMensagemCampanha = vRetorno(1)
230               Call EnviaLogStatus("ClicarMouseNaPosicao () - " & strMensagemCampanha)
240               GoTo Fim
250           Else
260               Call EnviaLogStatus("OK - SMS foi clicado")
270           End If
280       End If
        
290       Call EnviaLogStatus("PASSO 2 - Preenche as Informações Basicas (Nome : " & strNome & ")")
          'PASSO 2 - Clica em 'lkBasicInfo'
          '-----------------------------------------------------------------------------------------
300       Call IEApp.WaitFor("lkBasicInfo", STOP_WHEN_FIND_ID)
310       Call IEApp.Click(IEApp.Control("lkBasicInfo"))

          'PASSO 3 - Preenche o nome
          '-----------------------------------------------------------------------------------------
320       Set LinkOpcoes = IEApp.GetElementsByTag(IEApp.Document.getElementById("campBasicInfo"), "input", Array("name", "name"))
330       If LinkOpcoes.count > 0 Then LinkOpcoes.item(1).element.value = strNome
          
340       Call EnviaLogStatus("PASSO 3 - Preenche a seleção do usuário (Codigo : " & strNumeroSelecaoUsuario & ")")
          'PASSO 4 - Aba de 'Seleção de Usuarios'
          '-----------------------------------------------------------------------------------------
350       Call IEApp.Click(IEApp.Control("pUsesrSelection").Children(0))
360       IEApp.Control("usersSel").value = strNumeroSelecaoUsuario
          
370       Call EnviaLogStatus("PASSO 4 - Preenche o modelo a ser utilizado")
380       If TipoCampanha = "SMS" Then
390           Call EnviaLogStatus("PASSO 4 - Opção SMS, apenas salva o conteudo na caixa de texto (SmsTextMessageId_0)")
              'PASSO 5 - Modelo de SMS
              '-----------------------------------------------------------------------------------------
400           Call IEApp.Click(IEApp.Control("pSmsPiece").Children(0))
410           win.Wait 2000
420           If Not IEApp.ControlExists("SmsTextMessageId_0") Then
430               strMensagemCampanha = "SMS - Erro na criação da campanha SMS. Possivelmente Não selecionou a opção SMS. Caixa de mensagem SmsTextMessageId_0 não localiza"
440               Call EnviaLogStatus("PASSO 4 - " & strMensagemCampanha)
450               Call EnviaLogStatus("PASSO 4 - ERROR : " & strMensagemCampanha)
460               Call EnviaLogStatus("PASSO 4 - Cancelando campanha...")
470               bStatus = False
480               GoTo Fim
490           Else
500               IEApp.Control("SmsTextMessageId_0").value = strAssunto
510               Call EnviaLogStatus("PASSO 4 - OK, mensagem SMS salva")
520           End If
              'PASSO FINAL SMS - Salva a campanha
              '-----------------------------------------------------------------------------------------
530           IEApp.Click IEApp.Control("btnSaveCamp")
540           Call EnviaLogStatus("PASSO 5 - Campanha SMS Criada com sucesso !")
550       Else
560           Call EnviaLogStatus("PASSO 5 - Opção de Email. Prepara parametros...")
              'PASSO 5 - Modelo de Email
              '-----------------------------------------------------------------------------------------
570           Call IEApp.Click(IEApp.Control("pCampaignPiece").Children(0))
580           Call IEApp.WaitFor("subject", STOP_WHEN_FIND_ID)
590           Call w.Wait(500) 'Aguarda meio segundo
        
600           Call EnviaLogStatus("PASSO 6 - Clica na posição de seleção do modelo")
              'Edita a caixa de pesquisa dos modelos
              '---------------------------------------------------------------------------------------------
610           vRetorno = ClicarMouseNaPosicao("SelecaoModeloEmail_Editar")
620           win.Wait
              
630           On Error Resume Next
640           Set vLink = IEApp.GetElementsByTag(IEApp.Document.getElementById("select2-drop"), "li", Array("class", "select2-no-results"))
650           Do While vLink.count = 0
660               Set vLink = IEApp.GetElementsByTag(IEApp.Document.getElementById("select2-drop"), "li", Array("class", "select2-no-results"))
670               If vLink.count > 0 Then Exit Do
680           Loop
690           If vLink.count > 0 Then
700               Call EnviaLogStatus("PASSO 7 - Envia o nome do modelo (" & ModeloEmail & ") na caixa de pesquisa")
710               Call EnviarTexto(VBA.CStr(ModeloEmail))
720           End If
            
730           On Error GoTo CriarCampanha_Error
ClicaModeloSelecionado:
740           Call EnviaLogStatus("PASSO 8 - Após o sistema localizar o modelo, envia o clique na posição para seleção do primeiro item")
750           vRetorno = ClicarMouseNaPosicao("SelecaoModeloEmail_Selecionar")
        
760           If vRetorno(0) = True Then
770               If Not VBA.Replace(Trim(IEApp.Control("s2id_campTemp").innertext), vbNewLine, "") Like "*" & VBA.CStr(ModeloEmail) & "*" Then
780                   Call EnviaLogStatus("PASSO 9 - Modelo Não Selecioado, Tenta nova seleção...")
790                   If IEApp.Control("subject").value = VBA.CStr(ModeloEmail) And IEApp.ControlExists("charlimitinfo") Then
                          'Cancela e cria novamente
800                       If IEApp.Control("charlimitinfo").innertext Like "*caracteres restantes*" Then
810                           Debug.Print "Preenchimento incorreto no assunto. Recria a campanha..."
820                           win.Wait
830                           GoTo IniciaCriacao
840                       End If
850                   End If
860                   GoTo ClicaModeloSelecionado
870               Else
880                   Call EnviaLogStatus("PASSO 9 - Modelo Selecioado com sucesso")
890               End If
900           Else
910               strMensagemCampanha = "Não foi selecionado o clique para seleção do Modelo de Email"
920               GoTo Fim
930           End If
              
940           Call EnviaLogStatus("PASSO 10 - Modelo OK, Preenche o assunto e finaliza a criação...")
              'Modelo carregado, preenche os dados
              '---------------------------------------------------------------------------------------------
950           IEApp.WaitFor "subject", STOP_WHEN_FIND_ID
960           IEApp.Control("subject").value = strAssunto
            
970           Call EnviaLogStatus("PASSO 11 - Aguarda o carregamento do HTML do modelo...")
980           If IEApp.ControlExists("designPane_0") Then
990               Set LinkOpcoes = IEApp.GetElementsByTag(IEApp.Document.getElementById("designPane_0"), "textarea", Array("class", "note-codable"))
1000              If LinkOpcoes.count > 0 Then
1010                  If LinkOpcoes(1).element.innertext <> "" Then
1020                      If IEApp.ControlExists("subject") Then
1030                          IEApp.Control("subject").value = strAssunto
1040                      Else
1050                          strMensagemCampanha = "O assunto do email não preenchido corretamente"
1060                          GoTo Fim
1070                      End If
1080                      Call EnviaLogStatus("PASSO 12 - Tudo certo !, finaliza a criação")
                          'PASSO FINAL - Salva a campanha
                          '-----------------------------------------------------------------------------------------
1090                      IEApp.Click IEApp.Control("btnSaveCamp")
1100                      Call EnviaLogStatus("PASSO 5 - Campanha EMAIL Criada com sucesso !")
                          bStatus = True
1110                  End If
1120              Else
1130                  strMensagemCampanha = "Não carregou as informações do modelo selecionado"
1140                  GoTo Fim
1150              End If
1160          End If
1170      End If

          If bStatus Then
              Call EnviaLogStatus("PASSO FINAL - Localizando o numero da campanha gerada")
             'Busca o numero da campanha criada
1180         strNumeroCampanha = PegaNumeroCampanha(strNome)
          End If

1190      If strNumeroCampanha <> "" Then
1200          bStatus = True
1210          strMensagemCampanha = "Campanha '" & strNumeroCampanha & "' criada com sucesso. Pronta para teste/execução"
              Call EnviaLogStatus(strMensagemCampanha)
1220      Else
1230          bStatus = False
1240          strMensagemCampanha = "Numero de Campanha não gerado. Ocorreu algum erro"
              Call EnviaLogStatus(strMensagemCampanha)
1250      End If
Fim:
1260      CriarCampanha = VBA.Array(bStatus, strNumeroCampanha, strMensagemCampanha)
          
          Call EnviaLogStatus(cstr_ProcedureName & " - " & strMensagemCampanha)
          
1270      On Error GoTo 0
          Call EnviaLogStatus(VBA.String(120, ".")): Call EnviaLogStatus("INICIANDO > " & cstr_ProcedureName): Call EnviaLogStatus(VBA.String(120, "."))
          
1280      Exit Function

CriarCampanha_Error:
1290      If VBA.Err <> 0 Then
1300          lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
              Call EnviaLogStatus(cstr_ProcedureName & " - ERROR > " & " Em" & VBA.IIf(VBA.Erl() <> 0, " > Linha (" & VBA.Erl() & ")", ""))
1320          Call VBA.Err.Raise(VBA.Err.Number, cstr_ProcedureName, VBA.vbNewLine & cstr_ProcedureName & " Em" & VBA.IIf(VBA.Erl() <> 0, " > Linha (" & VBA.Erl() & ")", "") & " > " & VBA.Err.Description)
1330      End If
1340      GoTo Fim:
          'Debug Mode
1350      Resume
End Function

Public Property Get LogFile() As String
    LogFile = m_sLogFile
End Property

Public Property Let LogFile(ByVal sNewValue As String)
    m_sLogFile = sNewValue
End Property

Sub EnviarTexto(sTexto As String)
Dim oWShell As Object
Set oWShell = VBA.CreateObject("wscript.shell")
    oWShell.SendKeys sTexto
End Sub

Function BuscaInfoPonteito(strInfoPonteiro As String)
    Dim px As Long, py As Long
    Dim vPonteiro
    Inicializar_Globais
9   vPonteiro = Conexao.PegarArray("Pegar_PonteiroMouse", strInfoPonteiro, PerfilMousePointer())
    If Not VBA.IsEmpty(vPonteiro) Then
10      px = vPonteiro(0, 0)
11      py = vPonteiro(1, 0)
        BuscaInfoPonteito = VBA.Array(px, py)
    Else
'        MsgBox "Necessário configuração das coordenadas do clique do mouse para '" & strInfoPonteiro & "'", vbExclamation
        BuscaInfoPonteito = VBA.Array(0, 0)
    End If
End Function
'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cBL_Unear.PegaCodigoPorNome()
' TIPO             : Function
' DATA/HORA        : 08/08/2017 11:17
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Localiza o código na lista atual atraves do nome informado
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      08/08/2017 11:17
'---------------------------------------------------------------------------------------
Function PegaCodigoPorNome(strNome As String) As String
    Dim Tabela As Object, i As Integer
    Dim tBody As Object
    Dim tRow As Object
    'Tratamento de Erro
    '---------------------------------------------------------------------------------------
1   On Error GoTo PegaCodigoPorNome_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String:
    Const cstr_ProcedureName As String = "cBL_Unear.PegaCodigoPorNome()"
    '---------------------------------------------------------------------------------------

2   Call IEApp.WaitFor("bsDataTable", STOP_WHEN_FIND_ID)
3   Set Tabela = IEApp.Control("bsDataTable")
    Set tBody = Tabela.tBodies.item(0)
    
    'Aguarda o carregamento da tabela
    Do While tBody.Rows(0).Cells.Length < 2
        DoEvents
        Set Tabela = IEApp.Control("bsDataTable")
        Set tBody = Tabela.tBodies.item(0)
    Loop
    
4   For i = 0 To tBody.Rows.Length - 1
        Set tRow = tBody.Rows(i)
        If tRow.Cells.Length > 1 Then
5       If VBA.Trim(tRow.Cells(1).innertext) = strNome Then
6           PegaCodigoPorNome = VBA.Trim(tRow.Cells(0).innertext)
            Debug.Print Now, "Código '" & VBA.Trim(tRow.Cells(0).innertext) & "' Localizado !"
7           Exit Function
8       End If
        End If
9   Next i

Fim:
10  On Error GoTo 0
11  Exit Function

PegaCodigoPorNome_Error:
12  If Err <> 0 Then
13      lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
14      Call VBA.Err.Raise(VBA.Err.Number, cstr_ProcedureName, VBA.vbNewLine & cstr_ProcedureName & " Em" & VBA.IIf(VBA.Erl() <> 0, " > Linha (" & VBA.Erl() & ")", "") & " > " & VBA.Err.Description)
15  End If
    GoTo Fim:
    'Debug Mode
16  Resume
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cBL_Unear.LocalizarCampanha()
' TIPO             : Function
' DATA/HORA        : 08/08/2017 11:16
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Localiza uma campanha na lista atraves de um códigto. Retorna um objeto HTML DataCell
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      08/08/2017 11:16
'---------------------------------------------------------------------------------------
Function LocalizarCampanha(codigo As String) As Object
    Dim Campanha As Object
    'Tratamento de Erro
    '---------------------------------------------------------------------------------------
1   On Error GoTo LocalizarCampanha_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String:
    Const cstr_ProcedureName As String = "cBL_Unear.LocalizarCampanha()"
    '---------------------------------------------------------------------------------------
    Dim Tabela As Object
    Dim i As Integer, y As Integer

2   Set Tabela = IEApp.Control("bsDataTable")
3   For i = 0 To Tabela.tBodies.item(0).Rows.Length - 1
4       For y = 0 To Tabela.tBodies.item(0).Rows(i).Cells.Length - 1
5           If Tabela.tBodies.item(0).Rows(i).Cells(y).innertext = codigo Then
6               Set Campanha = Tabela.tBodies.item(0).Rows(i).Cells(y)
7           End If
8       Next y
9   Next i
10  Set LocalizarCampanha = Campanha

Fim:
11  On Error GoTo 0
12  Exit Function

LocalizarCampanha_Error:
13  If Err <> 0 Then
14      lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
15      Call VBA.Err.Raise(VBA.Err.Number, cstr_ProcedureName, VBA.vbNewLine & cstr_ProcedureName & " Em" & VBA.IIf(VBA.Erl() <> 0, " > Linha (" & VBA.Erl() & ")", "") & " > " & VBA.Err.Description)
16  End If
    GoTo Fim:
    'Debug Mode
17  Resume
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cBL_Unear.TestarCampanha()
' TIPO             : Function
' DATA/HORA        : 08/08/2017 11:16
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Envia solicitação de Teste da campanha
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      08/08/2017 11:16
'---------------------------------------------------------------------------------------
Function TestarCampanha(strCampanha As String)
'Tratamento de Erro
'---------------------------------------------------------------------------------------
1   On Error GoTo TestarCampanha_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String:
    Const cstr_ProcedureName As String = "cBL_Unear.TestarCampanha()"
    '---------------------------------------------------------------------------------------
    Dim strStatusAtual As String
    
    Call EnviaLogStatus(VBA.String(120, ".")): Call EnviaLogStatus("INICIANDO > " & cstr_ProcedureName): Call EnviaLogStatus(VBA.String(120, "."))
    
    strStatusAtual = BuscaInfoCampanhaNaLista(strCampanha, "Status")
    
2   If MenuEstaDisponivel("testCamp") And strStatusAtual = "CADASTRADA" Then
        Call EnviaLogStatus("Envia campanha para Teste...")
        '--------------------------------------------------------------------------------------------------------------------
3       Call MenuCampanha_Testar(strCampanha)
        '--------------------------------------------------------------------------------------------------------------------
        Call EnviaLogStatus("TESTE iniciado...")
        'Aguarda até que a campanha esteje com o status ATIVA
4       Call Me.AguardarStatus(strCampanha, "ATIVA", "Finalizando o teste, Aguardando disponibilização para Executar")
5       Call Me.SelecionarCampanhaNaLista(strCampanha)
        Call EnviaLogStatus("TESTE Finalizado...")
        
    ElseIf BuscaInfoCampanhaNaLista(strCampanha, "Status") = "TESTANDO" Then
        Call Me.AguardarStatus(strCampanha, "ATIVA", "Aguandao finalização do teste")
6   End If
    
    Call EnviaLogStatus("Verificando TESTE final...")
    
    strStatusAtual = BuscaInfoCampanhaNaLista(strCampanha, "Status")
    
    If VBA.UCase(strStatusAtual) Like "*ATIVA*" Then
        vRetorno = Array(1, "OK ! Testada. Pode Executar (Enviar para Liberaração)")
    ElseIf VBA.UCase(strStatusAtual) Like "*TESTANDO*" Then
        vRetorno = Array(2, "Tempo Excedido !. A campanha ainda não foi testada. Permaneceu com status (" & strStatusAtual & ")")
    End If
Fim:
    Call EnviaLogStatus("Status final do Teste > " & vRetorno(1))
    
    Call EnviaLogStatus(VBA.String(120, ".")): Call EnviaLogStatus("FINALIZADO ! > " & cstr_ProcedureName): Call EnviaLogStatus(VBA.String(120, "."))
    TestarCampanha = vRetorno
12  On Error GoTo 0
13  Exit Function

TestarCampanha_Error:
14  If Err <> 0 Then
15      lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
16      Call VBA.Err.Raise(VBA.Err.Number, cstr_ProcedureName, VBA.vbNewLine & cstr_ProcedureName & " Em" & VBA.IIf(VBA.Erl() <> 0, " > Linha (" & VBA.Erl() & ")", "") & " > " & VBA.Err.Description)
17  End If
    GoTo Fim:
    'Debug Mode
18  Resume
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cBL_Unear.ExecutarCampanha()
' TIPO             : Function
' DATA/HORA        : 08/08/2017 11:15
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Ultimo passo - Envia a solicitação da criação da campanha
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      08/08/2017 11:15
'---------------------------------------------------------------------------------------
Function ExecutarCampanha(strCampanha As String)
    Dim vRetorno
    Dim strStatusTeste As String
    'Tratamento de Erro
    '---------------------------------------------------------------------------------------
1   On Error GoTo ExecutarCampanha_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String:
    Const cstr_ProcedureName As String = "cBL_Unear.ExecutarCampanha()"
    '---------------------------------------------------------------------------------------
    Dim strStatusAtual As String
    
    Call EnviaLogStatus(VBA.String(120, ".")): Call EnviaLogStatus("INICIANDO > " & cstr_ProcedureName): Call EnviaLogStatus(VBA.String(120, "."))
    
2   vRetorno = Array(-1, "Nada executado")
    
    Call EnviaLogStatus("Preparando para executar. Verifica o status atual...")
    
    strStatusAtual = BuscaInfoCampanhaNaLista(strCampanha, "Status")
    
    If strStatusAtual Like "*TESTANDO*" Then
        Call EnviaLogStatus("Campanha ainda está TESTANDO, Aguarda mais 1 Minuto até que o teste seje concluído...")
        Call Me.AguardarStatus(strCampanha, "ATIVA", "Aguardando finalização do Teste pendente", 58)
    End If
    
    Call EnviaLogStatus("Espera foi finalizada, continua. Seleciona a campanha na lista...")
    Call Me.SelecionarCampanhaNaLista(strCampanha)
        
    Call EnviaLogStatus("Obtem novamente o status. Tem que estar ATIVA para que seja liberada")
    strStatusAtual = BuscaInfoCampanhaNaLista(strCampanha, "Status")
    
3   If strStatusAtual = "ATIVA" And Me.MenuEstaDisponivel("sendCamp") Then
        
        Call EnviaLogStatus("OK, campanha Testada, pronta para liberar. Envia o comando LIBERAÇÂO")
        
4       Call MenuCampanha_Enviar(strCampanha)
        
        Call EnviaLogStatus("Enviado !. Aguarda Login de autenticação")
5       Me.AguardaModalLogin

        Call EnviaLogStatus("Tela de Login OK, Envia dados de autenticação Usuario/Senha")
6       vRetorno = Me.PreencherLogin(Unear_Usuario(), Unear_Senha())
        Call EnviaLogStatus("LIBERAÇÃO - Finalizado !...")
        Call EnviaLogStatus("Login OK - Campanha enviada com sucesso")
    Else
        vRetorno = Array(0, "Campanha ainda está em teste. Não pode ser Executada. Tentar mais tarde. Cancelando....")
        Call EnviaLogStatus("ERROR - " & vRetorno(1))
7   End If

8   ExecutarCampanha = vRetorno

Fim:
9   On Error GoTo 0
    Call EnviaLogStatus(VBA.String(120, ".")): Call EnviaLogStatus("FINALIZADO > " & cstr_ProcedureName): Call EnviaLogStatus(VBA.String(120, "."))
10  Exit Function

ExecutarCampanha_Error:
11  If Err <> 0 Then
12      lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
13      Call VBA.Err.Raise(VBA.Err.Number, cstr_ProcedureName, VBA.vbNewLine & cstr_ProcedureName & " Em" & VBA.IIf(VBA.Erl() <> 0, " > Linha (" & VBA.Erl() & ")", "") & " > " & VBA.Err.Description)
14  End If
    GoTo Fim:
    'Debug Mode
15  Resume
End Function

Function BaixarArquivosRetorno(codigo As String)
    Dim vRetorno
    Dim qntBaixados As Integer
    Dim qtdEncontrados As Integer
    'Tratamento de Erro
    '---------------------------------------------------------------------------------------
1   On Error GoTo BaixarArquivosRetorno_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String:
    Const cstr_ProcedureName As String = "cBL_Unear.BaixarArquivosRetorno()"
    '---------------------------------------------------------------------------------------
    Dim strArquivo As String
    Dim strPastaDownloads As String
    
    vRetorno = Array(-1, "Nada executado")
    strPastaDownloads = PastaDownloads_IE
    
    
    Call AguardarStatus(codigo, "ENCERRADA", "Aguardando disponibilização para Download")
    Call SelecionarCampanhaNaLista(codigo)
    
    IEApp.WaitFor "campSendData", STOP_WHEN_FIND_ID
    
    If Me.BuscaInfoCampanhaNaLista(codigo, "Status") = "ENCERRADA" And Me.MenuEstaDisponivel("campSendData") And VBA.Now > VBA.CDate(Me.BuscaInfoCampanhaNaLista(codigo, "Fim")) Then
        Call ProcessarAcaoMenuCampanha("campSendData", codigo)
        Call AguardaCarregamentoTabela
        'If IEApp.Browser.NotificationBaseIsVisible Then Call IEApp.Browser.CloseNotificationBase
        'Pega todos os links com Name = Export
        Set cLinks = IEApp.GetElementsByTag(IEApp.Document.getElementById("bsDataTable"), "a", Array("name", "export"))
        
        For Each vLink In cLinks
            qtdEncontrados = qtdEncontrados + 1
            
            If vLink.element.parentelement.nextSibling.innertext = "Emails Enviados" Then
                strArquivo = codigo & "_enviados.csv"
            ElseIf vLink.element.parentelement.nextSibling.innertext = "Sem Resposta" Then
                strArquivo = codigo & "_sem resposta.csv"
            ElseIf vLink.element.parentelement.nextSibling.innertext = "Emails Entregues" Then
                strArquivo = codigo & "_emails entregues.csv"
            ElseIf vLink.element.parentelement.nextSibling.innertext = "Erro no Envio" Then
                strArquivo = codigo & "_erro no envio.csv"
            ElseIf vLink.element.parentelement.nextSibling.innertext = "Emails Abertos" Then
                strArquivo = codigo & "_emails abertos.csv"
            End If
            '
            Debug.Print "Baixando...."; qtdEncontrados & " de " & cLinks.count & " > " & strArquivo & " > (" & vLink.element.parentelement.nextSibling.innertext & ")"
            Call IEApp.Click(vLink)
            Call IEApp.WaitFor
'            Call IEApp.Browser.WaitDownloadBar
            
            strArquivo = strPastaDownloads & "\" & strArquivo
            
            Call AuxFileSystem.DeleteFile(strArquivo)
            
            Call ClicarMouseNaPosicao("Download_ClickSalvar")
            
            
            win.Wait
            Call ClicarMouseNaPosicao("Download_ClickFechar")

            Do While VBA.Dir(strArquivo) = ""
                Debug.Print "Aguardando finalizacao do download"
            Loop
            
            'Se nao estiver mais disponivel
            Debug.Print "Arquivo baixado com sucesso !"
            qntBaixados = qntBaixados + 1
            
        Next vLink
    End If
8   BaixarArquivosRetorno = vRetorno

Fim:
9   On Error GoTo 0
10  Exit Function

BaixarArquivosRetorno_Error:
11  If VBA.Err <> 0 Then
12      lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
13      Call VBA.Err.Raise(VBA.Err.Number, cstr_ProcedureName, VBA.vbNewLine & cstr_ProcedureName & " Em" & VBA.IIf(VBA.Erl() <> 0, " > Linha (" & VBA.Erl() & ")", "") & " > " & VBA.Err.Description)
14  End If
    GoTo Fim:
    'Debug Mode
15  Resume

End Function


'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cBL_Unear.AguardaModalLogin()
' TIPO             : Function
' DATA/HORA        : 08/08/2017 11:14
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Aguarda ate que a tela de Login esteja disponivel
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      08/08/2017 11:14
'---------------------------------------------------------------------------------------
Function AguardaModalLogin()
    Dim dTimeOut As Date, bRetorno As Boolean
    'Tratamento de Erro
    '---------------------------------------------------------------------------------------
1   On Error GoTo AguardaModalLogin_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String:
    Const cstr_ProcedureName As String = "cBL_Unear.AguardaModalLogin()"
    '---------------------------------------------------------------------------------------

2   dTimeOut = PreparaTimeOut()
3   bRetorno = False
4   Do While Not bRetorno
5       Debug.Print
        Call showMessage("LIBERAÇÃO - Aguardando Modal de Login para conclusão...", , SemProgresso)
6       bRetorno = IEApp.ControlExists("login") And IEApp.ControlExists("password") And IEApp.ControlExists("popup")
7       If DeuTimeOut(dTimeOut) Then
            bRetorno = False:
8           Debug.Print "Timeout - Aguarda Modal de Login...."
9           Exit Do
10      End If
11  Loop
12  If bRetorno Then Call showMessage("LIBERAÇÃO - Modal de Login OK...", , SemProgresso)

Fim:
15  AguardaModalLogin = bRetorno

16  On Error GoTo 0
17  Exit Function

AguardaModalLogin_Error:
18  If Err <> 0 Then
19      lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
20      Call VBA.Err.Raise(VBA.Err.Number, cstr_ProcedureName, VBA.vbNewLine & cstr_ProcedureName & " Em" & VBA.IIf(VBA.Erl() <> 0, " > Linha (" & VBA.Erl() & ")", "") & " > " & VBA.Err.Description)
21  End If
    GoTo Fim:
    'Debug Mode
22  Resume
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cBL_Unear.PreencherLogin()
' TIPO             : Function
' DATA/HORA        : 08/08/2017 13:12
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Preenche a tela de login
'---------------------------------------------------------------------------------------
Function PreencherLogin(sLogin As String, sSenha As String)
    Dim dTimeOut As Date, bRetorno
    'Tratamento de Erro
    '---------------------------------------------------------------------------------------
1   On Error GoTo PreencherLogin_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String:
    Const cstr_ProcedureName As String = "cBL_Unear.PreencherLogin()"
    '---------------------------------------------------------------------------------------
    Dim msg As String
2   dTimeOut = PreparaTimeOut()
    'Modal de Login
    'login
    'Preenche login e senha no Modal
3   If IEApp.ControlExists("login") And IEApp.ControlExists("password") And IEApp.ControlExists("popup") Then
4       IEApp.Control("login").value = sLogin
5       IEApp.Control("password").value = sSenha
6       Call IEApp.Click("btnSave")
7       Do
8           If IEApp.ControlExists("login") Then
9               If IEApp.Control("login") = "" Then
10                  bRetorno = Array(1, "OK")
11              ElseIf IEApp.Control("login") = "" Then
12                  bRetorno = Array(2, "Erro")
13              End If

                If IEApp.ControlExists("fileMessage") Then
                    msg = IEApp.Control("fileMessage").innertext
                    If msg <> "" Then
                        bRetorno = Array(2, "Erro - " & msg)
                        Call IEApp.Click("btnCancel")
                    Else
                        bRetorno = Array(1, "OK")
                    End If
                    
                    Exit Do
                End If
15          End If

16          If DeuTimeOut(dTimeOut) Then
17              bRetorno = Array(3, "Timeout")
18              Exit Do
19          End If
20      Loop
21  Else
22      bRetorno = Array(3, "Não abriu modal de Login")
23  End If
24  PreencherLogin = bRetorno

Fim:
25  On Error GoTo 0
26  Exit Function

PreencherLogin_Error:
27  If Err <> 0 Then
28      lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
29      Call VBA.Err.Raise(VBA.Err.Number, cstr_ProcedureName, VBA.vbNewLine & cstr_ProcedureName & " Em" & VBA.IIf(VBA.Erl() <> 0, " > Linha (" & VBA.Erl() & ")", "") & " > " & VBA.Err.Description)
30  End If
    GoTo Fim:
31  Resume    'Debug Mode
End Function

Private Sub MenuCampanha_Enviar(codigo As String)
1   Call ProcessarAcaoMenuCampanha("sendCamp", codigo)
End Sub

Private Sub MenuCampanha_Testar(codigo As String)
1   Call ProcessarAcaoMenuCampanha("testCamp", codigo)
End Sub

Sub MenuCampanha_Editar(codigo As String)
1   Call ProcessarAcaoMenuCampanha("editCamp", codigo)
End Sub

Sub MenuCampanha_Preview(codigo As String)
1   Call ProcessarAcaoMenuCampanha("previewCamp", codigo)
End Sub

Sub MenuCampanha_Remover(codigo As String)
1   Call ProcessarAcaoMenuCampanha("removeCamp", codigo)
End Sub

'Verifica se o menu tenm o atributo 'disabled'
Function MenuEstaDisponivel(strMenu As String) As Boolean
1   If IEApp.ControlExists(strMenu) Then
2       MenuEstaDisponivel = Not IEApp.ControlHasAttibute(IEApp.Control(strMenu), "disabled")
3   End If
End Function

Sub DetalheCamanha_Voltar()
1   Call IEApp.Click("btnCancel")
End Sub

Function DesmarcarCampanhaNaLista(codigo As String) As Boolean
    On Error GoTo Err_Handle
    Dim Campanha As Object
1   Set Campanha = LocalizarCampanha(codigo)
2   If Not Campanha Is Nothing Then
        'Ao localiza o ID, verificar se a linha ja está selecionada. Se nao tiver, clica e seleciona
3       If Campanha.parentNode.ClassName = "active" Then
4           Call Campanha.Click
5           DesmarcarCampanhaNaLista = True
6       End If
7   End If
Fim:
    Exit Function
Err_Handle:
    If VBA.Err <> 0 Then Debug.Print "DesmarcarCampanhaNaLista()", VBA.Error
    GoTo Fim
    Resume
End Function

Function SelecionarCampanhaNaLista(codigo As String) As Boolean
    On Error GoTo Err_Handle
    Dim Campanha As Object
1   Set Campanha = LocalizarCampanha(codigo)
2   If Not Campanha Is Nothing Then
        'Ao localiza o ID, verificar se a linha ja está selecionada. Se nao tiver, clica e seleciona
3       If Campanha.parentNode.ClassName <> "active" Then
            Call Campanha.Click
4           SelecionarCampanhaNaLista = True
        Else
            SelecionarCampanhaNaLista = True
        End If
5   End If
Fim:
    Exit Function
Err_Handle:
    If VBA.Err <> 0 Then Debug.Print "SelecionarCampanhaNaLista()", VBA.Error
    GoTo Fim
    Resume
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cBL_Unear.AguardarStatus()
' TIPO             : Function
' DATA/HORA        : 08/08/2017 13:12
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Aguarda até q o status esperado seja lido
'---------------------------------------------------------------------------------------
Function AguardarStatus(strCampanha As String, strStatusEsperado As String, strMensagem As String, Optional tempoEspera As Integer = 55) As Boolean
    Dim dTimeOut As Date, bRetorno As Boolean
    Dim strStatusAtual As String
    'Tratamento de Erro
    '---------------------------------------------------------------------------------------
1   On Error GoTo AguardarStatus_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String:
    Const cstr_ProcedureName As String = "cBL_Unear.AguardarStatus()"
    '---------------------------------------------------------------------------------------
    'Aguarda 50 segs
2   dTimeOut = PreparaTimeOut(tempoEspera)
    strStatusAtual = BuscaInfoCampanhaNaLista(strCampanha, "STATUS")
3   Do While strStatusAtual <> strStatusEsperado
4       Me.SelecionarListaPorMenu
5       Call showMessage(strMensagem & VBA.vbNewLine & "Status Atual : " & strStatusAtual & " / Status Esperado : " & strStatusEsperado, , SemProgresso)
        'Aguarda 5 segundos para a proxima conferencia
        Call win.Wait(5000)
        strStatusAtual = BuscaInfoCampanhaNaLista(strCampanha, "STATUS")
6       If strStatusAtual = strStatusEsperado Then
7           bRetorno = True
8           Exit Do
9       ElseIf DeuTimeOut(dTimeOut) Then
10          bRetorno = False
11          Exit Do
12      End If
13  Loop
14  AguardarStatus = bRetorno
Fim:
15  On Error GoTo 0
16  Exit Function

AguardarStatus_Error:
17  If VBA.Err <> 0 Then
18      lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
19      Call VBA.Err.Raise(VBA.Err.Number, cstr_ProcedureName, VBA.vbNewLine & cstr_ProcedureName & " Em" & VBA.IIf(VBA.Erl() <> 0, " > Linha (" & VBA.Erl() & ")", "") & " > " & VBA.Err.Description)
20  End If
    GoTo Fim:
21  Resume    'Debug Mode
End Function

'Aguarda o carregamento da lista
Function AguardaCarregamentoDaLista()
    On Error GoTo Erro
    Dim div As Object
1   Call IEApp.WaitFor("bsDataTable", STOP_WHEN_FIND_ID)
    win.Wait 4000
'2   Set div = IEApp.GetElementByTag(IEApp.Document, "div", Array("class", "fixed-table-loading"))
'3   Do While div Is Nothing
'4       Set div = IEApp.GetElementByTag(IEApp.Document, "div", Array("class", "fixed-table-loading"))
'5       Call sendLog("Aguardando finalização do carregamento da lista")
'6   Loop
Exit Function
Erro:
    If VBA.Err <> 0 Then
        Stop
        Debug.Print VBA.Error
    End If
    Resume
End Function

'Força a seleção da lista clicando na opção à esquerda
Sub SelecionarListaPorMenu()
    On Error Resume Next
1   Call IEApp.WaitFor("Group_1", STOP_WHEN_FIND_ID)
    'Seleciona a opção de menu 'Selecoes do Usuario'
2   Call IEApp.Click(IEApp.Control("Group_1").Children(1))
3   Call AguardaCarregamentoDaLista
End Sub


'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cBL_Unear.BuscaInfoCampanhaNaLista()
' TIPO             : Function
' DATA/HORA        : 08/08/2017 13:11
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Recupera uma informação da lista
'---------------------------------------------------------------------------------------
Public Function BuscaInfoCampanhaNaLista(codigo As String, strColuna As String)
    'Tratamento de Erro
    '---------------------------------------------------------------------------------------
1   On Error GoTo BuscaInfoCampanhaNaLista_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String:
    Const cstr_ProcedureName As String = "cBL_Unear.BuscaInfoCampanhaNaLista()"
    '---------------------------------------------------------------------------------------
    Dim linha As Object
    Dim Campanha As Object
    Dim Tabela As Object
    Dim i As Integer

2   Set Campanha = LocalizarCampanha(codigo)
3   If Not Campanha Is Nothing Then

4       Set Tabela = Campanha.parentNode.parentNode.parentNode

5       For i = 0 To Tabela.tHead.Rows.item(0).Cells.Length - 1
6           If VBA.UCase(VBA.Trim(Tabela.tHead.Rows.item(0).Cells(i).innertext)) = VBA.UCase(strColuna) Then
7               BuscaInfoCampanhaNaLista = VBA.Trim(Campanha.parentNode.Cells(i).innertext)
8               Exit For
9           End If
10      Next i
11  End If

Fim:
12  On Error GoTo 0
13  Exit Function

BuscaInfoCampanhaNaLista_Error:
14  If Err <> 0 Then
15      lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
16      Call VBA.Err.Raise(VBA.Err.Number, cstr_ProcedureName, VBA.vbNewLine & cstr_ProcedureName & " Em" & VBA.IIf(VBA.Erl() <> 0, " > Linha (" & VBA.Erl() & ")", "") & " > " & VBA.Err.Description)
17  End If
    GoTo Fim:
18  Resume    'Debug Mode
End Function

'Executa um Menu
Sub ProcessarAcaoMenuCampanha(strAcao As String, codigo As String)
    Dim Campanha As Object
    win.Wait
1   If SelecionarCampanhaNaLista(codigo) Then
2       Call IEApp.Click(strAcao)
3   End If
End Sub

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cBL_Unear.UploadListaEmailsTXT()
' TIPO             : Function
' DATA/HORA        : 08/08/2017 11:27
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Localize o ID da janela de upload do arquivo, preenche o endereço do arquivo informado na caixa "Nome do Arquivo" e clica em OK
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      08/08/2017 11:27
'---------------------------------------------------------------------------------------
Private Function UploadListaEmailsTXT(strArquivo As String) As Boolean
    Dim hwndP As Long
    Dim oWindow As cTFW_Window
    Dim edit As cTFW_Window
    'Tratamento de Erro
    '---------------------------------------------------------------------------------------
1   On Error GoTo UploadListaEmailsTXT_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String:
    Const cstr_ProcedureName As String = "cBL_Unear.UploadListaEmailsTXT()"
    '---------------------------------------------------------------------------------------

2   Set oWindow = New cTFW_Window
    'Localiza uma janela atraves do título
3   Call oWindow.GetWindow(pCaptionOrHandle:="Escolher arquivo a carregar")
4   If oWindow.hwnd = 0 Then
5       UploadListaEmailsTXT = False
6   Else
7       Call oWindow.LoadProperties
        'Enumera todas as janelas filhas da janela acima
8       Call oWindow.LoadChildren(APISupport.GetWindows(oWindow.hwnd))
        'Localiza a primeira Caixa de Edição
9       Set edit = oWindow.FindEdit()
10      If Not edit Is Nothing Then
            'Envia o texto para caixa de texto instaciada
11          Call edit.SetText(strArquivo)
            'Clica no botão OK
12          Call oWindow.ClickOn(OKButton)
13          UploadListaEmailsTXT = True
14      Else
15          Call VBA.MsgBox("Nenhuma caixa de dialogo localizada !", VBA.vbExclamation)
16          UploadListaEmailsTXT = False
17      End If
18  End If

Fim:
19  On Error GoTo 0
20  Exit Function

UploadListaEmailsTXT_Error:
21  If Err <> 0 Then
22      lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
23      Call VBA.Err.Raise(VBA.Err.Number, cstr_ProcedureName, VBA.vbNewLine & cstr_ProcedureName & " Em" & VBA.IIf(VBA.Erl() <> 0, " > Linha (" & VBA.Erl() & ")", "") & " > " & VBA.Err.Description)
24  End If
    GoTo Fim:
    'Debug Mode
25  Resume
End Function
'Envia uma mensagem para o evento 'RegistraLog' para que seja capturado por uma aplicação
Private Sub sendLog(pMensagem As String)
1   Debug.Print Now & " " & pMensagem
2   RaiseEvent SaveRegistraLog(Now & " " & pMensagem, "URL")
End Sub

Public Property Get Navegador_StatusAtual() As String
1   Navegador_StatusAtual = m_Navegador_StatusAtual
End Property

Public Property Let Navegador_StatusAtual(ByVal vNewValue As String)
1   m_Navegador_StatusAtual = vNewValue
End Property

Public Property Get CRM_Conectado() As Boolean
    CRM_Conectado = CRM_Logado
End Property

Public Property Let CRM_Conectado(ByVal vNewValue As Boolean)
1   m_CRM_Conectado = vNewValue
End Property

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cBL_Unear.CRM_TelaIndex()
' TIPO             : Property
' DATA/HORA        : 08/08/2017 13:10
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Recupera o código relacionado à tela atual
'---------------------------------------------------------------------------------------
Public Property Get CRM_TelaIndex() As Integer
'Tratamento de Erro
'---------------------------------------------------------------------------------------
1   On Error GoTo CRM_TelaIndex_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String:
    Const cstr_ProcedureName As String = "cBL_Unear.CRM_TelaIndex()"
    '---------------------------------------------------------------------------------------

2   m_iCRM_TelaIndex = 0
3   If IEDrive Is Nothing Then Exit Property
4   If IEApp.Url Like "*unear.net*" Then
5       m_CRM_Conectado = True
6       If IEApp.ControlExists("username") And IEApp.ControlExists("username") Then
7           m_iCRM_TelaIndex = 1
8       ElseIf IEApp.ControlExists("campaignList") Then
9           m_iCRM_TelaIndex = 2
10      ElseIf IEApp.ControlExists("selectionList") Then
11          m_iCRM_TelaIndex = 3
12      ElseIf IEApp.ControlExists("uploadResult") Then
13          m_iCRM_TelaIndex = 4
14      Else
15          CRM_TelaIndex = 0
16      End If
17  Else
18      CRM_TelaIndex = 0
19  End If
20  CRM_TelaIndex = m_iCRM_TelaIndex

Fim:
21  On Error GoTo 0
22  Exit Property

CRM_TelaIndex_Error:
23  If Err <> 0 Then
24      lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
25      Call VBA.Err.Raise(VBA.Err.Number, cstr_ProcedureName, VBA.vbNewLine & cstr_ProcedureName & " Em" & VBA.IIf(VBA.Erl() <> 0, " > Linha (" & VBA.Erl() & ")", "") & " > " & VBA.Err.Description)
26  End If
    GoTo Fim:
27  Resume    'Debug Mode
End Property

Public Property Get CRM_Logado() As Integer
    Dim bLogado As Boolean
    bLogado = IEApp.ControlExists("appBody") '
    bLogado = bLogado And IEApp.ControlExists("app") '
    bLogado = bLogado And IEApp.ControlExists("headerContent") '
    bLogado = bLogado And IEApp.ControlExists("searchTerm") '
4   CRM_Logado = bLogado
End Property


Public Property Get CRM_TelaAtual() As String
1   m_CRM_Conectado = False
2   m_iCRM_TelaIndex = 0
3   If IEApp.Url Like "*unear.net*" Then
4       m_CRM_Conectado = True
5       If IEApp.ControlExists("username") And IEApp.ControlExists("username") Then
6           CRM_TelaAtual = "Login"
7       ElseIf IEApp.ControlExists("campaignList") Then
8           CRM_TelaAtual = "Lista de Campanhas"
9       ElseIf IEApp.ControlExists("selectionList") Then
10          CRM_TelaAtual = "Lista de Seleções"
11      Else
12          CRM_TelaAtual = "Tela não identificada"
13      End If
14  Else
15      CRM_TelaAtual = "Não está no CRM"
16  End If
End Property

Function PerfilSelecionado() As String
    Dim ctrSpan As Object
1   Set ctrSpan = Browser.Control("btnTitle")
2   If Not ctrSpan Is Nothing Then
3       Set ctrSpan = Browser.ParseControl(ctrSpan)
4       PerfilSelecionado = ctrSpan.value
5   End If
End Function

'"bsDataTable"
'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cBL_Unear.DownloadLista()
' TIPO             : Function
' DATA/HORA        : 08/08/2017 13:09
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Baixa o conteudo da lista atual para uma tabela local
'---------------------------------------------------------------------------------------
Public Function DownloadLista(tipoTabela As String, idListaHTML As String) As String
    Dim sTabelaNome As String, strTableName As String, strCampos As String
    Dim Tabela As Object
    Dim i As Integer, y As Integer
    Dim strSQLTable As String
    Dim rs As Object

    'Tratamento de Erro
    '---------------------------------------------------------------------------------------
1   On Error GoTo DownloadLista_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String:
    Const cstr_ProcedureName As String = "cBL_Unear.DownloadLista()"
    '---------------------------------------------------------------------------------------

2   Do While Not IEApp.ControlExists(idListaHTML): VBA.DoEvents: Loop

3   Set Tabela = IEApp.Control("bsDataTable")
4   strTableName = tipoTabela & "_" & Tabela.ID & "_" & VBA.Format(Now, "yyyymmdd_hhnnss")
5   strSQLTable = "CREATE TABLE " & strTableName & "("
6   strCampos = ""

7   For i = 0 To Tabela.tHead.Rows.item(0).Cells.Length - 1
8       strCampos = strCampos & VBA.vbNewLine & "   ,[" & Tabela.tHead.Rows.item(0).Cells(i).innertext & "] TEXT"
9   Next i
10  strSQLTable = strSQLTable & VBA.Mid(strCampos, 7) & VBA.vbNewLine & ")"

11  CurrentDb.Execute strSQLTable

12  Set rs = CurrentDb.OpenRecordset(strTableName)
13  For i = 0 To Tabela.tBodies.item(0).Rows.Length - 1
14      rs.addNew
15      For y = 0 To Tabela.tBodies.item(0).Rows(i).Cells.Length - 1
16          rs(y).value = Tabela.tBodies.item(0).Rows(i).Cells(y).innertext
17      Next y
18      rs.Update
19  Next i
20  DownloadLista = strTableName

Fim:
21  On Error GoTo 0
22  Exit Function

DownloadLista_Error:
23  If Err <> 0 Then
24      lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
25      Call VBA.Err.Raise(VBA.Err.Number, cstr_ProcedureName, VBA.vbNewLine & cstr_ProcedureName & " Em" & VBA.IIf(VBA.Erl() <> 0, " > Linha (" & VBA.Erl() & ")", "") & " > " & VBA.Err.Description)
26  End If
    GoTo Fim:
27  Resume    'Debug Mode
End Function

Private Property Get Browser() As cIE_Tab
1   Set Browser = IEApp
End Property

Public Property Set Browser(objBrowser As cIE_Tab)
1   Set m_objBrowser = objBrowser
2   Set IEDrive = objBrowser
End Property

Private Sub Class_Initialize()
1   Call AtualizaInstancia
End Sub

Private Function PreparaTimeOut(Optional dTimeOutLimite As Integer = 10)
1   PreparaTimeOut = VBA.DateAdd("s", dTimeOutLimite, VBA.Now())
End Function

Private Function DeuTimeOut(dInicial As Date) As Boolean
1   If VBA.CDbl(dInicial) > 0 Then If VBA.Now >= dInicial Then DeuTimeOut = True: Exit Function
End Function

Private Sub IEDrive_OnFinish()
    Call EventoOcupado(False)
End Sub

Private Sub IEDrive_OnRunning(pMessage As String)
    Call EventoOcupado(True)
End Sub

Function BrowserEstaAberto() As Boolean
    If Not IEDrive Is Nothing Then
        If Not IEDrive.Browser Is Nothing Then
            BrowserEstaAberto = IEDrive.Browser.BroswerIsOpen
        End If
    End If
End Function

Sub EnviaLogStatus(strMensagem As String)
    Call RegistraStatus(Mensagem:=VBA.Trim(strMensagem), LogTxt:=Me.LogFile, Progresso:=True)
End Sub
