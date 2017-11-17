Attribute VB_Name = "mBL_CRMUnear"
Option Compare Database
Option Explicit

Public bRodando As Boolean
Public bCancelar As Boolean

Public cUnear As New cBL_Unear
Public win As New cTFW_Window
Public arrPonteiro As Variant
Private COD_DISPARO As String
Private COD_CAMPANHA As String
Public Enum eAcao
    Testar = 1
    Executar = 2
End Enum

Public Function NovoCodigo(Optional ByRef sCodigo As String): sCodigo = VBA.Format(VBA.Now(), "yyyymmddhhnnss"): NovoCodigo = sCodigo: End Function


'Recupera configurações salvas
Public Function EmailsLogs(): EmailsLogs = Nz(DLookup("EmailsLogs", "tblConfig"), ""): End Function
Public Function CaixaDeCorreioRemetente(): CaixaDeCorreioRemetente = Nz(DLookup("CaixaDeCorreioRemetente", "tblConfig"), ""): End Function
Public Function URLSiteUnear(): URLSiteUnear = Nz(DLookup("URLSiteUnear", "tblConfig"), ""): End Function
Public Function PerfilMousePointer(): PerfilMousePointer = Access.Nz(Access.DLookup("PerfilPadraoMousePointer", "tblConfig"), ""): End Function
Public Function PastaRaizListas(): PastaRaizListas = Access.Nz(Access.DLookup("PastaLocalListas", "tblConfig"), ""): End Function
Public Function Unear_Usuario(): Unear_Usuario = Nz(DLookup("Unear_Usuario", "tblConfig"), ""): End Function
Public Function Unear_Senha(): Unear_Senha = Nz(DLookup("Unear_Senha", "tblConfig"), ""): End Function
Public Function PastaDownloads_IE()
    Dim strPasta As String
    strPasta = Nz(DLookup("PastaDownloads_IE", "tblConfig"), "")
    If strPasta = "" Then
       strPasta = AuxFileSystem.PegarPasta(AuxFileSystem.PegarPasta(AuxFileSystem.PegarPasta(VBA.Environ("temp")))) & "\Downloads"
    End If
    PastaDownloads_IE = strPasta
End Function
'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : mBL_CRMUnear.GerarCampanhas()
' TIPO             : Function
' DATA/HORA        : 08/08/2017 10:58
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Inicia a geração das campanhas
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      08/08/2017          Adelson         Inicio da rotina
' * 1.01      31/08/2017          Adelson         Ajustes nos processos
' * 1.02      11/09/2017          Adelson         Melhoria no registro de status de geração da campanha
' * 1.03      12/09/2017          Adelson         Correções no envio de SMS
'---------------------------------------------------------------------------------------

Sub GerarCampanhas(Optional vArgs As Variant)
    Dim rsListaDisparo As Object, dtProcesso As Date
    Dim strSelecao As String, strCampanha As String
    Dim bRetorno, strStatus As String
    Dim intID As Long
    Dim strNomeSelecao As String
    Dim strNomeCampanha As String
    Dim cValidacao As Collection
    Dim i As Integer
    Dim strArquivoSelecao As String
    Dim strArquivoLog As String
    Dim strArquivoLogDetalhes As String
    Dim strLogCampanhas As String
    Dim infoAdicionais As String
    Dim statusTeste As String, statusExecucao
    Dim statusFinal As String
    Dim job As cJob
    Dim tm As cTimer
    Dim cL As Collection
    Dim sTimerID As String, sJobID As String
    Dim strCod As String
    Dim bStatus As Boolean
    Dim arrCampanhas As New VBA.Collection
    
    '---------------------------------------------------------------------------------------
    On Error GoTo MacroCampanhaCRM_Unear_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
    Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Function mBL_CRMUnear.MacroCampanhaCRM_Unear()"
    '---------------------------------------------------------------------------------------
    
    cUnear.cstrURL_Unear = URLSiteUnear()
    
    If cUnear.cstrURL_Unear = "" Then
        MsgBox "URL do site do Unear não configurado. Necessário ajustar os paramentros de configurações !", vbExclamation
        Exit Sub
    End If
    
    Call Inicializar_Globais
    
    '---------------------------------------------------------------------------
    'Registra informações sobre a execução
    '---------------------------------------------------------------------------
    Call NovoCodigo(COD_CAMPANHA)
    strArquivoLog = PegaEndereco_Programa & "\Log\GerarCampanhas"
    Call AuxFileSystem.MkFullDirectory(strArquivoLog)
    strLogCampanhas = strArquivoLog & "\Campanhas.txt"
    Call AuxFileSystem.MkFullDirectory(strArquivoLog & "\Detalhes\")
    strArquivoLogDetalhes = strArquivoLog & "\Detalhes\" & COD_CAMPANHA & ".txt"
    cUnear.LogFile = strArquivoLogDetalhes
    
    If Not VBA.IsMissing(vArgs) Then
        Set job = vArgs(0)
        sTimerID = job.JobTimer.TimerID
        sJobID = job.JobID
        Set cL = vArgs(1)
        infoAdicionais = "Job > Nome : " & job.nome & " | Inicio : " & job.StartTime & " | Qtd Executado : " & job.CountExecuted
    End If
    
    Call RegistraStatus(Mensagem:="Log Iniciado............", LogTxt:=strArquivoLogDetalhes, Progresso:=True, IncluirDataHoraLog:=False)
    Call RegistraStatus(Mensagem:="Data/Hora    : " & VBA.Now(), LogTxt:=strArquivoLogDetalhes, Progresso:=True, IncluirDataHoraLog:=False)
    Call RegistraStatus(Mensagem:="Timer ID     : " & sTimerID, LogTxt:=strArquivoLogDetalhes, Progresso:=True, IncluirDataHoraLog:=False)
    Call RegistraStatus(Mensagem:="Job ID       : " & sJobID, LogTxt:=strArquivoLogDetalhes, Progresso:=True, IncluirDataHoraLog:=False)
    
    Call RegistraStatus(Mensagem:=VBA.String(120, "-"), LogTxt:=strArquivoLogDetalhes, Progresso:=True, IncluirDataHoraLog:=False)
    Call MarcaJob("GerarCampanhas")
    
    Set rsListaDisparo = Conexao.PegarRS("Pegar_ListaDisparo")

    'Passo 1
    Call RegistraStatus(Mensagem:="Acessando a plataforma :.." & cUnear.cstrURL_Unear, LogTxt:=strArquivoLogDetalhes, Progresso:=True)
    
    bRetorno = cUnear.AcessarSite
    Call cUnear.AguardaTelaLogin
    Call RegistraStatus(Mensagem:="Autenticando usuario/senha (" & Unear_Usuario() & ")...", LogTxt:=strArquivoLogDetalhes, Progresso:=True)
    If Not cUnear.CRM_Conectado Then
        Call RegistraStatus(Mensagem:="Realizando o logon...", LogTxt:=strArquivoLogDetalhes, Progresso:=True)
        bRetorno = cUnear.Entrar(Unear_Usuario(), Unear_Senha())
    Else
        Call RegistraStatus(Mensagem:="Ja está logado !...", LogTxt:=strArquivoLogDetalhes, Progresso:=True)
    End If
    
    'Passo 2
    If VBA.IsArray(bRetorno) Then Call RegistraStatus(Mensagem:=VBA.CStr(bRetorno(0)), LogTxt:=strArquivoLogDetalhes, Progresso:=True)
    Call RegistraStatus(Mensagem:=VBA.String(120, "-") & VBA.vbNewLine, LogTxt:=strArquivoLogDetalhes, Progresso:=True, IncluirDataHoraLog:=False)
    
    
    Do While Not rsListaDisparo.EOF
        dtProcesso = VBA.Timer
        intID = Nz(rsListaDisparo!ID.value, 0)
        Call NovoCodigo(COD_DISPARO)
        strCod = "TFWBOT|-" & VBA.Format(intID, "00") & "-" & COD_DISPARO & "| - "
        
        Call RegistraStatus(Mensagem:=VBA.String(120, "*"), LogTxt:=strArquivoLogDetalhes, Progresso:=True, IncluirDataHoraLog:=False)
        Call RegistraStatus(Mensagem:="INICIO EVENTO : #" & VBA.Format(intID, "00") & "-" & COD_DISPARO, LogTxt:=strArquivoLogDetalhes, Progresso:=True, IncluirDataHoraLog:=False)
        Call RegistraStatus(Mensagem:=VBA.String(120, "*"), LogTxt:=strArquivoLogDetalhes, Progresso:=True, IncluirDataHoraLog:=False)
        
        Call RegistrarLogEvento(COD_DISPARO, "Iniciando, Aguarde...", strArquivoLogDetalhes, "", False, sTimerID, sJobID, "#" & VBA.Format(intID, "00"))
        
        Call salvaValor("tipoCampanha", rsListaDisparo!TipoCampanha.value)
        Call salvaValor("Assunto", rsListaDisparo!Assunto.value)
        Call RegistraStatus(Mensagem:="#" & VBA.Format(intID, "00") & " - Preparando para criar gerar a campanha...", LogTxt:=strArquivoLogDetalhes, Progresso:=True)
        Call AtualizaListaDisparo(intID, "IDCampanha", "")
        Call AtualizaListaDisparo(intID, "Status", "Aguardando...")
        Call LimparEtapas
        Set cValidacao = ValidaParametros(rsListaDisparo)
        If cValidacao.count > 0 Then
            Call RegistraStatus(Mensagem:="Processo será cancelado ! Erros de validação encontrados", LogTxt:=strArquivoLogDetalhes, Progresso:=True)
            For i = 1 To cValidacao.count
                Call RegistraStatus(Mensagem:="Erro : " & cValidacao.item(i), LogTxt:=strArquivoLogDetalhes, Progresso:=True)
            Next i
            Call AtualizaListaDisparo(intID, "Status", "0 - Erro. Consulte o log para mais detalhes !")
            GoTo Proxima
        End If
        Call RegistraStatus(Mensagem:="Parametros necessários foram validados, continuar a criação da campanha...", LogTxt:=strArquivoLogDetalhes, Progresso:=True)
        '-------------------------------------------------------------------------------------
        'Passo 3 -
        '-------------------------------------------------------------------------------------
        bRetorno = cUnear.AcessarListaDeCampanhas()
        'Call RegistrarLogEtapa("ListaDeCampanhas", bRetorno)
            
        '-------------------------------------------------------------------------------------
        'Passo 4 - Seleção do Usuário
        '-------------------------------------------------------------------------------------
        bRetorno = cUnear.AcessarListaDeSelecoes
        'Call RegistrarLogEtapa("ListaDeSelecao", bRetorno)
        '-------------------------------------------------------------------------------------
        'Upload do arquivo de seleção
        '-------------------------------------------------------------------------------------
        strArquivoSelecao = PegarArquivoLista(Access.Nz(rsListaDisparo!ArquivoSelecaoUsuarios.value), VBA.Date())
        strNomeSelecao = strCod & "UPLOAD_" & rsListaDisparo!TipoCampanha.value & "_" & AuxFileSystem.getNomeBase(strArquivoSelecao)
        
        Call RegistraStatus(Mensagem:="Efetua o upload do arquivo de seleção..." & strNomeSelecao, LogTxt:=strArquivoLogDetalhes, Progresso:=True)
        Call RegistraStatus(Mensagem:="Arquivo : " & strArquivoSelecao, LogTxt:=strArquivoLogDetalhes, Progresso:=True)
        
        strSelecao = cUnear.UploadSelecao(strArquivoSelecao, strNomeSelecao)
        If strSelecao <> "" Then GoTo SelecaoOK
        'Apos o upload, um numero de seleção é gerado.
        'Caso haja problema de atualização, não é possivel capturar o numero gerado.
        'Assim, é necessário forçar a atualização.
        If strSelecao = "" Then
            Call RegistraStatus(Mensagem:="Códio da seleção não recuperado após Upload. Tenta algumas opções a atualização", LogTxt:=strArquivoLogDetalhes, Progresso:=True)
        End If
        'Tenta recupera o código da seleção
        If strSelecao = "" Then
            Call RegistraStatus(Mensagem:="1º Tentativa - Recarrega a lista de seleções. Volta para a lista de campanhas e depois a de seleções", LogTxt:=strArquivoLogDetalhes, Progresso:=True)
            Call cUnear.AcessarListaDeCampanhas
            Call cUnear.AcessarListaDeSelecoes
            strSelecao = cUnear.PegaNumeroSelecao(strNomeSelecao)
        End If
        
        'Caso ainda sim não conseguiou...Sair do Sistema e volta para atualizar a lista
        If strSelecao = "" Then
            Call RegistraStatus(Mensagem:="2º Tentativa - Sai do Sistema e efetua logon novamente", LogTxt:=strArquivoLogDetalhes, Progresso:=True)
            'Sai do sistema, depois volta
            Call SairEEntrar(strArquivoLogDetalhes)
        End If
        Call cUnear.AcessarListaDeSelecoes
        strSelecao = cUnear.PegaNumeroSelecao(strNomeSelecao)

SelecaoOK:
        If strSelecao <> "" Then
            Call RegistraStatus(Mensagem:="Seleção Criada com sucesso! Código : " & strSelecao, LogTxt:=strArquivoLogDetalhes, Progresso:=True)
            Call AtualizaListaDisparo(intID, "IDSelecao", strSelecao)
        Else
            'Call RegistrarLogEtapa("UploadSelecao", Array(2, "Erro"))
            Call RegistraStatus(Mensagem:="Erro ao criar a seleção. Código Ainda não recuperado. Aguarda atualização manual por 15 segundos", LogTxt:=strArquivoLogDetalhes, Progresso:=True)
            GoTo Proxima
            statusFinal = "Dispario : #" & VBA.Format(intID, "00") & " | Codigo Campanha Unear : ERRO NO UPLOAD DA SELEÇÃO. Arquivo : " & strNomeSelecao
        End If

        strNomeCampanha = strCod & Nz(rsListaDisparo!Campanha.value, "CAMPANHA_")
        Call RegistraStatus(Mensagem:="Preparando para criar a nova campanha com nome : " & strNomeCampanha, LogTxt:=strArquivoLogDetalhes, Progresso:=True)
        
        '-------------------------------------------------------------------------------------
        'Passo 6 - Cria a campanha
        '-------------------------------------------------------------------------------------
        strCampanha = cUnear.CriarCampanha(strSelecao, strNomeCampanha, Nz(rsListaDisparo!Assunto.value), Nz(rsListaDisparo!ModeloEmail.value))(1)

        If strCampanha <> "" Then
            
            Call RegistraStatus(Mensagem:="Nova campanha '" & strNomeCampanha & "' como Numero : " & strCampanha, LogTxt:=strArquivoLogDetalhes, Progresso:=True)
            
            Call AtualizaListaDisparo(intID, "IDCampanha", strCampanha)
            Call AtualizaDadosCampanha(strCampanha, "IDCampanha", strCampanha)
            Call AtualizaDadosCampanha(strCampanha, "Campanha", strNomeCampanha)
            Call AtualizaDadosCampanha(strCampanha, "TipoCampanha", rsListaDisparo!TipoCampanha.value)
            Call AtualizaDadosCampanha(strCampanha, "SelecaoUsuarios", strSelecao & "-" & strNomeSelecao)
            Call AtualizaDadosCampanha(strCampanha, "Criacao", -1)
            '-------------------------------------------------------------------------------------
            'Passo 7 - Testa a campanha
            '-------------------------------------------------------------------------------------
            Call RegistraStatus(Mensagem:="****** TESTE INICIADO ***** ", LogTxt:=strArquivoLogDetalhes, Progresso:=True)
            bRetorno = TestarCampanha(strCampanha, strArquivoLogDetalhes)
            If bRetorno(0) = 1 Then
                statusTeste = "OK"
            Else
                statusTeste = "NÃO TESTOU - TIMEOUT"
            End If
            Call RegistraStatus(Mensagem:="Processo de Teste finalizado. Status : " & statusTeste & " Retorno : " & bRetorno(0) & " - " & bRetorno(1), LogTxt:=strArquivoLogDetalhes, Progresso:=True)
            Call RegistraStatus(Mensagem:="****** TESTE FINALIZADO ***** ", LogTxt:=strArquivoLogDetalhes, Progresso:=True)
            '-------------------------------------------------------------------------------------
            'Passo 8 - Executa a campanha
            '-------------------------------------------------------------------------------------
            Call RegistraStatus(Mensagem:="****** LIBERAÇÃO (EXECUÇÃO) INICIADA ***** ", LogTxt:=strArquivoLogDetalhes, Progresso:=True)
            
            bRetorno = ExecutarCampanha(strCampanha, intID, strArquivoLogDetalhes)
            
            If bRetorno(0) = 1 Then
                statusExecucao = "OK"
                bStatus = True
            Else
                Call arrCampanhas.Add(strCampanha & ";" & intID, "id_" & strCampanha)
                statusExecucao = "#NAO EXECUTADO"
            End If
            Call RegistraStatus(Mensagem:="Processo de LIBERAÇÃO finalizado. Status : " & statusTeste & " Retorno : " & bRetorno(0) & " - " & bRetorno(1), LogTxt:=strArquivoLogDetalhes, Progresso:=True)
            Call RegistraStatus(Mensagem:="****** LIBERAÇÃO (EXECUÇÃO) FINALIZADO ***** ", LogTxt:=strArquivoLogDetalhes, Progresso:=True)
            Call RegistraStatus(Mensagem:=" CAMPANHA :  '" & strCampanha & " - " & strNomeCampanha & "' PROCESSADA COM SUCESSO NO SISTEMA ###", LogTxt:=strArquivoLogDetalhes, Progresso:=True)
            Call AtualizaDadosCampanha(strCampanha, "Selecionar", -1)
            statusFinal = "DISPARO : #" & VBA.Format(intID, "00") & " | Codigo Campanha Unear : '" & strCampanha & " - " & strNomeCampanha & " | TESTE : " & statusTeste & " | LIBERAÇÃO : " & statusExecucao
        Else
            Call RegistraStatus(Mensagem:="#Erro ao criar a Nova campanha '" & strNomeCampanha & "'", LogTxt:=strArquivoLogDetalhes, Progresso:=True)
            'Call RegistrarLogEtapa("CriarCampanha", Array(2, "Erro"))
            statusFinal = "Dispario : #" & VBA.Format(intID, "00") & " | Codigo Campanha Unear : Erro, Campanha não foi criada no Unear"
            strCampanha = 0
        End If
        
        Call RegistraStatus(Mensagem:="RESULTADO >  " & statusFinal & " | TimerID : " & sTimerID & " | JobID : " & sJobID, LogTxt:=strArquivoLogDetalhes, Progresso:=True)
        
Proxima:
        Call AtualizaDadosCampanha(strCampanha, "EventoID", COD_DISPARO)
        Call AtualizaDadosCampanha(strCampanha, "TimerID", sTimerID)
        Call RegistrarLogEvento(COD_DISPARO, statusFinal, strArquivoLogDetalhes, strLogCampanhas, True, sTimerID, sJobID, "#" & VBA.Format(intID, "00"), bStatus)
        
        Call RegistraStatus(Mensagem:=VBA.String(120, "-"), LogTxt:=strArquivoLogDetalhes, Progresso:=True, IncluirDataHoraLog:=False)
        Call RegistraStatus(Mensagem:="FIM EVENTO : #" & VBA.Format(intID, "00") & "-" & COD_DISPARO, LogTxt:=strArquivoLogDetalhes, Progresso:=True, IncluirDataHoraLog:=False)
        Call RegistraStatus(Mensagem:=VBA.vbTab & VBA.vbTab & VBA.String(25, "."), LogTxt:=strArquivoLogDetalhes, Progresso:=True, IncluirDataHoraLog:=False)
        Call RegistraStatus(Mensagem:=VBA.vbTab & VBA.vbTab & "CAMPANHA     : " & VBA.IIf(strCampanha <> "", strCampanha, "Não criou"), LogTxt:=strArquivoLogDetalhes, Progresso:=True, IncluirDataHoraLog:=False)
        Call RegistraStatus(Mensagem:=VBA.vbTab & VBA.vbTab & "TESTE        : " & VBA.IIf(statusTeste <> "", statusTeste, "Não disponivel"), LogTxt:=strArquivoLogDetalhes, Progresso:=True, IncluirDataHoraLog:=False)
        Call RegistraStatus(Mensagem:=VBA.vbTab & VBA.vbTab & "LIBERAÇÃO    : " & VBA.IIf(statusExecucao <> "", statusExecucao, "Não disponivel"), LogTxt:=strArquivoLogDetalhes, Progresso:=True, IncluirDataHoraLog:=False)
        Call RegistraStatus(Mensagem:=VBA.vbTab & VBA.vbTab & "STATUS       : " & VBA.IIf(statusExecucao = "OK", "SUCESSO, CAMPANHA LIBERADA", "NÃO ENVIOU"), LogTxt:=strArquivoLogDetalhes, Progresso:=True, IncluirDataHoraLog:=False)
        
        Call RegistraStatus(Mensagem:=cstr_ProcedureName & " > TEMPO TOTAL  : " & PegaTempoDecorrido(VBA.Now - dtProcesso), LogTxt:=strArquivoLogDetalhes, Progresso:=True, IncluirDataHoraLog:=False)
        
        Call RegistraStatus(Mensagem:=VBA.vbTab & VBA.vbTab & VBA.String(25, "."), LogTxt:=strArquivoLogDetalhes, Progresso:=True, IncluirDataHoraLog:=False)
        Call RegistraStatus(Mensagem:=VBA.String(120, "-"), LogTxt:=strArquivoLogDetalhes, Progresso:=True, IncluirDataHoraLog:=False)
        Call RegistraStatus(Mensagem:="#" & VBA.vbNewLine, LogTxt:=strArquivoLogDetalhes, Progresso:=True, IncluirDataHoraLog:=False)
        
        rsListaDisparo.MoveNext
        'Stop 'Analisar
    Loop
    
    Call LimparEtapas
    'Se foi passado um TimerID, realiza novamente o Teste/Execução das campanhas pendentes com esses status
    If sTimerID <> "" Then
        If arrCampanhas.count > 0 Then
            dtProcesso = VBA.Timer
            Call RegistraStatus(Mensagem:="$", LogTxt:=strArquivoLogDetalhes, Progresso:=True)
            Call RegistraStatus(Mensagem:=VBA.String(120, "*"), LogTxt:=strArquivoLogDetalhes, Progresso:=True)
            Call RegistraStatus(Mensagem:="Job ID " & sJobID & " Foi finalizado, porem há campanhas que não foram executadas", Progresso:=True, LogTxt:=strArquivoLogDetalhes)
            Call RegistraStatus(Mensagem:="Iniciando o TESTE / EXECUÇÃO das " & arrCampanhas.count & "(s) Campanhas pendentes !", Progresso:=True, LogTxt:=strArquivoLogDetalhes)
            For i = 1 To arrCampanhas.count
                Call RegistraStatus(Mensagem:="Campanha : " & VBA.CStr(VBA.Split(arrCampanhas(i), ";")(0)) & " - Reiniciando LIBERAÇÃO", Progresso:=True, LogTxt:=strArquivoLogDetalhes)
                Call mBL_CRMUnear.ExecutarCampanha(VBA.CStr(VBA.Split(arrCampanhas(i), ";")(0)), VBA.CLng(VBA.Split(arrCampanhas(i), ";")(1)), strArquivoLogDetalhes)
            Next i
            Call RegistraStatus(Mensagem:="Tempo de processamento : " & PegaTempoDecorrido(VBA.Now - dtProcesso), Progresso:=True, LogTxt:=strArquivoLogDetalhes)
            Call RegistraStatus(Mensagem:=VBA.String(120, "*"), LogTxt:=strArquivoLogDetalhes, Progresso:=True, IncluirDataHoraLog:=False)
        End If
        Call mBL_CRMUnear.AtualizarStatusDaLista(1)
    End If
Fim:
    
    Call RegistraStatus(Mensagem:=VBA.String(120, "."), Progresso:=True, LogTxt:=strArquivoLogDetalhes)
    Call RegistraStatus(Mensagem:=cstr_ProcedureName & " > PROCESSO FINALIZADO ! Tempo Total : " & PegaTempoDecorrido(VBA.Now - dtSartRunProc), Progresso:=True, LogTxt:=strArquivoLogDetalhes)
    Call RegistraStatus(Mensagem:=VBA.String(120, "."), Progresso:=True, LogTxt:=strArquivoLogDetalhes)
    
    Call showMessage("Geração das campanhas finalizado !", , Fim, , , , Icone_Sucesso, 3)
    On Error GoTo 0
    Exit Sub

MacroCampanhaCRM_Unear_Error:
    If VBA.Err <> 0 Then
        lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
        bStatus = False
        statusFinal = "Dispario : #" & VBA.Format(intID, "00") & " | Codigo Campanha Unear : '" & strCampanha & " - " & strNomeCampanha & " | TESTE : " & statusTeste & " | LIBERAÇÃO : " & statusExecucao
        Call RegistraStatus(Mensagem:=" ------ ERROR ------ ", LogTxt:=strArquivoLogDetalhes, Progresso:=True)
        Call RegistraStatus(Mensagem:="ERROR - " & strErrorMessagem, LogTxt:=strArquivoLogDetalhes, Progresso:=True)
        Call RegistraStatus(Mensagem:=statusFinal, LogTxt:=strArquivoLogDetalhes, Progresso:=True)
        Call RegistraStatus(Mensagem:=" ------ ERROR ------ ", LogTxt:=strArquivoLogDetalhes, Progresso:=True)
    End If
    GoTo Proxima:
    'Debug Mode
    Resume
End Sub

' ----------------------------------------------------------------
' Procedure Name: PegarArquivoLista
' Purpose: Ajusta o nome do arquivo de acordo com variáveis
' Procedure Kind: Function
' Procedure Access: Public
' Parameter strEndereco (String): Endereço do arquivo de entrada
' Parameter dtRef (Date): Data de Referencia
' Author: Adelson
' Date: 04/11/2017
' ----------------------------------------------------------------
Function PegarArquivoLista(strEndereco As String, dtRef As Date)
    '----------------------------------------------------------------------------------------------------
    On Error GoTo TratarErro: Const cstr_ProcedureName As String = "TFWCliente.mBL_CRMUnear.PegarArquivoLista"
    '------------------------------------------------------------------------------------------------
    Dim rsVar As Object
    Dim strEnderecoSaida As String
    strEnderecoSaida = strEndereco
    If VBA.InStr(strEndereco, "@PastaRaiz") > 0 Then strEnderecoSaida = VBA.Replace(strEndereco, "[@PastaRaiz]", PastaRaizListas())
    If VBA.InStr(strEndereco, "@MMYYYY") > 0 Then strEnderecoSaida = VBA.Replace(strEnderecoSaida, "[@MMYYYY]", VBA.Format(dtRef, "MMYYYY"))
    If VBA.InStr(strEndereco, "@YYYYMMDD") > 0 Then strEnderecoSaida = VBA.Replace(strEnderecoSaida, "[@YYYYMMDD]", VBA.Format(dtRef, "YYYYMMDD"))
    PegarArquivoLista = strEnderecoSaida
    
    On Error GoTo 0
Fim:
    Exit Function
TratarErro:
    If VBA.Err <> 0 Then Call VBA.Err.Raise(VBA.Rnd(), cstr_ProcedureName & " - " & VBA.Erl() & ">" & VBA.Err.Description & VBA.vbCrLf)
    GoTo Fim:
    Resume
End Function

Function ValidaParametros(rs As Object) As Collection
    '----------------------------------------------------------------------------------------------------
    On Error GoTo TratarErro: Const cstr_ProcedureName As String = "TFWCliente.mBL_CRMUnear.ValidaParametros"
    '------------------------------------------------------------------------------------------------
    Dim cF As New Collection
    Dim i As Integer
    Dim vInfoClick As Variant
    Dim strArquivoSelecao As String
    
    For i = 0 To rs.Fields.count - 1
        'Preenchimento
        Select Case rs.Fields(i).Name
            Case "TipoCampanha", "Campanha", "ArquivoSelecaoUsuarios", "ModeloEmail", "Assunto"
                If rs.Fields(i).Name = "ModeloEmail" And rs.Fields("TipoCampanha").value = "SMS" Then
                Else
                    If Access.Nz(rs.Fields(i).value) = "" Then cF.Add rs.Fields(i).Name & " - Preenchimento obrigatório !"
                    If rs.Fields(i).Name = "ArquivoSelecaoUsuarios" Then
                        strArquivoSelecao = PegarArquivoLista(Access.Nz(rs!ArquivoSelecaoUsuarios.value), VBA.Date())
                        If Not FileExists(strArquivoSelecao) Then
                            cF.Add rs.Fields(i).Name & " - Arquivo não localizado ! > " & Nz(rs!ArquivoSelecaoUsuarios.value)
                        End If
                    End If
                End If
            Case Else
        End Select
    Next i
    vInfoClick = cUnear.BuscaInfoPonteito("MenuOpcoes"): If VBA.IsArray(vInfoClick) Then If vInfoClick(0) = 0 Then cF.Add "MenuOpcoes - Posição do Clique não configurado!"
    vInfoClick = cUnear.BuscaInfoPonteito("MenuSelecoes"): If VBA.IsArray(vInfoClick) Then If vInfoClick(0) = 0 Then cF.Add "MenuSelecoes - Posição do Clique não configurado!"
    vInfoClick = cUnear.BuscaInfoPonteito("BotaoUploadSelecao"): If VBA.IsArray(vInfoClick) Then If vInfoClick(0) = 0 Then cF.Add "BotaoUploadSelecao - Posição do Clique não configurado!"
    vInfoClick = cUnear.BuscaInfoPonteito("SelecaoModeloEmail_Editar"): If VBA.IsArray(vInfoClick) Then If vInfoClick(0) = 0 Then cF.Add "SelecaoModeloEmail_Editar - Posição do Clique não configurado!"
    vInfoClick = cUnear.BuscaInfoPonteito("SelecaoModeloEmail_Selecionar"): If VBA.IsArray(vInfoClick) Then If vInfoClick(0) = 0 Then cF.Add "SelecaoModeloEmail_Selecionar - Posição do Clique não configurado!"
    vInfoClick = cUnear.BuscaInfoPonteito("OpcaoSMS"): If VBA.IsArray(vInfoClick) Then If vInfoClick(0) = 0 Then cF.Add "OpcaoSMS - Posição do Clique não configurado!"
    vInfoClick = cUnear.BuscaInfoPonteito("Download_ClickSalvar"): If VBA.IsArray(vInfoClick) Then If vInfoClick(0) = 0 Then cF.Add "Download_ClickSalvar - Posição do Clique não configurado!"
    vInfoClick = cUnear.BuscaInfoPonteito("Download_ClickFechar"): If VBA.IsArray(vInfoClick) Then If vInfoClick(0) = 0 Then cF.Add "Download_ClickFechar - Posição do Clique não configurado!"
    Set ValidaParametros = cF
    
    On Error GoTo 0
Fim:
    Exit Function
TratarErro:
    If VBA.Err <> 0 Then Call VBA.Err.Raise(VBA.Rnd(), cstr_ProcedureName & " - " & VBA.Erl() & ">" & VBA.Err.Description & VBA.vbCrLf)
    GoTo Fim:
    Resume
End Function

Sub SairEEntrar(strArquivoLogDetalhes As String)
    Call RegistraStatus(Mensagem:="2º Tentativa - Saindo...", LogTxt:=strArquivoLogDetalhes, Progresso:=True)
    Call cUnear.Sair
    Call win.Wait
    Call RegistraStatus(Mensagem:="2º Tentativa - Entrando...", LogTxt:=strArquivoLogDetalhes, Progresso:=True)
    Call cUnear.Entrar(Unear_Usuario(), Unear_Senha())
    Call RegistraStatus(Mensagem:="2º Tentativa - Seleciona as campanhas para o sistema 'CRMRede'...", LogTxt:=strArquivoLogDetalhes, Progresso:=True)
    Call cUnear.AcessarListaDeCampanhas
    Call win.Wait
    Call RegistraStatus(Mensagem:="2º Tentativa - Voltando a lista de seleções...", LogTxt:=strArquivoLogDetalhes, Progresso:=True)
End Sub

Sub AtualizaListaDisparo(pID As Long, campo As String, valor As String)
    Call Conexao.ExecutarDDL(Conexao.PegarComandoSQLModelo("AtualizaDisparo", Nothing, "[@ID]", pID, "@campo", campo, "[@valor]", valor))
End Sub

Sub AtualizaDadosCampanha(pIDCampanha, campo As String, valor As String)
    If pIDCampanha <> "" Then
    With CurrentDb.OpenRecordset("SELECT * FROM tblCampanhas WHERE IDCampanha = " & pIDCampanha)
        If .EOF Then .addNew Else .edit
        .Fields("IDCampanha").value = pIDCampanha
        .Fields(campo).value = valor
        .Update
    End With
    Form_frmConsultaUnear.tblCampanhas_subformulário.Form.Requery
    End If
End Sub

Sub GravaLogEtapas(strEtapa As String, pStatus As String)
    With CurrentDb.OpenRecordset("tblEtapas")
        If .EOF Then .addNew Else .edit
        .Fields(strEtapa).value = pStatus
        .Update
    End With
End Sub

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : mBL_CRMUnear.AtualizarStatusDaLista()
' TIPO             : Sub
' DATA/HORA        : 31/08/2017 15:07
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Recupera a lista de status das campanhas
'---------------------------------------------------------------------------------------
Sub AtualizarStatusDaLista(Optional vArgs As Variant)
    Call CurrentDb.Execute("UPDATE tblCampanhas SET Selecionar = 0")
    Call CurrentDb.Execute("UPDATE tblCampanhas SET VisivelNaTela = 0")
    Call BaixarListaCampanhas(vArgs)
End Sub

Sub BaixarListaCampanhas(Optional vArgs As Variant)
    Dim rsListaDisparo As Object
    Dim strSelecao As String, strCampanha As String
    Dim bRetorno
    Dim intID As Long
    Dim intSegundos As Integer
    'Tratamento de Erro
    
    '---------------------------------------------------------------------------------------
1   On Error GoTo AtualizarStatusDaLista_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String:
    Const cstr_ProcedureName As String = "mBL_CRMUnear.BaixarListaCampanhas()"
    '---------------------------------------------------------------------------------------
    intSegundos = 2
2   Inicializar_Globais
    Dim strStatus As String
    Dim strTempTabela As String
    
3   If Not cUnear.BrowserEstaAberto Then
       MsgBox "É necessário que o navegador esteja aberto e com o Sistema Unear logado !", vbExclamation
       GoTo Fim
6   End If

    If Not cUnear.CRM_Conectado Then
       MsgBox "É necessário Conectado ao Unear !", vbExclamation
       GoTo Fim
    End If
    
    showMessage "Aguarde, Selecionando a lista atualizada (Top 50)."
    Call SelecionarTop50
    
    FecharProgresso

    strTempTabela = cUnear.DownloadLista("Campanha", "bsDataTable")

8   Set rsListaDisparo = CurrentDb.OpenRecordset("SELECT a.* FROM [" & strTempTabela & "] AS a WHERE a.Campanha Like 'TFWBOT*';")

9   Do While Not rsListaDisparo.EOF
        VBA.DoEvents
10      Call showMessage("Recuperando lista atualizada das campanhas '" & intID & "'", "Unear")
11      intID = rsListaDisparo!ID.value
12      strStatus = rsListaDisparo!status.value
13      Call AtualizaDadosCampanha(intID, "Status", strStatus)
        Call AtualizaDadosCampanha(intID, "Campanha", rsListaDisparo!Campanha.value)
        Call AtualizaDadosCampanha(intID, "TipoCampanha", rsListaDisparo!TIPO.value)
        Call AtualizaDadosCampanha(intID, "SelecaoUsuarios", rsListaDisparo![Seleção de Usuários].value)
        Call AtualizaDadosCampanha(intID, "Criacao", -1)
        Call AtualizaDadosCampanha(intID, "Selecionar", -1)
        Call AtualizaDadosCampanha(intID, "VisivelNaTela", -1)
14      Select Case strStatus
        Case "ATIVA"
15          Call AtualizaDadosCampanha(intID, "AcaoPendente", "EXECUTAR")
16      Case "CADASTRADA"
17          Call AtualizaDadosCampanha(intID, "AcaoPendente", "TESTAR")
18      Case "SMS TESTANDO", "TESTANDO"
19          Call AtualizaDadosCampanha(intID, "AcaoPendente", "EM TESTE")
20      Case "SMS LIBERADA", "GERANDO SMS (0%)"
21          Call AtualizaDadosCampanha(intID, "AcaoPendente", strStatus & " (EM PROCESSAMENTO)")
22      Case "ENCERRADA"
23          Call AtualizaDadosCampanha(intID, "AcaoPendente", "NENHUMA")
            Call AtualizaDadosCampanha(intID, "Inicio", rsListaDisparo!Início.value)
            Call AtualizaDadosCampanha(intID, "Fim", rsListaDisparo!Fim.value)
            Call AtualizaDadosCampanha(intID, "Download", VBA.CInt(VBA.Now > VBA.CDate(rsListaDisparo!Fim.value)))
24      Case Else
25          Call AtualizaDadosCampanha(intID, "AcaoPendente", strStatus)
26      End Select
27      If strStatus = "ENCERRADA" Then
28          Call AtualizaDadosCampanha(intID, "Execucao", -1)
29          Call AtualizaDadosCampanha(intID, "Teste", -1)
30      End If
31      rsListaDisparo.MoveNext
32  Loop
    
    rsListaDisparo.Close
    Set rsListaDisparo = Nothing
    
    Call ExcluiTabela(strTempTabela)

33  showMessage "Status atualizados com sucesso !", , Fim, , , , Icone_Sucesso, intSegundos

Fim:
34  On Error GoTo 0
35  Exit Sub

AtualizarStatusDaLista_Error:
36  If Err <> 0 Then
37      lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
38      Call VBA.Err.Raise(VBA.Err.Number, cstr_ProcedureName, VBA.vbNewLine & cstr_ProcedureName & " Em" & VBA.IIf(VBA.Erl() <> 0, " > Linha (" & VBA.Erl() & ")", "") & " > " & VBA.Err.Description)
39  End If
    GoTo Fim:
40  Resume    'Debug Mode
End Sub

Sub LimparEtapas()
    Inicializar_Globais
    Call Conexao.DeletarRegistros("Deletar_LogEtapas")
    Form_frmConsultaUnear.sfEtapas.Form.Requery
End Sub

Function Navegador_StatusAtual()
    Navegador_StatusAtual = pegaValor("Navegador_StatusAtual")
End Function

Function Navegador_URL()
    Navegador_URL = pegaValor("Navegador_URL")
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : mBL_CRMUnear.TestarCampanha()
' TIPO             : Function
' DATA/HORA        : 13/09/2017 10:02
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Executa o teste de uma campanha previamente criada
'---------------------------------------------------------------------------------------
Function TestarCampanha(strCampanha As String, strArquivoLogDetalhes As String)
          Dim bRetorno
          'Passo 7 - Seleciona a campanha criada
          'Tratamento de Erro
          '---------------------------------------------------------------------------------------
1         On Error GoTo TestarCampanha_Error
          Dim lngErrorNumber As Long, strErrorMessagem As String:
          Const cstr_ProcedureName As String = "mBL_CRMUnear.TestarCampanha()"
          '---------------------------------------------------------------------------------------

2         If Not cUnear.BrowserEstaAberto Then
3             bRetorno = Array(4, "O Browser não está aberto !")
              Call RegistraStatus(Mensagem:="Browser não está aberto. Teste Cancelado", LogTxt:=strArquivoLogDetalhes, Progresso:=True)
4             GoTo Fim
5         End If

6         If Not cUnear.CRM_Logado Then
              Call RegistraStatus(Mensagem:="Usuário não esta logado. Teste Cancelado", LogTxt:=strArquivoLogDetalhes, Progresso:=True)
7             bRetorno = Array(3, "Não está logado no site. Não é possivel Testar")
8             GoTo Fim
9         End If

          Call RegistraStatus(Mensagem:="Iniciando o teste da campanha criada", LogTxt:=strArquivoLogDetalhes, Progresso:=True)
          Call RegistraStatus(Mensagem:="Localiza a seleciona a campanha", LogTxt:=strArquivoLogDetalhes, Progresso:=True)

11        If cUnear.SelecionarCampanhaNaLista(strCampanha) Then
              Call RegistraStatus(Mensagem:="Campanha selecionada com sucesso, executa o clique no botão 'Testar'", LogTxt:=strArquivoLogDetalhes, Progresso:=True)
              'Passo 8 - Testa a campanha caso a mesma esteja selecionada
12            bRetorno = cUnear.TestarCampanha(strCampanha)
          End If
13        'Call RegistrarLogEtapa("TestarCampanha", bRetorno)
14        If bRetorno(0) = 1 Then
              Call RegistraStatus(Mensagem:="O TESTE foi enviado com sucesso", LogTxt:=strArquivoLogDetalhes, Progresso:=True)
15            Call AtualizaDadosCampanha(strCampanha, "Teste", -1)
16        Else
              Call RegistraStatus(Mensagem:="ERROR - Não foi possivel enviar o TESTE da campanha", LogTxt:=strArquivoLogDetalhes, Progresso:=True)
17            Call AtualizaDadosCampanha(strCampanha, "Teste", 0)
18        End If
20        TestarCampanha = bRetorno
Fim:
21        On Error GoTo 0
22        Exit Function

TestarCampanha_Error:
23        If Err <> 0 Then
24            lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
25            Call VBA.Err.Raise(VBA.Err.Number, cstr_ProcedureName, VBA.vbNewLine & cstr_ProcedureName & " Em" & VBA.IIf(VBA.Erl() <> 0, " > Linha (" & VBA.Erl() & ")", "") & " > " & VBA.Err.Description)
26        End If
          GoTo Fim:
27        Resume    'Debug Mode
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : mBL_CRMUnear.ExecutarCampanha()
' TIPO             : Function
' DATA/HORA        : 13/09/2017 10:03
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Executa o envio da campanha que esteja pronta (Status = ATIVA)
'---------------------------------------------------------------------------------------
Function ExecutarCampanha(strCampanha As String, intID As Long, strArquivoLogDetalhes As String)
          Dim bRetorno
          Dim strStatus As String

          'Tratamento de Erro
          '---------------------------------------------------------------------------------------
10        On Error GoTo ExecutarCampanha_Error
20        Dim lngErrorNumber As Long, strErrorMessagem As String:
          Const cstr_ProcedureName As String = "mBL_CRMUnear.ExecutarCampanha()"
          '---------------------------------------------------------------------------------------

30        If Not cUnear.BrowserEstaAberto Then
40            bRetorno = Array(4, "O Browser não está aberto !")
50            Call RegistraStatus(Mensagem:="Browser não está aberto. LIBERAÇÃO Cancelada", LogTxt:=strArquivoLogDetalhes, Progresso:=True)
60            GoTo Fim
70        End If

80        If Not cUnear.CRM_Logado Then
90            Call RegistraStatus(Mensagem:="Usuário não esta logado. LIBERAÇÃO Cancelada", LogTxt:=strArquivoLogDetalhes, Progresso:=True)
100           bRetorno = Array(3, "Não está logado no site. Não é possivel Testar")
110           GoTo Fim
120       End If

130       Call RegistraStatus(Mensagem:="Iniciando a liberação da campanha ja testada", LogTxt:=strArquivoLogDetalhes, Progresso:=True)
140       Call RegistraStatus(Mensagem:="Envia a campanha testada para Liberação", LogTxt:=strArquivoLogDetalhes, Progresso:=True)
          
          '---------------------------------------------------------------------------------------
150        bRetorno = cUnear.ExecutarCampanha(strCampanha)
          '---------------------------------------------------------------------------------------

160       If bRetorno(0) = 1 Then
170           win.Wait 3000
180           strStatus = cUnear.BuscaInfoCampanhaNaLista(strCampanha, "Status")
190           Do While strStatus = "Aguardando Liberação"
200               Call cUnear.SelecionarListaPorMenu
210               strStatus = cUnear.BuscaInfoCampanhaNaLista(strCampanha, "Status")
220           Loop
230           Call RegistraStatus(Mensagem:="Liberação enviada com sucesso", LogTxt:=strArquivoLogDetalhes, Progresso:=True)
240           Call AtualizaListaDisparo(intID, "Status", strStatus)
250           Call AtualizaListaDisparo(intID, "DataHoraCriacao", VBA.Now())
260           Call AtualizaDadosCampanha(strCampanha, "Teste", -1)
270           Call AtualizaDadosCampanha(strCampanha, "Execucao", -1)
280           Call AtualizaDadosCampanha(strCampanha, "DataHoraCriacao", VBA.Now())
290       Else
300           Call RegistraStatus(Mensagem:="Ocorreu um erro ao enviar a liberação", LogTxt:=strArquivoLogDetalhes, Progresso:=True)
310           strStatus = cUnear.BuscaInfoCampanhaNaLista(strCampanha, "Status")
320           Call AtualizaListaDisparo(intID, "Status", strStatus)
330       End If

340       Call AtualizaDadosCampanha(strCampanha, "Status", strStatus)
350       Call AtualizaDadosCampanha(strCampanha, "IDDisparo", VBA.CStr(intID))
          
360       ExecutarCampanha = bRetorno

Fim:
370       On Error GoTo 0
380       Exit Function

ExecutarCampanha_Error:
390       If Err <> 0 Then
400           lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
410           Call VBA.Err.Raise(VBA.Err.Number, cstr_ProcedureName, VBA.vbNewLine & cstr_ProcedureName & " Em" & VBA.IIf(VBA.Erl() <> 0, " > Linha (" & VBA.Erl() & ")", "") & " > " & VBA.Err.Description)
420       End If
430       GoTo Fim:
440       Resume    'Debug Mode
End Function


'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : mBL_CRMUnear.TestaCampanhasSelecionadas()
' TIPO             : Sub
' DATA/HORA        : 13/09/2017 10:13
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Dispara o teste/execução das campanhas selecionadas que estejam com o Status ATIVA
'---------------------------------------------------------------------------------------
Sub TestaExecutar_CampanhasSelecionadas(acao As eAcao)
    Dim rsListaDisparo As Object
    Dim strSelecao As String, strCampanha As String
    Dim bRetorno
    Dim intID As Long
    Dim intContagem As Integer

    'Tratamento de Erro
    '---------------------------------------------------------------------------------------
1   On Error GoTo TestaCampanhasSelecionadas_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String:
    Const cstr_ProcedureName As String = "mBL_CRMUnear.TestaCampanhasSelecionadas()"
    '---------------------------------------------------------------------------------------

2   Inicializar_Globais
    Dim strStatus As String

3   If Not cUnear.BrowserEstaAberto Then
4       MsgBox "É necessário que o navegador esteja aberto e com o Sistema Unear logado !", vbExclamation
5       Exit Sub
6   End If

    If Not cUnear.CRM_Conectado Then
        MsgBox "É necessário está conectado ao Unear para executar essa ação !", vbExclamation
        Exit Sub
    End If

7   If Not cUnear.CRM_TelaIndex = 2 Then cUnear.AcessarListaDeCampanhas
    
    
    If acao = eAcao.Testar Then
8       Set rsListaDisparo = Conexao.PegarRS("Pegar_ListaCampanhas_TestePendente")
    Else
        Set rsListaDisparo = Conexao.PegarRS("Pegar_ListaCampanhas_TesteExecutar")
    End If

9   Do While Not rsListaDisparo.EOF
10      If acao = eAcao.Testar Then
11          If rsListaDisparo!status.value = "CADASTRADA" Then
12              bRetorno = TestarCampanha(rsListaDisparo!IDCampanha.value, "")
13              If bRetorno(0) = 1 Then intContagem = intContagem + 1
14          End If
15      Else
16          If rsListaDisparo!status.value = "ATIVA" Then
17              bRetorno = ExecutarCampanha(rsListaDisparo!IDCampanha.value, rsListaDisparo!IDDisparo.value, "")
18              If bRetorno(0) = 1 Then intContagem = intContagem + 1
19          End If
20      End If
21      rsListaDisparo.MoveNext
22  Loop

23  If intContagem = 0 Then
24      If acao = "Teste" Then
25          showMessage "Nenhuma campanha selecionada disponível para teste !", , Fim, , , , Icone_Aviso, 3
26      Else
27          showMessage "Nenhuma campanha disponivel para ser executada !", , Fim, , , , Icone_Aviso, 3
28      End If
29  Else
30      If acao = "Teste" Then
31          showMessage intContagem & " - Campanhas enviadas para testes!", , Fim, , , , Icone_Sucesso, 3
32      Else
33          showMessage intContagem & " - foram enviadas (Executadas) !", , Fim, , , , Icone_Sucesso, 3
34      End If
        AtualizarStatusDaLista
35  End If
    
Fim:
36  On Error GoTo 0
37  Exit Sub

TestaCampanhasSelecionadas_Error:
38  If Err <> 0 Then
39      lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
40      Call VBA.Err.Raise(VBA.Err.Number, cstr_ProcedureName, VBA.vbNewLine & cstr_ProcedureName & " Em" & VBA.IIf(VBA.Erl() <> 0, " > Linha (" & VBA.Erl() & ")", "") & " > " & VBA.Err.Description)
41  End If
    GoTo Fim:
42  Resume    'Debug Mode

End Sub


Sub BaixarArquivosRetorno()
    'Tratamento de Erro
    '---------------------------------------------------------------------------------------
1   On Error GoTo BaixarArquivosRetorno_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String:
    Const cstr_ProcedureName As String = "cBL_Unear.BaixarArquivosRetorno()"
    '---------------------------------------------------------------------------------------
    Dim rsListaDisparo As Object
    
    Dim cUnear As New cBL_Unear
    
    Call cUnear.AcessarSite
    Call cUnear.Entrar(Unear_Usuario(), Unear_Senha())
    
    Set rsListaDisparo = Conexao.PegarRS("Pegar_ListaCampanhasSelecionadas")
    Do While Not rsListaDisparo.EOF
        Call showMessage("Baixando arquivos de retorno para '" & rsListaDisparo!Campanha.value & "'", "Unear", INICIO, 8)
        If rsListaDisparo!status.value = "ENCERRADA" Then
            If Not cUnear.CRM_TelaIndex = 2 Then Call cUnear.AcessarListaDeCampanhas
            Call cUnear.BaixarArquivosRetorno(rsListaDisparo!IDCampanha.value)
        End If
        rsListaDisparo.MoveNext
    Loop
    
    showMessage "Download comcluído !", , Fim, , , , Icone_Sucesso
Fim:
9   On Error GoTo 0
10  Exit Sub

BaixarArquivosRetorno_Error:
11  If Err <> 0 Then
12      lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
13      Call VBA.Err.Raise(VBA.Err.Number, cstr_ProcedureName, VBA.vbNewLine & cstr_ProcedureName & " Em" & VBA.IIf(VBA.Erl() <> 0, " > Linha (" & VBA.Erl() & ")", "") & " > " & VBA.Err.Description)
14  End If
    GoTo Fim:
    'Debug Mode
15  Resume

End Sub


Sub SelecionarTop50()
    Call cUnear.AcessarListaDeCampanhas
    Call cUnear.ClicarMouseNaPosicao("ComboSelecao")
    Call cUnear.ClicarMouseNaPosicao("ComboSelecao_Top50")
End Sub
