Attribute VB_Name = "ValidacaoArquivos"
Option Compare Database
'---------------------------------------------------------------------------------------
' PROJETO/MODULO   : TFWCliente.mLib_VBA
' TIPO             : Module
' DATA/HORA        : 05/10/2016 08:46
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Funções usadas para a validação dos arquivos a serem importados
'---------------------------------------------------------------------------------------
' + Historico de Revisão do Módulo
' **************************************************************************************
'   Versão    Data/Hora             Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      05/10/2016 08:46
'---------------------------------------------------------------------------------------

Sub NovaPreValidacaoArquivos()
    Call salvaValor("avisoPrevalidacao", "Escolha um diretório clicando em 'Selecionar'...")
    Call salvaValor("TelaPrevalidacao_FiltroStatus", "")
    Call salvaValor("imagemStatusPrevalidacao", "")
    Call salvaValor("AnalisarDiretorioPreValidacao", "")
    Call CurrentDb.Execute("DELETE FROM tblValidacaoArquivos")
    Call salvaValor("bEditarPasta", -1)
    Call AbrirFormulario("frmPrevalidacao", Access.acNormal, Access.acDialog)
End Sub

Function verificaResultadoValidacao(vResultValidacao)
    'Aviso que aparecerá
    Call salvaValor("avisoPrevalidacao", "Nenhuma mensagem")
    If vResultValidacao(0) = 0 Then
        'Abre a lista filtrado pelo status 'Erro'
        Call salvaValor("TelaPrevalidacao_FiltroStatus", "erro")
        Call salvaValor("avisoPrevalidacao", "Foram identificados arquivos que não estão de acordo com as regras de nomenclatura para importação na ferramenta." & VBA.vbNewLine & _
                                              "Corrija os nomes e tente novamente." & VBA.vbNewLine & _
                                              "Caso queira continuar, clique em Continuar para importar os arquivos que estão corretos")
        Call salvaValor("imagemStatusPrevalidacao", "aviso_erro")

    Else
        Call salvaValor("TelaPrevalidacao_FiltroStatus", "")
        Call salvaValor("avisoPrevalidacao", "Todos os arquivos foram validados e estão de acordo com as regras de nomenclatura da ferramenta." & VBA.vbNewLine & _
                                              "A importação pode continuar")
        Call salvaValor("imagemStatusPrevalidacao", "aviso_ok")
    End If
End Function

Function MontaFiltroArquivos()
    Dim strFiltro As String
    Dim arrArquivos As New cTFW_Array
    Dim i As Long
    
    Call arrArquivos.CopyFromArray(Conexao.PegarArray("Pegar_Relatorio"), True)
    
    For i = 0 To arrArquivos.RowsCount + 1
        Select Case LCase(arrArquivos.element(i, 8))
        Case "excel"
            strFiltro = strFiltro & VBA.LCase(arrArquivos.element(i, 1)) & ";*.xls"
        End Select
    Next i
        
    strFiltro = strFiltro & "|Todos os Arquvos;*.*"
    MontaFiltroArquivos = strFiltro
End Function

Function AdicionaArquivosSelecionadoParaImportacao(vArqs As Variant) As Variant
    For Each arquivo In vArqs
        Call AdicionarArquivoParaImportacao(VBA.Dir(CStr(arquivo)), VBA.CStr(arquivo), "")
    Next arquivo
End Function

Function preValidacaoRegraArquivos(vArqs As Variant, Optional bEditarPastas As Boolean = False) As Variant
    Dim arquivo
    Dim rsRegras As Object
    Dim NomeArquivo As String
    Dim vRegras
    Dim arrErros As New cTFW_Array
    Dim arrRegras As New cTFW_Array
    Dim arrOK As New cTFW_Array
    Dim regra, i As Long
    Dim bRegexEncontrado As Boolean

    Call Inicializar_Globais
    Call CurrentDb.Execute("DELETE FROM tblValidacaoArquivos")
    
    Call salvaValor("AnalisarDiretorioPreValidacao", PegarPasta(VBA.CStr(vArqs(0))))
    Call salvaValor("bEditarPasta", VBA.CInt(bEditarPastas))
    
    Call showMessage("Aguarde...")
    
    For Each arquivo In vArqs
        
        Call showMessage("Analisando..." & CStr(arquivo))
        
        NomeArquivo = VBA.Dir(CStr(arquivo))
        Call arrRegras.CopyFromArray(Conexao.PegarArray("Pegar_Relatorio"), True)
        
        bRegexEncontrado = False

        For i = 0 To arrRegras.RowsCount + 1
            If VBA.LCase(arrRegras.element(i, 7)) = "*" Then
                bRegexEncontrado = True
                Call AdicionarArquivoPrevalidacao(NomeArquivo, CStr(arquivo), arrRegras.element(i, 7), arrRegras.element(i, 12), "ok")
                Exit For
            Else
                If AuxTexto.IsLinhaMatch(VBA.LCase(NomeArquivo), VBA.LCase(arrRegras.element(i, 7))) Then
                    bRegexEncontrado = True
                    Call AdicionarArquivoPrevalidacao(NomeArquivo, CStr(arquivo), arrRegras.element(i, 7), arrRegras.element(i, 12), "ok")
                    Exit For
                End If
            End If
        Next i

        If Not bRegexEncontrado Then
            arrErros.AddElement CStr(arquivo)
            Call AdicionarArquivoPrevalidacao(NomeArquivo, CStr(arquivo), "", "#ERRO - Arquivo não identificado", "erro")
        Else
            arrOK.AddElement CStr(arquivo)
        End If
        
    Next arquivo
    
    If arrOK.RowsCount + 1 > 0 Then
        Call arrOK.Resize(arrOK.RowsCount)
    End If
    
    closeSplash

    If arrErros.RowsCount = 0 Then
        preValidacaoRegraArquivos = Array(-1, "PRE-VALIDAÇÃO - Todos os arquivos foram validados. Importação pode continuar...", arrErros, arrOK)
    Else
        preValidacaoRegraArquivos = Array(0, "PRE-VALIDAÇÃO - Foram encontrados (" & arrErros.RowsCount & ") arquivos com NOMES INVÁLIDOS na pré-validação." & VBA.vbNewLine & "Verifique e corrija os nomes dos arquivos conforme as regras exigidas...", arrErros, arrOK)
    End If
End Function

Sub AdicionarArquivoPrevalidacao(strArquivo As String, enderecoCompleto As String, RegraRegex As String, DescricaoRegra, status As String)
    On Error GoTo Erro
    Dim dtPeriodo As Date
    With CurrentDb.OpenRecordset("tblValidacaoArquivos")
        .addNew
        !arquivo.value = strArquivo
        !local.value = enderecoCompleto
        !RegraRegex.value = RegraRegex
        !DescricaoRegra.value = DescricaoRegra
        !status.value = status
        !PeriodoIdentificado.value = DeterminaDataPorRegex(strArquivo) 'vba.iif(dtPeriodo <> 0, UCase(VBA.Format(dtPeriodo, "MMM/YYYY")), "")
         Call AnexaArquivo("tblValidacaoArquivos", "", PegaEnderecoConfiguracoes() & "\" & status & ".ico", True, CurrentDb, !iconeStatus)
        .Update
    End With
    Exit Sub
Erro:
    VBA.MsgBox VBA.Error, VBA.vbCritical
    Resume Next
    Exit Sub
    Resume
End Sub

Sub AdicionarArquivoParaImportacao(strArquivo As String, enderecoCompleto As String, status As String)
    On Error GoTo Erro
    With CurrentDb.OpenRecordset("SELECT * FROM tblListaArquivosParaImportacao WHERE local = '" & enderecoCompleto & "'")
        If Not .EOF Then .edit Else .addNew
        !arquivo.value = strArquivo
        !local.value = enderecoCompleto
        !status.value = status
        .Update
    End With
    Exit Sub
Erro:
    VBA.MsgBox VBA.Error, VBA.vbCritical
    Resume Next: Exit Sub
    Resume
End Sub


Function SelecioanrVariosArquivosParaImportacao()
    Dim vArqs As Variant, vArqSelecionados As Variant
    Inicializar_Globais
    Call AbrirFormulario("frmListaArquivosParaImportacao", , acDialog)
    vArqSelecionados = Conexao.PegarArray(Conexao.PegarRS("Pegar_ArquivosSelecionadosParaImportacao", -1))
    If Not VBA.IsEmpty(vArqSelecionados) Then
    ReDim vArqs(0 To UBound(vArqSelecionados, 2))
    For i = 0 To UBound(vArqs, 1)
        vArqs(i) = vArqSelecionados(0, i)
    Next
    End If
    SelecioanrVariosArquivosParaImportacao = vArqs
End Function
