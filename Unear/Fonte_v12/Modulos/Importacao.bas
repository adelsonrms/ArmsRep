Attribute VB_Name = "Importacao"
Option Compare Database
Option Explicit

'********************VARIAVEIS - PROPRIEDADES********************
Private m_QtdArquivos                   As Long
Private m_DePara                        As DEPARAs
Private m_RelImportados                 As Object 'Scripting.Dictionary
Private m_ArquivosComErro               As Object 'Scripting.Dictionary
Private m_ArquivosImportados            As Object 'Scripting.Dictionary
Private m_ImportacaoInicio              As Date
Private m_ImportacaoFim                 As Date

'---------------------------------------------------------------------------------------
' Modulo....: VariaveisEConstantes / Módulo
' Rotina....: TipoRelatorio / Enum
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 09/08/2013
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Esse enumerado deve trabalhar em conjunto com a tabela tblRelatorios
'             como alguns relatórios serão criado em tempo de execução eu já reservei alguns
'             números para que os modelos fiquem sempre em cima assim fica mais fácil de dar
'             manutenção nessa tabela.
'             NOTA: Sempre substitiur o nome RESERVADOXX pelo nome do relatório para facilitar
'                   a identificação.
'---------------------------------------------------------------------------------------
Public Enum TipoRelatorio
    RESERVADO01 = -2
    Nenhum = -1
    XMLMensageria = 1
    BaseCSGD = 2
    BaseClientes = 3
End Enum

Public Enum TipoLeitura
    GenericoArrayExcel = 1
    Outro = 99
End Enum


'********************SUB_ROTINAS********************
'---------------------------------------------------------------------------------------
' Modulo....: AuxImportacao / Módulo
' Rotina....: Iniciar / Sub
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 16/08/2013
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina que concentra o inicio das importações dos arquivos
'             Essa rotina é a base de todas as importações.
'---------------------------------------------------------------------------------------
'REVISÕES :
'28/06/2016 - Adpatado trecho do código para utilizar a função CaixaDeDialogo() para seleção dos arquivos
'28/06/2016 - Para as chamadas que utilizam a tela de progresso, direcionar para funções reutilizaveis
'---------------------------------------------------------------------------------------
Public Sub Iniciar_Importacoes()
On Error GoTo TratarErro
Dim fso                 As Object
Dim fsoFile             As Object
Dim RelatorioAux        As TipoRelatorio
Dim intContador         As Integer
Dim InicioImpParcial    As Date
Dim InicioImpTotal      As Date
Dim vArquivos           As Variant
Dim vResultValidacao    As Variant
    
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    
    dtFirstTime = Time
    
    Call RegistraLog("--- Iniciar_Importacoes() - INICIANDO AS IMPORTAÇÕES ---------------------------------", logFile_Importacao)
    
    'Seleciona todos os arquivos
    'vArquivos = CaixaDeDialogo(msoFileDialogFilePicker, _
                               "Escolha os arquivos para a importação dos dados", _
                               True, _
                               "Importar Arquivos", _
                               "Arquivos de Dados;*.txt,*.xls,*.xml*,*.*db|Todos os Arquvos;*.*")
                               
        
    vArquivos = SelecioanrVariosArquivosParaImportacao()

    If Not VBA.IsEmpty(vArquivos) Then
    
        vResultValidacao = preValidacaoRegraArquivos(vArquivos)
        
        If vResultValidacao(0) = 0 Then
            Call verificaResultadoValidacao(vResultValidacao)
            'Abre o form
            Call AbrirFormulario("frmPrevalidacao", acNormal, acDialog)
            If pegaValor("Continuar_Importacao", 0) = 0 Then
                GoTo Fim
            End If
        End If
        vArquivos = vResultValidacao(3).Matriz
        
        Set fso = VBA.CreateObject("Scripting.FileSystemObject")
        
        QtdArquivos = 0
        
        Set Importacao.ArquivosComErro = VBA.CreateObject("Scripting.Dictionary")
        Set Importacao.ArquivosImportados = VBA.CreateObject("Scripting.Dictionary")
        Set RelImportados = VBA.CreateObject("Scripting.Dictionary")
        Set Excecoes.ErroDEPARA = VBA.CreateObject("Scripting.Dictionary")
        Set Excecoes.ErrosImportacao = VBA.CreateObject("Scripting.Dictionary")

        QtdArquivos = UBound(vArquivos) + 1
        
        Call AbrirTelaProgresso("Barra")
        
        Call Publicas.Inicializar_Globais(True)
        ImportacaoInicio = VBA.Now()
        For intContador = 0 To (Importacao.QtdArquivos - 1) Step 1
            If vArquivos(intContador) <> "" Then
                InicioImpParcial = VBA.Now()
                Call RegistraLog(Space(3) & (intContador + 1) & " de " & UBound(vArquivos) & " / Importando...." & vArquivos(intContador), logFile_Importacao)
                Set fsoFile = fso.GetFile(vArquivos(intContador))
                Call AuxForm.ExibirProgresso(intContador + 1, UBound(vArquivos), fsoFile.Name)
                RelatorioAux = 0
                Call Importacao.Importar_Relatorios(fso, fsoFile, RelatorioAux)
                Call InserirArquivoImportado(fsoFile.Name, fsoFile.Path, InicioImpParcial, RelatorioAux)
                Call RegistraLog(Space(3) & intContador & " de " & (Importacao.QtdArquivos - 1) & " / Finalizado ! " & PegaTempoDecorrido(Now - InicioImpParcial), logFile_Importacao)
                InicioImpParcial = 0
                Call Publicas.RemoverObjetosMemoria(fsoFile)
            End If
        Next intContador
        
        ImportacaoFim = VBA.Now()

        Call FecharTelaProgresso("Barra")
        
        Call RegistraLog(String(100, "-"), logFile_Importacao)
        Call RegistraLog(Space(3) & " Processo finalizado : Tempo total de " & PegaTempoDecorrido(Now - dtFirstTime), logFile_Importacao)
        Call RegistraLog("--- Iniciar_Importacoes() - Finalizado ! ---------------------------------", logFile_Importacao)
        Call RegistraLog("", logFile_Importacao)
                
        Call Access.DoCmd.openForm("frmLogImportacao", Access.AcFormView.acNormal, _
                                   WindowMode:=Access.AcWindowMode.acDialog)
    Else
        Call AuxMensagens.MessageBoxMaster("F014")
    End If
Fim:
    QtdArquivos = 0
    ImportacaoInicio = 0
    Call Publicas.RemoverObjetosMemoria(fsoFile, fso, Excecoes.ErroDEPARA, Excecoes.ErrosImportacao, _
                                       Importacao.ArquivosComErro, Importacao.ArquivosImportados)
    
    
On Error GoTo 0
Exit Sub
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "Importacao.Iniciar", Erl)
    Resume Next
    Resume
End Sub


Sub Teste()
    Debug.Print preValidacaoRegraArquivos(Array("D:\Data\Dev\Projetos\Tecnun\Framework\TFWApp_Modelo_Bradesco\Entradas\Base_Excel_Padrao_Tabela.xlsx", _
                                                "D:\Data\Dev\Projetos\Tecnun\Framework\TFWApp_Modelo_Bradesco\Entradas\teste.xlsx"))(1)
End Sub

'---------------------------------------------------------------------------------------
' Modulo....: AuxImportacao / Módulo
' Rotina....: InserirArquivoImportado / Sub
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 16/08/2013
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina para inserir o log dos arquivos importados com o seus respectivos
'             tempo
'---------------------------------------------------------------------------------------
Public Sub InserirArquivoImportado(ByVal NomeArquivo As String, ByVal CaminhoArquivo As String, _
                                   ByVal InicioImpParcial As Date, ByVal RelatorioAux As TipoRelatorio)
On Error GoTo TratarErro
Dim FimImpParcial       As Date
Dim chave               As String
Dim Motivo              As String
Dim ArquivoComErro      As Boolean
Dim arrAux              As Variant

    FimImpParcial = VBA.Now()
    
    If Importacao.ArquivosComErro.Exists(NomeArquivo) Then
        ArquivoComErro = True
        Motivo = Importacao.ArquivosComErro.item(NomeArquivo)
    End If

    chave = RelatorioAux & "|" & NomeArquivo
    If Not Importacao.ArquivosImportados.Exists(chave) Then
        arrAux = Array(RelatorioAux, NomeArquivo, CaminhoArquivo, InicioImpParcial, _
                       FimImpParcial, ArquivoComErro, Motivo, VBA.Environ("COMPUTERNAME"), _
                       VBA.Environ("USERNAME"))
        Call Importacao.ArquivosImportados.Add(chave, arrAux)
        Call Conexao.InserirRegistros("Insere_LogImportacao", arrAux)
    End If

On Error GoTo 0
Exit Sub
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "Importacao.InserirArquivoImportado", Erl)
End Sub

'---------------------------------------------------------------------------------------
' Modulo....: AuxImportacao / Módulo
' Rotina....: InserirArquivosComErro / Sub
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 14/08/2013
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina para inserir um log dos arquivos que apresentaram erros durante a
'             importação
'---------------------------------------------------------------------------------------
Public Sub InserirArquivosComErro(ByVal NomeArquivo As String, ByVal Motivo As String)
On Error GoTo TrataErro
    If Not Importacao.ArquivosComErro.Exists(NomeArquivo) Then
        Call Importacao.ArquivosComErro.Add(NomeArquivo, Motivo)
        Call Excecoes.TratarErro(Motivo, 9999, NomeArquivo, , False)
    End If
Exit Sub
TrataErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "Importacao.InserirArquivosComErro()", Erl)
End Sub

'---------------------------------------------------------------------------------------
' Modulo....: AuxImportacao / Módulo
' Rotina....: TratarRelatorioSaida / Sub
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 22/08/2013
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina para tratar o relatório que será exibido aos usuários
'---------------------------------------------------------------------------------------
Public Sub TratarRelatorioSaida(ByRef arrResumo As Variant, ByRef arrAnalitico As Variant, _
                                ByRef arrErros As Variant, ByRef arrDePara As Variant)
On Error GoTo TrataErro
Dim varKey          As Variant

    arrResumo = AuxArray.TransporMatriz(Conexao.PegarArray("Pega_LogImportacao_Resumo", _
                                   Importacao.ImportacaoInicio, Importacao.ImportacaoFim))

    arrAnalitico = AuxArray.TransporMatriz(Conexao.PegarArray("Pega_LogImportacao_Analitico", _
                                   Importacao.ImportacaoInicio, Importacao.ImportacaoFim))
                                   
    arrErros = AuxArray.TransporMatriz(Conexao.PegarArray("Pega_LogImportacao_Erros", _
                                   Importacao.ImportacaoInicio, Importacao.ImportacaoFim))

    arrDePara = AuxArray.TransporMatriz(Conexao.PegarArray("Pega_LogImportacao_De_Para", _
                                   Importacao.ImportacaoInicio))
Exit Sub
TrataErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "Importacao.TratarRelatorioSaida()", Erl)
End Sub

'********************FUNCOES********************
'---------------------------------------------------------------------------------------
' Modulo....: AuxImportacao / Módulo
' Rotina....: PegaTipoRelatorio / Function
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 09/08/2013
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina para determinar o tipo de relatório, primeiramente ele verifica se
'             o arquivo a ser aberto é do tipo Excel ou TXT e assim ele verifica
'---------------------------------------------------------------------------------------
Public Function Determinar_TipoRelatorio(ByRef fsoFile As Object, Optional ByRef xlsapp As Object, _
                                         Optional ByRef wbk As Object, Optional ByVal Senha As String = VBA.vbNullString) As TipoRelatorio
    On Error GoTo TratarErro
    Dim RelAux As TipoRelatorio
    Dim contador As Integer
    Dim NomeArquivo As String
    Dim relDet As RelatorioDetalhe

    RelAux = TipoRelatorio.Nenhum
    NomeArquivo = VBA.LCase(fsoFile.Name)

    With Relatorio
        For contador = 1 To .count Step 1
            With .item(contador)
                If AuxTexto.IsLinhaMatch(NomeArquivo, VBA.LCase(.Regex)) Then
                    RelAux = .ID
                    Exit For
                End If
            End With
        Next contador
    End With

    If Not VBA.IsMissing(wbk) And Not VBA.IsMissing(xlsapp) Then
        If Not RelAux = TipoRelatorio.Nenhum Then
            If AuxTexto.IsLinhaMatch(NomeArquivo, "[.](xls)\w*$") Then    'Siginifica que é um Excel
                Set xlsapp = VBA.CreateObject("Excel.Application")
                Set wbk = AuxExcel.AbrirWBK(xlsapp, fsoFile, Nothing, VBA.vbNullString, Senha)
                If wbk Is Nothing Then RelAux = TipoRelatorio.Nenhum
            End If
        End If
    End If

    Determinar_TipoRelatorio = RelAux

    On Error GoTo 0
    Exit Function
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "Importacao.PegaTipoRelatorio", Erl)
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : Importacao.Determinar_TipoLeitura()
' TIPO             : Function
' DATA/HORA        : 13/07/2016 16:37
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Determina o tipo de leitora do relatorio caso seja especificado na tabela tblRelatorio na coluna tipo Leitura
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      13/07/2016 16:37
'---------------------------------------------------------------------------------------
Public Function Determinar_TipoLeitura(ByRef fsoFile As Object) As TipoLeitura
On Error GoTo TratarErro
Dim RelAux          As TipoLeitura
Dim contador        As Integer
Dim NomeArquivo     As String
Dim relDet As RelatorioDetalhe
    
    RelAux = TipoRelatorio.Nenhum
    NomeArquivo = VBA.LCase(fsoFile.Name)
    
    With Relatorio
        For contador = 1 To .count Step 1
            With .item(contador)
                If .Regex <> VBA.vbNullString Then
                    If AuxTexto.IsLinhaMatch(NomeArquivo, VBA.LCase(.Regex)) Then
                        RelAux = Nz(DLookup("TipoLeitura", "tblRelatorios", "RelID=" & .ID), 99)
                        Exit For
                    End If
                End If
            End With
        Next contador
    End With
    
    Determinar_TipoLeitura = RelAux
    
On Error GoTo 0
Exit Function
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "Importacao.PegaTipoRelatorio", Erl)
End Function

'---------------------------------------------------------------------------------------
' Modulo....: Importacao / Módulo
' Rotina....: Importar_Relatorios / Sub
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 09/08/2013
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina que inicializa e verifica qual rotina de importação irá ser usada
'             de acordo com o arquivo a ser importado
'---------------------------------------------------------------------------------------

Public Sub Importar_Relatorios(ByRef fso As Object, _
                               ByRef fsoFile As Object, _
                               ByRef RelAux As TipoRelatorio)
1   On Error GoTo TratarErro
    Dim xlsapp As Object           'Excel.Application
    Dim wbk As Object

    Dim Tabela As Object
    Dim NomeTabela As String
    Dim dtRef As Date
    Dim dtBase As Date
    Dim arrDados As Variant
    Dim ArquivoValido As Boolean
    Dim MomentoInclusao As Date
    Dim bXML As Boolean
    Dim vRetorno As Variant
    
2   MomentoInclusao = VBA.Now
    'A rotina Determinar_TipoRelatorio, abaixo, uma vez que ela determine pelo nome do arquivo
    'qual é o relatório, ela mesma abre o arquivo, então, a abertura do wbk se dá na rotina chamada abaixo
3   RelAux = Importacao.Determinar_TipoRelatorio(fsoFile, xlsapp, wbk)

4   dtBase = EncontrarDataBase(fsoFile.Name)
5   Call salvaValor("dtBase", dtBase)

6   If Not RelAux = TipoRelatorio.Nenhum Then
7       ArquivoValido = True
        'ARMS 13/07/2016 - Incluido função que identifica o tipo de leitura.
        'Dessa forma podemos configurar na tabela 'tblRelatorios' e direcionar quais relatórios de acordo do tipo de leitura padrão
        'Por enquanto esta programado : Generico Excel
8       Select Case Importacao.Determinar_TipoLeitura(fsoFile)
        Case TipoLeitura.GenericoArrayExcel
9           arrDados = Pegar_Array_Excel(wbk:=wbk, NomePlan:=vbNullString, PL:=1, FechaWBK:=True, PC:=1, ColxlUP:=1)
10          If VBA.IsArray(arrDados) Then
11              Call Pegar_Tabela(Tabela, NomeTabela, RelAux, False, dtRef)
12              Call Importar_Rel_Generico(Tabela, NomeTabela, arrDados, RelAux, fsoFile.Name, fsoFile.Path, 1, 0)
13          Else
14              ArquivoValido = False
15          End If
16      Case Else
17          Select Case RelAux
            '----------------------------------------------------------------------------------------------------------------------------
            'Chama função de importação em modulo de regra de negocio especifico
            '----------------------------------------------------------------------------------------------------------------------------
            Case 1
            
22          End Select

23      End Select
24  Else
25      ArquivoValido = False
26  End If

27  If Not ArquivoValido Then
28      Call Importacao.InserirArquivosComErro(fsoFile.Name, "Arquivo inválido." & VBA.vbNewLine & _
                                                             "Por favor, verifique se o arquivo está com a nomeclatura correta e tente novamente.")
29  End If
Fim:
    ' Call AuxExcel.FecharWBK(wbk, False, xlsapp, True)
30  Call Publicas.RemoverObjetosMemoria(wbk, xlsapp)

31  On Error GoTo 0
32  Exit Sub
TratarErro:
33  Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "Importacao.Importar_Relatorios()", Erl)
34  Call Importacao.InserirArquivosComErro(fsoFile.Name, "Arquivo com erro." & VBA.vbNewLine & _
                                                         "Por favor, verifique se o arquivo está com a nomeclatura correta e tente novamente.")
35  GoTo Fim
36  Resume
End Sub

'---------------------------------------------------------------------------------------
' Modulo....: Importacao / Módulo
' Rotina....: CriarBackUPInformacoes / Sub
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 21/08/2013
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina que cria o backup da tabela
'---------------------------------------------------------------------------------------
Private Sub CriarBackUPInformacoes(ByVal chave As String)
On Error GoTo TratarErro
Dim Conn        As New ConexaoDB
Dim fso         As Object
Dim caminho     As String
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")

    caminho = fso.BuildPath(AuxTabela.PegarCaminhoBE, Relatorio.item(chave).NomeArquivo)
    Conn.StringConexao = "Provider=Microsoft.ACE.OLEDB.12.0;User ID=Admin;Data Source=" & caminho & _
                         ";Mode=Share Deny None;Persist Security Info=False"
    Call Conn.DeletarRegistros("Deletar_Temporario")
    Call Conn.InserirRegistros("Inserir_Temporario")

    Conn.DesConectar
    Call Publicas.RemoverObjetosMemoria(Conn)
On Error GoTo 0
Exit Sub
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "Importacao.CriarBackUPInformacoes", Erl)
End Sub

'---------------------------------------------------------------------------------------
' Modulo....: Importacao / Módulo
' Rotina....: InsereValoresPadroes / Sub
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 05/08/2014
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina para inserir valores padroes em campos fixos de acordo com o array de parametros
'             para tanto o array de parametros deve sempre ser acompanhado de 2 informações, sendo a
'             1 para o nome do campo e a 2 o valor
'---------------------------------------------------------------------------------------
Public Sub InsereValoresPadroes(ByRef Tabela As Object, _
                                ParamArray Parametros() As Variant)
On Error GoTo TratarErro
Dim contador        As Integer

    Parametros = AuxArray.Acertar_Array_Parametros(Parametros)
    With Tabela
        For contador = 0 To UBound(Parametros) Step 2
            .Fields(Parametros(contador)).value = Parametros(contador + 1)
        Next contador
    End With

On Error GoTo 0
Exit Sub
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "Importacao.InsereValoresPadroes", Erl)
End Sub


Private Function Pegar_Data(ByVal TextoData As String) As Date
On Error GoTo TratarErro
Dim DataAux     As String
Dim dtRef       As Date

    DataAux = AuxTexto.PegarTexto_Regex(VBA.UCase(TextoData), "\d{2,4}\s*[_]*\d{2}")
    If Not DataAux = VBA.vbNullString Then
        DataAux = VBA.Replace(VBA.Replace(DataAux, " ", ""), "_", "")
        If VBA.Len(DataAux) = 4 Then 'FORMATO AA MM
            dtRef = VBA.DateSerial(VBA.CInt(VBA.Left(DataAux, 2)), VBA.CInt(VBA.Right(DataAux, 2)), 1)
        ElseIf VBA.Len(DataAux) = 6 Then 'FORMATO AAAA MM
            dtRef = VBA.DateSerial(VBA.CInt(VBA.Left(DataAux, 4)), VBA.CInt(VBA.Right(DataAux, 2)), 1)
        End If
    End If
    Pegar_Data = dtRef
On Error GoTo 0
Exit Function
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "Importacao.Pegar_Data", Erl)
End Function

'---------------------------------------------------------------------------------------
' Modulo....: Importacao / Módulo
' Rotina....: PegarData / Sub
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 14/08/2013
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina para preencher os parametros Ano e Mes (BYREF) do relatorio 4980
'             ou qlq outra string no formato YYYYMM
'---------------------------------------------------------------------------------------
Private Function PegarData(ByVal valor As String, ByRef Ano As Long, ByRef Mes As Integer) As Date
On Error GoTo TratarErro

    If VBA.Len(valor) = 6 Then
        If VBA.IsNumeric(VBA.Left(valor, 4)) Then
            Ano = VBA.CLng(VBA.Left(valor, 4))
        End If
        If VBA.IsNumeric(VBA.Right(valor, 2)) Then
            Mes = VBA.CInt(VBA.Right(valor, 2))
        End If
    End If
    If Not Ano = 0 And Not Mes = 0 Then PegarData = VBA.DateSerial(Ano, Mes, 1)
On Error GoTo 0
Exit Function
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "Importacao.PegarData", Erl)
End Function

'---------------------------------------------------------------------------------------
' Modulo....: Importacao / Módulo
' Rotina....: Pegar_Array_Excel / Function
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 16/08/2013
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina que abre o arquivo em excel e pega seus dados.
'             PC = Primeira Coluna sendo a coluna inicial do array
'             ColuaSort é para classificar os dados de forma crescente antes de
'             inserir-los na array
'---------------------------------------------------------------------------------------
Public Function Pegar_Array_Excel(ByRef wbk As Object, ByVal NomePlan As String, _
                                   ByVal PL As Long, ByVal FechaWBK As Boolean, _
                                   Optional ByVal PC As Integer = 1, _
                                   Optional ByVal ColxlUP As Integer = 1, _
                                   Optional ByVal ColunaSort As Integer = 0, _
                                   Optional ByVal ColunaSort1 As Integer = 0) As Variant
On Error GoTo TrataErro
Dim sht             As Object
Dim UL              As Long 'Ultima Linha
Dim UC              As Long 'Ultima Coluna
Dim arrDados        As Variant
    
    'Direciona a manipulação das funções do Excel para classe cTFW_Excel
    
    If NomePlan = VBA.vbNullString Then
        Set sht = wbk.Worksheets(1)
    Else
        Set sht = AuxExcel.SetSheet(wbk, NomePlan)
    End If
    
    If sht Is Nothing Then
        Call Importacao.InserirArquivosComErro(wbk.Name, "Arquivo de importação inválido.")
        Call AuxExcel.FecharWBK(wbk)
        GoTo Fim
    End If

    With sht
        If Not .ProtectContents Then
            .Cells().Columns.Hidden = False
            .Cells().Rows.Hidden = False
        End If
        UL = .Cells(.Rows.count, ColxlUP).End(XlDirection.xlUp).row
        If PL = 0 Then PL = .Cells(UL, PC).End(XlDirection.xlUp).row - 2
        UC = .Cells(PL, .Columns.count).End(XlDirection.xlToLeft).Column
        If UL <= PL Or UC = 1 Then
            arrDados = Empty
        Else
            If ColunaSort > 0 Then
                With .Sort
                    With .SortFields
                        .Clear
                        .Add key:=sht.Range(sht.Cells(PL, ColunaSort), sht.Cells(UL, ColunaSort)), _
                             SortOn:=XlSortOn.xlSortOnValues, Order:=XlSortOrder.xlAscending, DataOption:=XlSortDataOption.xlSortNormal
                        If ColunaSort1 > 0 Then
                            .Add key:=sht.Range(sht.Cells(PL, ColunaSort1), sht.Cells(UL, ColunaSort1)), _
                            SortOn:=XlSortOn.xlSortOnValues, Order:=XlSortOrder.xlAscending, DataOption:=XlSortDataOption.xlSortNormal
                        End If
                    End With
                    .SetRange sht.Range(sht.Cells(PL, PC), sht.Cells(UL, UC))
                    .Header = XlYesNoGuess.xlYes
                    .MatchCase = False
                    .Orientation = Constants.xlTopToBottom
                    .SortMethod = XlSortMethod.xlPinYin
                    .Apply
                End With
                UL = .Cells(.Rows.count, ColxlUP).End(xlUp).row
            End If
            arrDados = GetArrayFromRange(.Range(.Cells(PL, PC), .Cells(UL, UC)), CellProperty.value)
        End If
    End With

    Call AuxForm.IncrementaBarraProgresso((5 / Importacao.QtdArquivos))
    Pegar_Array_Excel = arrDados
    
Fim:
    If FechaWBK Then
        Call AuxExcel.FecharWBK(wbk, False)
    Else
        Exit Function
    End If

    Call Publicas.RemoverObjetosMemoria(sht)
Exit Function
TrataErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "Importacao.Pegar_Array_Excel()", Erl)
    GoTo Fim
End Function

Public Function Pegar_Array_TXT(ByRef fsoFile As Object, _
                                 ByVal RodarComparacao As Boolean, _
                                 ByVal tpRelAux As TipoRelatorio) As Variant
On Error GoTo TratarErro
Dim arrRegex        As Variant
Dim linha           As String
Dim arrDados        As Variant

    arrRegex = Conexao.PegarArray("Pegar_Relatorio_Regex_PorID", tpRelAux)

    Open fsoFile For Input As #1
    Do While Not VBA.EOF(1)
        Line Input #1, linha
        linha = VBA.UCase(linha)
        If Not linha = VBA.vbNullString Then
            If RodarComparacao Then
                If Not AuxTexto.IsLinhaMatch(linha, arrRegex) Then
                    Call Preencher_Array_Dados(linha, arrDados)
                End If
            Else
                Call Preencher_Array_Dados(linha, arrDados)
            End If
        End If
    Loop
    Pegar_Array_TXT = arrDados
    Dim contador As Long
    Open "C:\Users\Jeff\Documents\tesssstee.txt" For Append As #2
    For contador = 0 To UBound(arrDados) Step 1
        Print #2, arrDados(contador)
    Next contador
    Close #2
Fim:
    Close #1
On Error GoTo 0
Exit Function
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "Importacao.Pegar_Array_TXT", Erl)
    Resume
    GoTo Fim
    Resume
End Function

'---------------------------------------------------------------------------------------
' Modulo....: Importacao / Módulo
' Rotina....: Preencher_Array_Dados / Sub
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 31/10/2014
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina para preencher o arrDados que é passado como referencia e utiliza
'             o redim preserve.
'---------------------------------------------------------------------------------------
Public Sub Preencher_Array_Dados(ByVal linha As String, ByRef arrDados As Variant)
On Error GoTo TratarErro
    If Not VBA.IsArray(arrDados) Then
        ReDim arrDados(0)
    Else
        ReDim Preserve arrDados(UBound(arrDados) + 1)
    End If
    arrDados(UBound(arrDados)) = linha
On Error GoTo 0
Exit Sub
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "Importacao.Preencher_Array_Dados", Erl)
End Sub

'---------------------------------------------------------------------------------------
' Modulo....: AuxLeitura / Módulo
' Rotina....: PreencheArray() / Sub
' Autor.....: Jefferson
' Contato...: jefferson.dantas@mondial.com.br
' Data......: ?/?/2012
' Empresa...: Mondial Tecnologia em Informática LTDA.
' Descrição.: Preenche um array de dados baseado em uma variável String
'---------------------------------------------------------------------------------------
Public Sub PreencheArray(ByVal strLinha As String, ByRef arrDados As Variant)
On Error GoTo TrataErro
    If Not VBA.IsArray(arrDados) Then
        ReDim arrDados(0)
    Else
        ReDim Preserve arrDados(UBound(arrDados) + 1)
    End If
    arrDados(UBound(arrDados)) = strLinha
Exit Sub
TrataErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxLeitura.PreencheArray()", Erl)
End Sub

Public Function Validar_Layout(ByRef Colunas As Object, ByRef Titulo As Object, _
                               ByVal NomeRelatorio As String, ByVal NomeArquivo As String, ByVal CaminhoDoArquivo As String)
On Error GoTo TratarErro
Dim Resultado       As Boolean
Dim coluna          As Variant

    If Colunas.count > 0 Then
        Resultado = True
        For Each coluna In Colunas.Keys
            If Not Titulo.Exists(coluna) Then
                Resultado = False
                Call Excecoes.Tratar_Log_Erros_De_Importacao(NomeRelatorio, NomeArquivo, CaminhoDoArquivo, 9999, _
                                                             "A Coluna Inexistente: " & coluna, 0, 0)
                'Exit For
            End If
        Next coluna
    End If
    Validar_Layout = Resultado
On Error GoTo 0
Exit Function
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "Importacao.Validar_Layout", Erl)
End Function

Private Sub Verificar_E_Inserir_Field(ByRef Tabela As Object, _
                                      ByVal FieldName As String, _
                                      ByVal valor As Variant)
On Error Resume Next 'Previsto, caso não exista o campo o valor não será inserido
    Tabela.Fields(FieldName).value = valor
End Sub

'---------------------------------------------------------------------------------------
' Procedure : PegarParametrodoTXT
' Author    : Jonathan
' Date      : 07/11/2014
' Purpose   : Passando um arquivo de texto buscar um texto Padrão com REGEX.
'---------------------------------------------------------------------------------------
Public Function PegarParametroDoTXT(ByRef fsoFile As Object, ByVal ExcluirReg As String, ParamArray Argumentos() As Variant) As String
Dim linha As String
Dim textoPegar As String
Dim arrExcluirRegex As Variant
Dim contador As Integer

    arrExcluirRegex = VBA.Split(ExcluirReg, "|")
    Argumentos = AuxArray.Acertar_Array_Parametros(Argumentos)
    On Error GoTo pegarNomeRelatorio_Error
    Open fsoFile For Input As #1

    Do While Not VBA.EOF(1)
        Line Input #1, linha
        linha = VBA.UCase(linha)
        If Not linha = VBA.vbNullString Then
            If AuxTexto.IsLinhaMatch(linha, Argumentos) Then
                textoPegar = AuxTexto.PegarTexto_Regex(linha, Argumentos)
                If VBA.IsArray(arrExcluirRegex) Then
                    For contador = 0 To UBound(arrExcluirRegex) Step 1
                        textoPegar = VBA.Replace(textoPegar, arrExcluirRegex(contador), "")
                    Next contador
                ElseIf Not ExcluirReg = VBA.vbNullString Then
                    textoPegar = VBA.Replace(textoPegar, ExcluirReg, "")
                End If
            End If
        End If
        If Not textoPegar = VBA.vbNullString Then
            Exit Do
        End If
    Loop
    If Not linha = VBA.vbNullString Then
        PegarParametroDoTXT = Trim(textoPegar)
    Else
        PegarParametroDoTXT = ""
    End If

Fim:
    Close #1
    On Error GoTo 0
    Exit Function

pegarNomeRelatorio_Error:
    If VBA.Err <> 0 Then Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "Importacao.pegarNomeRelatorio()", VBA.Erl)
    Exit Function
    Resume
End Function

''''---------------------------------------------------------------------------------------
'''' Modulo....: Importacao / Módulo
'''' Rotina....: TratarMatrizDados / Sub
'''' Autor.....: Fernando Fernandes
'''' Contato...: fernando@tecnun.com.br
'''' Data......: 08/04/2015
'''' Empresa...: Tecnun Tecnologia em Informática
'''' Descrição.: Rotina que trata os dados de uma matriz durante a importação de um arquivo
''''             com objetivo de deixá-la exatamente da forma que a base espera.
''''---------------------------------------------------------------------------------------
'''Public Sub TratarMatrizDados(ByRef arrDados As Variant, _
'''                             ByRef RelAux As TipoRelatorio, _
'''                             Optional dtBase As Date)
'''On Error GoTo TratarErro
'''Dim arrDestino  As Variant
'''Dim Titulos     As Variant
'''Dim Soma        As Double
'''Dim Mes         As String
'''Dim lin         As Long
'''Dim col         As Long
'''Dim dtRef       As Date
'''Dim linOrigem   As Long
'''Dim colOrigem   As Long
'''Dim linDestino  As Long
'''
'''    Select Case RelAux
'''        'Tratamento especifico da matriz de dados
'''        Case TipoRelatorio.RESERVADO01 'EXEMPLO
'''            'No exemplo elimina 2 linhas do array. As linhas de posição 10 e 20
'''            lin = LBound(arrDados, 1) + 1
'''            Do While lin <= UBound(arrDados, 1)
'''                VBA.DoEvents
'''                If lin = 10 Or lin = 20 Then
'''                    arrDados = RemoveRowFromArray(arrDados, lin)
'''                End If
'''                lin = lin + 1
'''            Loop
'''    End Select
'''    Call RemoverObjetosMemoria(Titulos, arrDestino)
'''
'''On Error GoTo 0
'''Exit Sub
'''TratarErro:
'''    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "Importacao.TratarMatrizDados", Erl)
'''End Sub



'---------------------------------------------------------------------------------------
' Modulo....: Importacao / Módulo
' Rotina....: BuscarDadosMultiPlanilhas / Sub
' Autor.....: Fernando Fernandes
' Contato...: fernando@tecnun.com.br
' Data......: 13/04/2015
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina para buscar dados de múltiplas planilhas com mesma estrutura, colocando
'             suas respectivas matrizes uma abaixo da outra, montando uma matrizona bem completa
'---------------------------------------------------------------------------------------
Private Sub BuscarDadosMultiPlanilhas(ByRef xlsapp As Object, _
                                      ByRef wbk As Object, _
                                      ByRef dtBase As Date, _
                                      ByRef fsoFile As Object, _
                                      ByRef arrDados As Variant, _
                                      Optional ByVal PL As Long = 2, _
                                      Optional ByVal PC As Long = 1, _
                                      Optional ByVal ColxlUP As Long = 1, _
                                      Optional ByVal ComPeriodo As Boolean = False)
On Error GoTo TratarErro
Dim arrDadosParcial As Variant
Dim wsh             As Object
Dim rngTitulo       As Object
Dim Periodo         As String

    If ComPeriodo Then
        With fsoFile
            Periodo = VBA.Mid(.Name, VBA.InStrRev(.Name, ".", -1) - 6, 6)
            dtBase = VBA.DateSerial(VBA.Left(Periodo, 4), VBA.Right(Periodo, 2), 1)
        End With
    End If
    
    Set rngTitulo = AuxExcel.GetUsedRange(wbk.Worksheets(1))
    With rngTitulo
        Set rngTitulo = .offset(PL - 2, 0).Resize(1, .Columns.count + 1)
    End With

    arrDados = GetArrayFromRange(rngTitulo, CellProperty.value)
    arrDados(1, UBound(arrDados, 2)) = "Planilha"
    
    For Each wsh In wbk.Worksheets
        If wsh.Visible = xlEnums.xlSheetVisible Then
            arrDadosParcial = Pegar_Array_Excel(wbk:=wbk, NomePlan:=wsh.Name, PL:=PL, FechaWBK:=False, PC:=PC, ColxlUP:=ColxlUP)
            arrDadosParcial = InsertColumnInArray(arr:=arrDadosParcial, position:=AfterLast, Content:=wsh.Name)
            arrDados = JuntarMatrizes(arrDados, arrDadosParcial, AuxArray.Orientation.Vertical)
        End If
    Next wsh
    
    Call AuxExcel.FecharWBK(wbk, False, xlsapp, True)
    Erase arrDadosParcial
    Call RemoverObjetosMemoria(arrDadosParcial, rngTitulo)

On Error GoTo 0
Exit Sub
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "Importacao.BuscarDadosMultiPlanilhas()", Erl)
End Sub

Public Function EncontrarPlanilha(ByRef wbk As Object, ByRef Texto As String) As String
On Error GoTo TratarErro
Dim Planilha As Object

    For Each Planilha In wbk.Worksheets
        If IsLinhaMatch(Planilha.Name, Texto) Then
            EncontrarPlanilha = Planilha.Name
            Exit For
        End If
    Next Planilha
    
On Error GoTo 0
Exit Function
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "Importacao.EncontrarPlanilha()", Erl)

End Function

Private Sub SepararPlanilhas(ByRef xlsapp As Object, _
                             ByRef fso As Object, _
                             ByRef fsoFile As Object, _
                             ByRef wbk As Object, _
                             ByRef dicPlanilhas As Object)
On Error GoTo TratarErro
Dim wsh             As Object
Dim TEMP            As String
Dim CaminhoArquivo  As String
Dim NomeArquivo     As String
Dim NomePlanilha    As String

    TEMP = VBA.Environ("Temp")
    
    If Not wbk Is Nothing Then
        For Each wsh In wbk.Worksheets

            NomePlanilha = wsh.Name
            NomeArquivo = NomePlanilha & "_" & fsoFile.Name
            CaminhoArquivo = TEMP & "\" & NomeArquivo
            If fso.FileExists(CaminhoArquivo) Then Call fso.DeleteFile(CaminhoArquivo, True)
            
            wsh.Copy
            With xlsapp
                With .Workbooks(.Workbooks.count)
                    .SaveAs FileName:=CaminhoArquivo
                    .Close SaveChanges:=False
                End With
            End With
            dicPlanilhas.Add NomePlanilha, CaminhoArquivo
        Next wsh
    End If
    Call RemoverObjetosMemoria(wsh)
On Error GoTo 0
Exit Sub
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "Importacao.SepararPlanilhas()", Erl)

End Sub

Public Function EncontrarColuna(ByVal arrDados As Variant, ByVal linha As Long, ByVal Conteudo As String) As Long
On Error GoTo TratarErro
Dim col As Long

    For col = LBound(arrDados, 2) To UBound(arrDados, 2) Step 1
        If arrDados(linha, col) = Conteudo Then
            EncontrarColuna = col
            Exit For
        End If
    Next col

On Error GoTo 0
Exit Function
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "Importacao.EncontrarColuna()", Erl)

End Function

'---------------------------------------------------------------------------------------
' Modulo....: Leitura / Módulo
' Rotina....: PegarArrayRelatorio / Function
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 16/08/2013
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina que abre o arquivo em excel e pega seus dados.
'             PC = Primeira Coluna sendo a coluna inicial do array
'             ColuaSort é para classificar os dados de forma crescente antes de
'             inserir-los na array
'---------------------------------------------------------------------------------------
Private Function PegarArrayRelatorio(ByRef wbk As Object, ByVal NomePlan As String, _
                                     ByVal PL As Long, ByVal FechaWBK As Boolean, _
                                     Optional ByVal PC As Integer = 1, _
                                     Optional ByVal ColxlUP As Integer = 1, _
                                     Optional ByVal ColunaSort As Integer = 0, _
                                     Optional ByVal ColunaSort1 As Integer = 0) As Variant
On Error GoTo TrataErro
Dim sht             As Object
Dim UL              As Long 'Ultima Linha
Dim UC              As Long 'Ultima Coluna
Dim NomeArquivo     As String
Dim arrDados        As Variant

    If NomePlan = VBA.vbNullString Then
        Set sht = wbk.Worksheets(1)
    Else
        Set sht = AuxExcel.SetSheet(wbk, NomePlan)
    End If
    If sht Is Nothing Then
        If wbk Is Nothing Then NomeArquivo = "Importação do Arquivo" Else NomeArquivo = wbk.Name
        Call Importacao.InserirArquivosComErro(NomeArquivo, "Arquivo de importação inválido.")
        Call AuxExcel.FecharWBK(wbk)
        GoTo Fim
    End If

    With sht
        If Not .ProtectContents Then
            .Cells().Columns.Hidden = False
            .Cells().Rows.Hidden = False
        End If
        UL = .Cells(.Rows.count, ColxlUP).End(-4162).row 'xlUp
        If PL = 0 Then PL = .Cells(UL, PC).End(-4162).row - 2 'xlUp
        UC = .Cells(PL, .Columns.count).End(-4159).Column  'xlToLeft
        If UL <= PL Or UC = 1 Then
            arrDados = Empty
        Else
            If ColunaSort > 0 Then
                With .Sort
                    With .SortFields
                        .Clear
                        'Excel.XlSortOn.xlSortOnValues = 0
                        'Excel.XlSortOrder.xlAscending = 1
                        'Excel.XlSortDataOption.xlSortNormal = 0
                        .Add key:=sht.Range(sht.Cells(PL, ColunaSort), sht.Cells(UL, ColunaSort)), _
                             SortOn:=0, Order:=1, DataOption:=0
                        If ColunaSort1 > 0 Then
                            .Add key:=sht.Range(sht.Cells(PL, ColunaSort1), sht.Cells(UL, ColunaSort1)), _
                                 SortOn:=0, Order:=1, DataOption:=0
                        End If
                    End With
                    .SetRange sht.Range(sht.Cells(PL, PC), sht.Cells(UL, UC))
                    .Header = 1         'Excel.XlYesNoGuess.xlYes = 1
                    .MatchCase = False
                    .Orientation = 1    'Excel.Constants.xlTopToBottom = 1
                    .SortMethod = 1     'Excel.XlSortMethod.xlPinYin = 1
                    .Apply
                End With
                UL = .Cells(.Rows.count, ColxlUP).End(-4162).row 'xlUp
            End If
            arrDados = .Range(.Cells(PL, PC), .Cells(UL, UC)).value
        End If
    End With

    Call AuxForm.IncrementaBarraProgresso((5 / Importacao.QtdArquivos))
    PegarArrayRelatorio = arrDados
Fim:
    If FechaWBK Then
        Call AuxExcel.FecharWBK(wbk, False)
    Else
        Exit Function
    End If

    Call Publicas.RemoverObjetosMemoria(sht)
Exit Function
TrataErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "Leitura.PegarArrayRelatorio()", Erl)
    GoTo Fim
End Function

Public Function EncontrarDataBase(ByVal NomeDoArquivo As String) As Date
On Error GoTo TratarErro
Dim Ponto   As Long
Dim txtData As String
Dim Ano     As Long
Dim Mes     As Long

    Ponto = VBA.InStrRev(NomeDoArquivo, ".", -1)
    
    If Ponto > 0 Then
        txtData = VBA.Mid(NomeDoArquivo, VBA.IIf(Ponto <= 6, 1, Ponto - 6), 6)
        If VBA.IsNumeric(txtData) Then
            If VBA.Left(txtData, 1) = 2 Then
                Ano = VBA.Left(txtData, 4)
                Mes = VBA.Right(txtData, 2)
                EncontrarDataBase = VBA.DateSerial(Ano, Mes, 1)
            ElseIf VBA.Mid(txtData, 3, 1) = 2 Then
                Ano = VBA.Right(txtData, 4)
                Mes = VBA.Left(txtData, 2)
                EncontrarDataBase = VBA.DateSerial(Ano, Mes, 1)
            End If
        End If
    End If
    
Fim:
On Error GoTo 0
Exit Function
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "Importacao.EncontrarDataBase()", Erl)
End Function


'---------------------------------------------------------------------------------------
' Modulo....: Importacao_XLS / Módulo
' Rotina....: Importar_Rel_Generico / Sub
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 03/01/2014
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina Generica para a importação de um relatorio
'---------------------------------------------------------------------------------------
Public Sub Importar_Rel_Generico_MDB(ByRef Tabela As Object, ByVal NomeTabela As String, _
                                     ByVal arrDados As Variant, ByVal RelAux As TipoRelatorio, _
                                     ByVal NomeArquivo As String, ByVal CaminhoDoArquivo As String, _
                                     ByVal LinhaInicial As Long, ByVal ColunaInicial As Integer, _
                                     ParamArray Parametros() As Variant)
On Error GoTo TratarErro
Dim Colunas             As Object 'Scripting.Dictionary
Dim Titulo              As Object 'Scripting.Dictionary
Dim TabelaAux           As String
Dim chave               As String
Dim ChaveProduto        As String
Dim ContLinhas          As Long
Dim Ano                 As Long
Dim ContCol             As Integer
Dim PosComErro          As Integer
Dim dtRef               As Date
Dim dtRefAnterior       As Date
Dim arrAux              As Variant
Dim arrAuxProd          As Variant
Dim Incremento          As Double
Dim UsuarioInclusao     As String
Dim MomentoInclusao     As Date
    
    UsuarioInclusao = Publicas.ChaveUsuario()
    MomentoInclusao = VBA.Now()
    
    If Tabela Is Nothing Then GoTo Fim
    Parametros = AuxArray.Acertar_Array_Parametros(Parametros)
    
    Set Titulo = AuxExcel.CriarTitulos(arrDados, 1)
    Set Colunas = Conexao.PegarDicionario("Pegar_RelatorioColunas", RelAux)
    
    If Not Validar_Layout(Colunas, Titulo, "Importação de Relatórios", NomeArquivo, CaminhoDoArquivo) Then
        Call Importacao.InserirArquivosComErro(NomeArquivo, "Layout do arquivo inválido.")
        GoTo Fim
    End If

    With Tabela
        Incremento = (75 / Importacao.QtdArquivos) / UBound(arrDados, 1)
        For ContLinhas = 2 To UBound(arrDados, 1) Step 1
            If (ContLinhas Mod 1000) = 0 Then Call AuxForm.IncrementaBarraProgresso(Incremento * 1000)
            .addNew
            If UBound(Parametros) >= 0 Then Call InsereValoresPadroes(Tabela, Parametros)
            For ContCol = LBound(arrDados, 2) To UBound(arrDados, 2) Step 1
                chave = RemoverQuebrasDeLinha(VBA.UCase(VBA.Trim(arrDados(1, ContCol))))
                
                If Colunas.Exists(chave) Then
                    PosComErro = ContCol
                    arrAux = VBA.Split(Colunas.item(chave), "|")
                    With .Fields(arrAux(0))
                        If .Type = 10 Then
                            .value = VBA.Trim(arrDados(ContLinhas, ContCol))
                            
                        ElseIf VBA.IsDate(arrDados(ContLinhas, ContCol)) And .Type = 8 Then
                            dtRef = VBA.CDate(arrDados(ContLinhas, ContCol))
                            .value = dtRef
                            
                        ElseIf VBA.IsNumeric(arrDados(ContLinhas, ContCol)) And _
                               (.Type = 5 Or _
                                .Type = 19 Or _
                                .Type = 7 Or _
                                .Type = 3 Or _
                                .Type = 6 Or _
                                .Type = 4) Then
                            .value = VBA.CDbl(arrDados(ContLinhas, ContCol))
                            
                        End If
                    End With
                End If
            Next ContCol

            'os campos abaixo TEM que existir em TODAS nossas tabelas SEMPRE.
            'Por isso não há teste pois deveremos garantir a existência deles, ao criar a tabela
            'FF-08/04/2015
            .Fields("MomentoInclusao").value = MomentoInclusao
            .Fields("UsuarioInclusao").value = UsuarioInclusao
            .Update
        Next ContLinhas
        
    End With
    If Importacao.RelImportados.Exists(NomeTabela) Then
        Call Importacao.RelImportados.Add(NomeTabela, Empty)
    End If
Fim:
    Call AuxTabela.FecharRecordSet(Tabela)
    Call Publicas.RemoverObjetosMemoria(Tabela)
    
On Error GoTo 0
Exit Sub
TratarErro:
    '=======================================================================================
    'Este tratamento é referente ao Tipo de Dados Incompatível dos valores
    '=======================================================================================
    If VBA.Err.Number = -2147352571 Or VBA.Err.Number = 13 Or _
       VBA.Err.Number = 2007 Or VBA.Err.Number = 3421 Then
       Call Excecoes.Tratar_Log_Erros_De_Importacao("Importação de Relatórios", NomeArquivo, CaminhoDoArquivo, VBA.Err.Number, _
       "Tipo de Dados Incorreto, Por favor, verifique o conteúdo da célula no Relatório Importado.", _
       (ContLinhas + LinhaInicial), (PosComErro + ColunaInicial))
       Resume Next 'Favor não remover esta linha
    End If
    '=======================================================================================
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "Importacao_XLS.Importar_Rel_Generico()")
    Call Importacao.InserirArquivosComErro(NomeArquivo, "Erro durante a importação")
    '=======================================================================================
    'Tratamento referente ao Rollback de Dados
    '=======================================================================================
    If Importacao.RelImportados.count > 0 Then
       If Importacao.RelImportados.Exists(NomeTabela) Then
          Call Importacao.RelImportados.Remove(NomeTabela)
       End If
    End If
    Call AuxTabela.DeletarRelatorio(NomeTabela, AuxTabela.PegarComplemento_Generico(dtRef:=dtRef))
    '=======================================================================================
    GoTo Fim
    Resume
End Sub

'---------------------------------------------------------------------------------------
' Modulo....: Importacao_TXT / Módulo
' Rotina....: Importar_Rel_Generico_TXT / Sub
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 03/01/2014
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina Generica para a importação de um relatorio
'---------------------------------------------------------------------------------------
Public Sub Importar_Rel_Generico_TXT(ByRef Tabela As Object, ByVal NomeTabela As String, _
                                     ByVal arrDados As Variant, ByVal RelAux As TipoRelatorio, _
                                     ByVal NomeArquivo As String, ByVal CaminhoDoArquivo As String, _
                                     ParamArray Parametros() As Variant)
On Error GoTo TratarErro
Dim Colunas             As Object 'Scripting.Dictionary
Dim Titulo              As Object 'Scripting.Dictionary
Dim TabelaAux           As String
Dim chave               As String
Dim ChaveProduto        As String
Dim ContLinhas          As Long
Dim Ano                 As Long
Dim ContCol             As Integer
Dim PosComErro          As Integer
Dim dtRef               As Date
Dim dtRefAnterior       As Date
Dim arrAux              As Variant
Dim arrAuxProd          As Variant
Dim Incremento          As Double
Dim coluna              As Variant
Dim Aux                 As Variant

    If Tabela Is Nothing Then GoTo Fim
    Parametros = AuxArray.Acertar_Array_Parametros(Parametros)
    Set Colunas = Conexao.PegarDicionario("Pegar_RelatorioColunas", RelAux)

    With Tabela
        Incremento = (75 / Importacao.QtdArquivos) / UBound(arrDados, 1)
        For ContLinhas = 2 To UBound(arrDados, 1) Step 1
            If (ContLinhas Mod 1000) = 0 Then Call AuxForm.IncrementaBarraProgresso(Incremento * 1000)
            .addNew
            If UBound(Parametros) >= 0 Then Call InsereValoresPadroes(Tabela, Parametros)
            For Each coluna In Colunas.Keys
                arrAux = VBA.Split(Colunas.item(coluna), "|")
                Aux = VBA.Trim(VBA.Mid(arrDados(ContLinhas), arrAux(2), arrAux(3)))
                
                If .Fields(coluna).Type = 10 Then
                    .Fields(coluna).value = Aux
                ElseIf VBA.IsDate(Aux) And .Fields(coluna).Type = 8 Then
                    .Fields(coluna).value = VBA.CDate(Aux)
                ElseIf VBA.IsNumeric(Aux) And _
                       (.Fields(coluna).Type = 5 Or _
                        .Fields(coluna).Type = 19 Or _
                        .Fields(coluna).Type = 7 Or _
                        .Fields(coluna).Type = 7 Or _
                        .Fields(coluna).Type = 4) Then
                    .Fields(coluna).value = VBA.CDbl(Aux)
                End If
            Next coluna
'            Call Verificar_E_Inserir_Field(Tabela, "Data_Hora", VBA.Now())
'            Call Verificar_E_Inserir_Field(Tabela, "Usuario", Publicas.ChaveUsuario)
            .Update
        Next ContLinhas
    End With
    If Importacao.RelImportados.Exists(NomeTabela) Then
        Call Importacao.RelImportados.Add(NomeTabela, Empty)
    End If
Fim:
    Call AuxTabela.FecharRecordSet(Tabela)
    Call Publicas.RemoverObjetosMemoria(Tabela)
On Error GoTo 0
Exit Sub
TratarErro:
    '=======================================================================================
    'Este tratamento é referente ao Tipo de Dados Incompatível dos valores
    '=======================================================================================
    If VBA.Err.Number = -2147352571 Or VBA.Err.Number = 13 Or _
       VBA.Err.Number = 2007 Or VBA.Err.Number = 3421 Then
       Call Excecoes.Tratar_Log_Erros_De_Importacao("Importação de Relatórios", NomeArquivo, CaminhoDoArquivo, VBA.Err.Number, _
       "Tipo de Dados Incorreto, Por favor, verifique o conteúdo da célula no Relatório Importado.", 0, 0)
       Resume Next 'Favor não remover esta linha
    End If
    '=======================================================================================
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "Importacao_TXT.Importar_Rel_Generico_TXT()")
    Call Importacao.InserirArquivosComErro(NomeArquivo, "Erro durante a importação")
    '=======================================================================================
    'Tratamento referente ao Rollback de Dados
    '=======================================================================================
    If Importacao.RelImportados.count > 0 Then
       If Importacao.RelImportados.Exists(NomeTabela) Then
          Call Importacao.RelImportados.Remove(NomeTabela)
       End If
    End If
    Call AuxTabela.DeletarRelatorio(NomeTabela, AuxTabela.PegarComplemento_Generico(dtRef:=dtRef))
    '=======================================================================================
    GoTo Fim
End Sub

'---------------------------------------------------------------------------------------
' Modulo....: Importacao_XLS / Módulo
' Rotina....: Importar_Rel_Generico / Sub
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 03/01/2014
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina Generica para a importação de um relatorio
'---------------------------------------------------------------------------------------
Public Sub Importar_Rel_Generico(ByRef Tabela As Object, ByVal NomeTabela As String, _
                                 ByVal arrDados As Variant, ByVal RelAux As TipoRelatorio, _
                                 ByVal NomeArquivo As String, ByVal CaminhoDoArquivo As String, _
                                 ByVal LinhaInicial As Long, ByVal ColunaInicial As Integer, _
                                 ParamArray Parametros() As Variant)
On Error GoTo TratarErro
Dim Colunas             As Object 'Scripting.Dictionary
Dim Titulo              As Object 'Scripting.Dictionary
Dim TabelaAux           As String
Dim chave               As String
Dim ChaveProduto        As String
Dim ContLinhas          As Long
Dim Ano                 As Long
Dim ContCol             As Integer
Dim PosComErro          As Integer
Dim dtRef               As Date
Dim dtRefAnterior       As Date
Dim arrAux              As Variant
Dim arrAuxProd          As Variant
Dim Incremento          As Double
Dim UsuarioInclusao     As String
Dim MomentoInclusao     As Date
Dim LinhaTitulos        As Long

    UsuarioInclusao = Publicas.ChaveUsuario()
    MomentoInclusao = VBA.Now()
    
    If Tabela Is Nothing Then GoTo Fim
    Parametros = AuxArray.Acertar_Array_Parametros(Parametros)
    
    If arrDados(LBound(arrDados, 1), LBound(arrDados, 2)) = VBA.vbNullString Then
        LinhaTitulos = LBound(arrDados, 1) + 1
    Else
        LinhaTitulos = LBound(arrDados, 1)
    End If
    Set Titulo = AuxExcel.CriarTitulos(arrDados, LinhaTitulos)
    
    Set Colunas = Conexao.PegarDicionario("Pegar_RelatorioColunas", RelAux)
    
    If Not Validar_Layout(Colunas, Titulo, "Importação de Relatórios", NomeArquivo, CaminhoDoArquivo) Then
        Call Importacao.InserirArquivosComErro(NomeArquivo, "Layout do arquivo inválido.")
        GoTo Fim
    End If

    With Tabela
        Incremento = (75 / Importacao.QtdArquivos) / UBound(arrDados, 1)
        For ContLinhas = LBound(arrDados, 1) + 1 To UBound(arrDados, 1) Step 1
            If (ContLinhas Mod 1000) = 0 Then Call AuxForm.IncrementaBarraProgresso(Incremento * 1000)
            .addNew
            If UBound(Parametros) >= 0 Then Call InsereValoresPadroes(Tabela, Parametros)
            For ContCol = LBound(arrDados, 2) To UBound(arrDados, 2) Step 1
                chave = RemoverQuebrasDeLinha(VBA.UCase(VBA.Trim(arrDados(LinhaTitulos, ContCol))))
                
                If Colunas.Exists(chave) Then
                    PosComErro = ContCol
                    arrAux = VBA.Split(Colunas.item(chave), "|")
                    With .Fields(arrAux(0))
                        If .Type = dao.DataTypeEnum.dbText Then
                            .value = VBA.Trim(arrDados(ContLinhas, ContCol))
                            
                        ElseIf VBA.IsDate(arrDados(ContLinhas, ContCol)) And .Type = dao.DataTypeEnum.dbDate Then
                            dtRef = VBA.CDate(arrDados(ContLinhas, ContCol))
                            .value = dtRef
                            
                        ElseIf VBA.IsNumeric(arrDados(ContLinhas, ContCol)) And _
                               (.Type = dao.DataTypeEnum.dbCurrency Or _
                                .Type = dao.DataTypeEnum.dbNumeric Or _
                                .Type = dao.DataTypeEnum.dbDouble Or _
                                .Type = dao.DataTypeEnum.dbInteger Or _
                                .Type = dao.DataTypeEnum.dbSingle Or _
                                .Type = dao.DataTypeEnum.dbLong) Then
                            .value = VBA.CDbl(arrDados(ContLinhas, ContCol))
                            
                        End If
                    End With
                End If
            Next ContCol

            'os campos abaixo TEM que existir em TODAS nossas tabelas SEMPRE.
            'Por isso não há teste pois deveremos garantir a existência deles, ao criar a tabela
            'FF-08/04/2015
            .Fields("MomentoInclusao").value = MomentoInclusao
            .Fields("UsuarioInclusao").value = UsuarioInclusao
            .Update
        Next ContLinhas
        
    End With
    If Importacao.RelImportados.Exists(NomeTabela) Then
        Call Importacao.RelImportados.Add(NomeTabela, Empty)
    End If
Fim:
    Call AuxTabela.FecharRecordSet(Tabela)
    Call Publicas.RemoverObjetosMemoria(Tabela)
    
On Error GoTo 0
Exit Sub
TratarErro:
    '=======================================================================================
    'Este tratamento é referente ao Tipo de Dados Incompatível dos valores
    '=======================================================================================
    If VBA.Err.Number = -2147352571 Or VBA.Err.Number = 13 Or _
       VBA.Err.Number = 2007 Or VBA.Err.Number = 3421 Then
       Call Excecoes.Tratar_Log_Erros_De_Importacao("Importação de Relatórios", NomeArquivo, CaminhoDoArquivo, VBA.Err.Number, _
       "Tipo de Dados Incorreto, Por favor, verifique o conteúdo da célula no Relatório Importado.", _
       (ContLinhas + LinhaInicial), (PosComErro + ColunaInicial))
       Resume Next 'Favor não remover esta linha
    End If
    '=======================================================================================
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "Importacao_XLS_Rotinas.Importar_Rel_Generico()")
    Call Importacao.InserirArquivosComErro(NomeArquivo, "Erro durante a importação")
    '=======================================================================================
    'Tratamento referente ao Rollback de Dados
    '=======================================================================================
    If Importacao.RelImportados.count > 0 Then
       If Importacao.RelImportados.Exists(NomeTabela) Then
          Call Importacao.RelImportados.Remove(NomeTabela)
       End If
    End If
    Call AuxTabela.DeletarRelatorio(NomeTabela, AuxTabela.PegarComplemento_Generico(dtRef:=dtRef))
    '=======================================================================================
    GoTo Fim
End Sub

'---------------------------------------------------------------------------------------
' Modulo....: Importacao_XLS / Módulo
' Rotina....: Importar_Rel_Generico / Sub
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 03/01/2014
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina Generica para a importação de um relatorio
'---------------------------------------------------------------------------------------
Public Sub Importar_Rel_Generico_XLS(ByRef Tabela As Object, ByVal NomeTabela As String, _
                                     ByVal arrDados As Variant, ByVal RelAux As TipoRelatorio, _
                                     ByVal NomeArquivo As String, ByVal CaminhoDoArquivo As String, _
                                     ByVal LinhaInicial As Long, ByVal ColunaInicial As Integer, _
                                     ParamArray Parametros() As Variant)
On Error GoTo TratarErro
Dim Colunas             As Object 'Scripting.Dictionary
Dim Titulo              As Object 'Scripting.Dictionary
Dim TabelaAux           As String
Dim chave               As String
Dim ChaveProduto        As String
Dim ContLinhas          As Long
Dim Ano                 As Long
Dim ContCol             As Integer
Dim PosComErro          As Integer
Dim dtRef               As Date
Dim dtRefAnterior       As Date
Dim arrAux              As Variant
Dim arrAuxProd          As Variant
Dim Incremento          As Double
Dim UsuarioInclusao     As String
Dim MomentoInclusao     As Date
Dim LinhaTitulos        As Long

    UsuarioInclusao = Publicas.ChaveUsuario()
    MomentoInclusao = VBA.Now()
    
    If Tabela Is Nothing Then GoTo Fim
    Parametros = AuxArray.Acertar_Array_Parametros(Parametros)
    
    If arrDados(LBound(arrDados, 1), LBound(arrDados, 2)) = VBA.vbNullString Then
        LinhaTitulos = LBound(arrDados, 1) + 1
    Else
        LinhaTitulos = LBound(arrDados, 1)
    End If
    Set Titulo = AuxExcel.CriarTitulos(arrDados, LinhaTitulos)
    
    Set Colunas = Conexao.PegarDicionario("Pegar_RelatorioColunas", RelAux)
    
    If Not Validar_Layout(Colunas, Titulo, "Importação de Relatórios", NomeArquivo, CaminhoDoArquivo) Then
        Call InserirArquivosComErro(NomeArquivo, "Layout do arquivo inválido.")
        GoTo Fim
    End If

    With Tabela
        Incremento = (75 / QtdArquivos) / UBound(arrDados, 1)
        For ContLinhas = LBound(arrDados, 1) + 1 To UBound(arrDados, 1) Step 1
            If (ContLinhas Mod 1000) = 0 Then Call AuxForm.IncrementaBarraProgresso(Incremento * 1000)
            .addNew
            If UBound(Parametros) >= 0 Then Call InsereValoresPadroes(Tabela, Parametros)
            For ContCol = LBound(arrDados, 2) To UBound(arrDados, 2) Step 1
                chave = RemoverQuebrasDeLinha(VBA.UCase(VBA.Trim(arrDados(LinhaTitulos, ContCol))))
                
                If Colunas.Exists(chave) Then
                    PosComErro = ContCol
                    arrAux = VBA.Split(Colunas.item(chave), "|")
                    With .Fields(arrAux(0))
                        If .Type = dao.DataTypeEnum.dbText Then
                            .value = VBA.Trim(arrDados(ContLinhas, ContCol))
                            
                        ElseIf VBA.IsDate(arrDados(ContLinhas, ContCol)) And .Type = dao.DataTypeEnum.dbDate Then
                            dtRef = VBA.CDate(arrDados(ContLinhas, ContCol))
                            .value = dtRef
                            
                        ElseIf VBA.IsNumeric(arrDados(ContLinhas, ContCol)) And _
                               (.Type = dao.DataTypeEnum.dbCurrency Or _
                                .Type = dao.DataTypeEnum.dbNumeric Or _
                                .Type = dao.DataTypeEnum.dbDouble Or _
                                .Type = dao.DataTypeEnum.dbInteger Or _
                                .Type = dao.DataTypeEnum.dbSingle Or _
                                .Type = dao.DataTypeEnum.dbLong) Then
                            .value = VBA.CDbl(arrDados(ContLinhas, ContCol))
                            
                        End If
                    End With
                End If
            Next ContCol

            'os campos abaixo TEM que existir em TODAS nossas tabelas SEMPRE.
            'Por isso não há teste pois deveremos garantir a existência deles, ao criar a tabela
            'FF-08/04/2015
            .Fields("MomentoInclusao").value = MomentoInclusao
            .Fields("UsuarioInclusao").value = UsuarioInclusao
            .Update
        Next ContLinhas
        
    End With
    If RelImportados.Exists(NomeTabela) Then
        Call RelImportados.Add(NomeTabela, Empty)
    End If
Fim:
    Call AuxTabela.FecharRecordSet(Tabela)
    Call Publicas.RemoverObjetosMemoria(Tabela)
    
On Error GoTo 0
Exit Sub
TratarErro:
    '=======================================================================================
    'Este tratamento é referente ao Tipo de Dados Incompatível dos valores
    '=======================================================================================
    If VBA.Err.Number = -2147352571 Or VBA.Err.Number = 13 Or _
       VBA.Err.Number = 2007 Or VBA.Err.Number = 3421 Then
       Call Excecoes.Tratar_Log_Erros_De_Importacao("Importação de Relatórios", NomeArquivo, CaminhoDoArquivo, VBA.Err.Number, _
       "Tipo de Dados Incorreto, Por favor, verifique o conteúdo da célula no Relatório Importado.", _
       (ContLinhas + LinhaInicial), (PosComErro + ColunaInicial))
       Resume Next 'Favor não remover esta linha
    End If
    '=======================================================================================
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "Importacao_XLS.Importar_Rel_Generico()")
    Call InserirArquivosComErro(NomeArquivo, "Erro durante a importação")
    '=======================================================================================
    'Tratamento referente ao Rollback de Dados
    '=======================================================================================
    If RelImportados.count > 0 Then
       If RelImportados.Exists(NomeTabela) Then
          Call RelImportados.Remove(NomeTabela)
       End If
    End If
    Call AuxTabela.DeletarRelatorio(NomeTabela, AuxTabela.PegarComplemento_Generico(dtRef:=dtRef))
    '=======================================================================================
    GoTo Fim
    Resume
End Sub

'********************PROPRIEDADES********************

Public Property Get QtdArquivos() As Long
    QtdArquivos = m_QtdArquivos
End Property
Public Property Let QtdArquivos(ByVal valor As Long)
    m_QtdArquivos = valor
End Property

Public Property Get DePara() As DEPARAs
    Set DePara = m_DePara
End Property
Public Property Set DePara(ByRef valor As DEPARAs)
    Set m_DePara = valor
End Property

Public Property Get RelImportados() As Object 'Scripting.Dictionary
    Set RelImportados = m_RelImportados
End Property
Public Property Set RelImportados(ByRef valor As Object)
    Set m_RelImportados = valor
End Property

Public Property Get ArquivosComErro() As Object 'Scripting.Dictionary
    Set ArquivosComErro = m_ArquivosComErro
End Property
Public Property Set ArquivosComErro(ByRef valor As Object)
    Set m_ArquivosComErro = valor
End Property

Public Property Get ArquivosImportados() As Object 'Scripting.Dictionary
    Set ArquivosImportados = m_ArquivosImportados
End Property
Public Property Set ArquivosImportados(ByRef valor As Object)
    Set m_ArquivosImportados = valor
End Property

Public Property Get ImportacaoInicio() As Date
    ImportacaoInicio = m_ImportacaoInicio
End Property
Public Property Let ImportacaoInicio(ByVal valor As Date)
    m_ImportacaoInicio = valor
End Property

Public Property Get ImportacaoFim() As Date
    ImportacaoFim = m_ImportacaoFim
End Property
Public Property Let ImportacaoFim(ByVal valor As Date)
    m_ImportacaoFim = valor
End Property

Function ValidarData(pData As String) As Boolean
    Dim RegexExp As String
    RegexExp = "(?:(?:31(\/|-|\.)(?:0?[13578]|1[02]|(?:Jan|Fev|Mar|Abr|Mai|Jun|Jul|Ago|Set|Out|Nov|Dez)))\1|(?:(?:29|30)(\/|-|\.)(?:0?[1,3-9]|1[0-2]|(?:Jan|Fev|Mar|Abr|Mai|Jun|Jul|Ago|Set|Out|Nov|Dez))\2))(?:(?:1[6-9]|[2-9]\d)?\d{2})$|^(?:29(\/|-|\.)(?:0?2|(?:Feb))\3(?:(?:(?:1[6-9]|[2-9]\d)?(?:0[48]|[2468][048]|[13579][26])|(?:(?:16|[2468][048]|[3579][26])00))))$|^(?:0?[1-9]|1\d|2[0-8])(\/|-|\.)(?:(?:0?[1-9]|(?:Jan|Fev|Mar|Abr|Mai|Jun|Jul|Ago|Set))|(?:1[0-2]|(?:Out|Nov|Dez)))\4(?:(?:1[6-9]|[2-9]\d)?\d{2})$"
    ValidarData = IsLinhaMatch(pData, RegexExp)
End Function

Function logFile_Importacao()
    logFile_Importacao = CurrentProject.Path & "\Log\LogImportacao.txt"
End Function

Function logFile_Exportacao()
    logFile_Exportacao = CurrentProject.Path & "\Log\LogExportacao.txt"
End Function
