Attribute VB_Name = "Excecoes"
Option Compare Database
Option Explicit

'*********************Importação********************
Private m_ErrosImp          As Object 'Scripting.Dictionary
Private m_ErroDEPARA        As Object 'Scripting.Dictionary

Public Property Get ErrosImportacao() As Object 'Scripting.Dictionary
    Set ErrosImportacao = m_ErrosImp
End Property
Public Property Set ErrosImportacao(ByRef valor As Object)
    Set m_ErrosImp = valor
End Property

Public Property Get ErroDEPARA() As Object 'Scripting.Dictionary
    Set ErroDEPARA = m_ErroDEPARA
End Property
Public Property Set ErroDEPARA(ByRef valor As Object)
    Set m_ErroDEPARA = valor
End Property

'********************SUB-ROTINAS********************
'---------------------------------------------------------------------------------------
' Modulo....: Excecoes / Módulo
' Rotina....: TratarErro / Sub
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 15/01/2014
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.:
'---------------------------------------------------------------------------------------
Public Sub TratarErro(ByVal ErroDescricao As String, ByVal ErroNum As Long, ByVal Rotina As String, _
                      Optional ByVal NumLinha As Long = 0, Optional ByVal MostrarMsg As Boolean = True, _
                      Optional ByVal MostrarStop As Boolean = False)
On Error GoTo TrataErro
    If ErroNum <> 9999 And MostrarStop Then Stop
    Debug.Print "TratarErro() - Em > " & Rotina & "(Line Number : " & NumLinha & ")", ErroNum, ErroDescricao; ""
    If MostrarMsg Then
        Call VBA.MsgBox("Um erro inesperado ocorreu ao executar o procedimento " & VBA.vbNewLine & VBA.vbNewLine & "Erro: " & VBA.Err.source & "-" & ErroDescricao & VBA.vbNewLine & _
                        "Numero: " & ErroNum & VBA.vbNewLine & _
                        "Rotina: " & Rotina & VBA.vbNewLine & _
                        "Linha: " & NumLinha, VBA.vbCritical, "TFW VBA.Error Handle - Erro Inesperado")
    End If
    Call GravaErro(ErroDescricao, ErroNum, Rotina, NumLinha)
'    If VBA.ErroNum <> 0 Then Stop
    Call LogaErro(ErroDescricao, ErroNum, Rotina, NumLinha)
Exit Sub
TrataErro:
    'FAÇA NADA
End Sub

Public Sub Tratar_Log_DE_PARA(ByVal CaminhoDoArquivo As String, ByVal NomeArquivo As String, ByVal RelatorioExecucao As String, _
                              ByVal TabelaDE_PARA As String, ByVal CodigoNaoEncontrado As String)
On Error GoTo TrataErro
    Call Conexao.InserirRegistros("Insere_Log_DE_PARA", CaminhoDoArquivo, NomeArquivo, RelatorioExecucao, TabelaDE_PARA, CodigoNaoEncontrado, _
                                 VBA.Now, VBA.Environ("ComputerName"), VBA.Environ("Username"))
Exit Sub
TrataErro:
    'FAÇA NADA
End Sub

Public Sub Tratar_Log_Erros_De_Importacao(ByVal TipoRelatorio As String, ByVal NomeDoArquivo As String, _
                                          ByVal CaminhoDoArquivo, ByVal NumErro As Long, ByVal descricao As String, _
                                          ByVal linha As Long, ByVal coluna As Long)
On Error GoTo TratarErro

    Call Conexao.InserirRegistros("Insere_LogErro_Importacao", TipoRelatorio, NomeDoArquivo, CaminhoDoArquivo, _
                                 NumErro, descricao, linha, coluna, VBA.Now, VBA.Environ("ComputerName"), _
                                 VBA.Environ("Username"))
Exit Sub
TratarErro:
    'FAÇA NADA
End Sub

Private Sub LogaErro(ByVal ErroDescricao As String, ByVal ErroNum As Long, _
                     ByVal Rotina As String, ByVal NumLinha As Long)
On Error GoTo TrataErro
    Call Conexao.InserirRegistros("Insere_LogErro", ErroNum, ErroDescricao, Rotina, NumLinha, _
                                 Access.Application.Name, Access.Application.Version, Access.Application.Build, _
                                 VBA.Now, VBA.Environ("ComputerName"), VBA.Environ("Username"))
Exit Sub
TrataErro:
    'FAÇA NADA
End Sub

Private Sub LogaImportacao(ByVal Relatorio As String, ByVal tpRelatorio As TipoRelatorio, ByVal Rotina As String, _
                          ByVal descricao As String, ByVal NomeDoArquivo As String, ByVal NumLinha As Long, _
                          ByVal dia As Long, ByVal Mes As Long, ByVal Ano As Long)
On Error GoTo TrataErro
    Call Conexao.InserirRegistros("Insere_LogImportacao", Relatorio, tpRelatorio, Rotina, descricao, NomeDoArquivo, NumLinha, _
                                  dia, Mes, Ano, VBA.Now, VBA.Environ("ComputerName"), VBA.Environ("Username"))
Exit Sub
TrataErro:
    'FAÇA NADA
End Sub


Private Sub GravaErro(ByVal ErroDescricao As String, ByVal ErroNum As Long, _
                      ByVal Rotina As String, ByVal NumLinha As Long)
On Error GoTo TrataErro
Dim fso             As Object
Dim strNomePasta    As String
Dim strNomeArq      As String

    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    strNomePasta = CurrentProject.Path & "\Log"
    strNomeArq = strNomePasta & "\LogErro.txt"
    Call VerificaECriaPasta(fso, strNomePasta)
    Call VerificaECriaArquivo(fso, strNomeArq)
    
    Open strNomeArq For Append As #1
    Print #1, "Erro: " & ErroDescricao & VBA.vbCrLf & _
              "Numero: " & ErroNum & VBA.vbCrLf & _
              "Rotina: " & Rotina & VBA.vbCrLf & _
              "Linha: " & NumLinha & VBA.vbCrLf & _
              "Data: " & VBA.Now() & VBA.vbCrLf & _
              "Usuario: " & VBA.Environ("UserName") & VBA.vbCrLf & _
              "Máquina: " & VBA.Environ("ComputerName") & VBA.vbCrLf & _
              "-------------------------------------------------------------------"
    Close #1
    Call Publicas.RemoverObjetosMemoria(fso)
Exit Sub
TrataErro:
    'FAÇA NADA
End Sub

Private Sub VerificaECriaPasta(ByRef fso As Object, ByVal strNomePasta As String)
On Error GoTo TrataErro
    If Not fso.FolderExists(strNomePasta) Then
        Call fso.CreateFolder(strNomePasta)
    End If
Exit Sub
TrataErro:
    'FAÇA NADA
End Sub

Private Sub VerificaECriaArquivo(ByRef fso As Object, ByVal strNomeArq As String)
On Error GoTo TrataErro
    If Not fso.FileExists(strNomeArq) Then
        fso.createTextFile (strNomeArq)
    End If
Exit Sub
TrataErro:
    'FAÇA NADA
End Sub

