Attribute VB_Name = "AuxMensagens"
Option Compare Database
Option Explicit

'----------------------------------------------------------------------------
'Template name: Messages Manager VBA
'Developed by Fernando R. Couto Fernandes / Jefferson Dantas
'Creation date: 01/Jun/2008
'Last Modified Date:  30/Jun/2014
'Contatos: fernando.fernandes@outlook.com.br / jefferdantas@gmail.com
'----------------------------------------------------------------------------
'O objetivo deste modelo é
'1) Tornar fácil a catalogação, visualização e documentação de todas as mensagens da aplicação
'2) Facilita também verificar as situações e passos para reprodução de cada mensagem (se o desenvolvedor assim desejar)
'3) Utilizar uma mesma rotina para exibir todas as perguntas ou mensagens, assim, fica fácil atualizar padrões de título por exemplo, mexendo num local só.

'Option Private Module
Option Base 1

'Definição do tipo Mensagem que conterá os detalhes de qualquer mensagem exibida
Private Type Mensagem
    ID      As String
    Titulo  As String
    Estilo  As VbMsgBoxStyle
    Texto   As String
End Type

'---------------------------------------------------------------------------------------
' Modulo....: AuxForm / Módulo
' Rotina....: MessageBoxMaster() / Function
' Autor.....: Fernando Fernandes
' Contato...: fernando@tecnun.com.br
' Data......: 20/10/2014
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina que busca os dados da mensagem na base, para exibir no formulário messagebox personalizado do sistema
'---------------------------------------------------------------------------------------
Public Function MessageBoxMaster(ByVal IDMensagem As String, _
                                 ParamArray Variaveis() As Variant) As VbMsgBoxResult
On Error GoTo TrataErro
Dim tMensagem   As Mensagem
Dim Mensagem    As String
Dim tipoMsgBox  As VbMsgBoxStyle
Dim Titulo      As String
Dim strAux      As String

    Call Publicas.Inicializar_Globais
    tMensagem = fnGetMessage(IDMensagem)
    Call ReplaceVariables(tMensagem, Variaveis)
    With tMensagem
        If .Titulo = VBA.vbNullString Then
            .Titulo = VariaveisEConstantes.appName
        End If
        strAux = Publicas.GerarChaves(.Titulo, .Texto, .Estilo)
    End With
    Call Access.DoCmd.openForm(FormName:="frmMsgBoxMaster", OpenArgs:=strAux, WindowMode:=Access.acDialog)
    MessageBoxMaster = VariaveisEConstantes.Resultado
    Call Access.DoCmd.Close(acForm, "frmMsgBoxMaster")
Exit Function
TrataErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxForm.MessageBoxMaster2()", Erl)
    Exit Function
Resume
End Function

'Esta rotina é responsável por "instanciar" o "objeto" Mensagem já com os dados correspondentes da Mensagem em questão,
'definidos pelo ID utilizado na chamada
Private Function fnGetMessage(idMessage As String) As Mensagem
On Error GoTo TreatError
Dim arrMensagem As Variant
Dim arrMessages As Variant
Dim lLinha      As Long
Dim sStyle      As String
Dim arrStyle    As Variant
Dim lCont       As Long
Dim tpMensagem  As Mensagem
    

    'ARMS, 09/10/2015 - Conecta-se ao banco de dados Client para consultar a tabela tblMessages
    '-------------------------------------------------------------------------------
    Dim cnDB As ConexaoDB
    Set cnDB = Conexao
    '-------------------------------------------------------------------------------
    
    arrMensagem = cnDB.PegarArray(cnDB.PegarRS("Pegar_Mensagem", idMessage))
    
    If VBA.IsArray(arrMensagem) Then
        With tpMensagem

            'recuperando o título da mensagem de acordo com o que está na planilha
            .ID = arrMensagem(0, 0)

            'recuperando o título da mensagem de acordo com o que está na planilha
            .Titulo = arrMensagem(1, 0)

            'recuperando o conteúdo da mensagem de acordo com o que está na planilha
            .Texto = arrMensagem(2, 0)

            'recuperando o estilo da mensagem de acordo com o que está na planilha
            sStyle = arrMensagem(3, 0)

            'a rotina abaixo trata os estilos, convertendo os textos da planilha em números e consequentemente em botões válidos em VBA.
            'aplicando split por sinais de "+" para identificar item a item
            arrStyle = VBA.Split(sStyle, "+")
            'se o resultado for uma array, vamos tratar cada elemento da array
            If VBA.IsArray(arrStyle) Then
                For lCont = 0 To UBound(arrStyle)
                    .Estilo = .Estilo + fnDetermineStyle(arrStyle(lCont))
                Next lCont
            Else
                'caso contrário, vamos tratar somente o elemento único
                .Estilo = fnDetermineStyle(sStyle)
            End If
        End With
    End If
    fnGetMessage = tpMensagem
Exit Function
TreatError:
' a rotina de tratamento de erro abaixo NÃO PODE NUNCA utilizar rotinas deste módulo, para não causar estouro de pilha de chamada.
'    VBA.MsgBox "Error in the routine that searches for messages." & VBA.vbCrLf & "Routine: fnGetMessage()",VBA.vbCritical +VBA.vbOKOnly, "Unexpected error"
    VBA.MsgBox "Houve um erro na rotina de buscar mensagens." & VBA.vbCrLf & "Rotina: fnGetMessage()", VBA.vbCritical + VBA.vbOKOnly, "Erro imprevisto"
End Function

'Esta rotina vai converter o texto "VbOkOnly", por exemplo, no VBA.MsgBoxStyle do VBA, VbOkOnly,
'permitindo assima exibição dos estilos e botões de forma correta.
Private Function fnDetermineStyle(ByVal sButton As String) As VbMsgBoxStyle
On Error Resume Next
    sButton = VBA.UCase(VBA.Trim(sButton))
    Select Case sButton
        Case VBA.UCase("vbOKOnly")
            fnDetermineStyle = VBA.vbOKOnly
        Case VBA.UCase("vbInformation")
            fnDetermineStyle = VBA.vbInformation
        Case VBA.UCase("vbExclamation")
            fnDetermineStyle = VBA.vbExclamation
        Case VBA.UCase("vbCritical")
            fnDetermineStyle = VBA.vbCritical
        Case VBA.UCase("vbQuestion")
            fnDetermineStyle = VBA.vbQuestion
        Case VBA.UCase("vbYesNo")
            fnDetermineStyle = VBA.vbYesNo
        Case VBA.UCase("vbYesNoCancel")
            fnDetermineStyle = VBA.vbYesNoCancel
        Case VBA.UCase("vbOKCancel")
            fnDetermineStyle = VBA.vbOKCancel
        Case VBA.UCase("vbRetryCancel")
            fnDetermineStyle = VBA.vbRetryCancel
        Case VBA.UCase("vbSystemModal")
            fnDetermineStyle = VBA.vbSystemModal
        Case VBA.UCase("vbAbortRetryIgnore")
            fnDetermineStyle = VBA.vbAbortRetryIgnore
        Case VBA.UCase("vbApplicationModal")
            fnDetermineStyle = VBA.vbApplicationModal
        Case VBA.UCase("vbDefaultButton1")
            fnDetermineStyle = VBA.vbDefaultButton1
        Case VBA.UCase("vbDefaultButton2")
            fnDetermineStyle = VBA.vbDefaultButton2
        Case VBA.UCase("vbDefaultButton3")
            fnDetermineStyle = VBA.vbDefaultButton3
        Case VBA.UCase("vbDefaultButton4")
            fnDetermineStyle = VBA.vbDefaultButton4
        Case VBA.UCase("VbMsgBoxHelpButton")
            fnDetermineStyle = VBA.vbMsgBoxHelpButton
        Case VBA.UCase("VbMsgBoxRight")
            fnDetermineStyle = VBA.vbMsgBoxRight
        Case VBA.UCase("VbMsgBoxRtlReading")
            fnDetermineStyle = VBA.vbMsgBoxRtlReading
        Case VBA.UCase("VbMsgBoxSetForeground")
            fnDetermineStyle = VBA.vbMsgBoxSetForeground
        Case Else
            fnDetermineStyle = 0
    End Select
End Function

Private Sub ReplaceVariables(ByRef oMessage As Mensagem, _
                             ByVal Variaveis As Variant)
On Error Resume Next
Dim contador        As Integer
    With oMessage
        For contador = 0 To UBound(Variaveis, 1) Step 1
            .Texto = VBA.Replace(.Texto, "#Variavel" & VBA.Format(contador + 1, "00"), Variaveis(contador))
        Next contador
    End With
End Sub

