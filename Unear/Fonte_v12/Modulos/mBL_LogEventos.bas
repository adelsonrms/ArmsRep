Attribute VB_Name = "mBL_LogEventos"
Option Compare Database

' ----------------------------------------------------------------
' Procedure Name: RegistrarLogEvento
' Purpose: Registra na tabela de log as ocorrencias de execução das campanhas
' Procedure Kind: Function
' Procedure Access: Public
' Parameter EventoID (String):
' Parameter status (String):
' Parameter ArquivoLog (String):
' Author: Adelson
' Date: 05/11/2017
' ----------------------------------------------------------------
Function RegistrarLogEvento(EventoID As String, status As String, ArquivoLogDetalhes As String, ArquivoLog As String, bEnviarEmail As Boolean, TimerID As String, JobID As String, DisparoID As String, Optional bStatus As Boolean = False)
    '----------------------------------------------------------------------------------------------------
    On Error GoTo TratarErro: Const cstr_ProcedureName As String = "TFWCliente.mBL_LogEventos.RegistrarLogEvento"
    '------------------------------------------------------------------------------------------------
    Dim rsLog As Object
    Dim bInicio As Boolean
    Dim strMensagem As String
    Set rsLog = CurrentDb.OpenRecordset("SELECT * FROM tblLog WHERE EventoID = '" & EventoID & "'")
    bInicio = rsLog.EOF
    With rsLog
        If bInicio Then .addNew Else .edit
        !EventoID.value = EventoID
        !TimerID.value = TimerID
        !JobID.value = JobID
        !DisparoID.value = DisparoID
        If bInicio Then
            !INICIO.value = VBA.Now
        Else
            !Fim.value = VBA.Now
            !Duracao.value = PegaTempoDecorrido(!Fim.value - !INICIO.value)
        End If
        !status.value = status
        !ArquivoLog.value = ArquivoLogDetalhes
        For i = 0 To .Fields.count - 1
            strMensagem = strMensagem & " | " & .Fields(i).value
        Next i
        !EnviarEmail.value = bEnviarEmail
        !StatusProcessamento.value = bStatus
        .Update
    End With
    If ArquivoLog <> "" Then Call auxAuditoria.RegistraLog(strMensagem, ArquivoLog)
    On Error GoTo 0
Fim:
    Exit Function
TratarErro:
    If VBA.Err <> 0 Then Call VBA.MsgBox(cstr_ProcedureName & " - " & VBA.Erl() & ">" & VBA.Err.Description, VBA.vbCritical, "Erro")
    GoTo Fim:
    Resume
End Function


Function MarcaJob(strNome As String)
    '----------------------------------------------------------------------------------------------------
    On Error GoTo TratarErro: Const cstr_ProcedureName As String = "TFWCliente.mBL_LogEventos.MarcaJob"
    '------------------------------------------------------------------------------------------------
    Dim rsLog As Object
    Set rsLog = CurrentDb.OpenRecordset("SELECT * FROM tblExecucao WHERE NomeJob = '" & strNome & "'")
    With rsLog
        If .EOF Then .addNew Else .edit
        !NomeJob.value = strNome
        !Execucoes.value = Nz(!Execucoes.value, 0) + 1
        .Update
    End With
    On Error GoTo 0
Fim:
    Exit Function
TratarErro:
    If VBA.Err <> 0 Then Call VBA.MsgBox(cstr_ProcedureName & " - " & VBA.Erl() & ">" & VBA.Err.Description, VBA.vbCritical, "Erro")
    GoTo Fim:
    Resume
End Function

Sub LimpaMarcacao()
    CurrentDb.Execute ("DELETE FROM tblExecucao")
End Sub

Function PegarQtdExecucoes(strNome As String)
    With CurrentDb.OpenRecordset("SELECT sum(Execucoes) FROM tblExecucao WHERE NomeJob = '" & strNome & "'")
        PegarQtdExecucoes = Nz(.Fields(0).value, 0)
    End With
End Function

Sub RegistraStatus(Mensagem As String, LogTxt As String, Progresso As Boolean, Optional IncluirDataHoraLog As Boolean = True)
    If Progresso Then Call showMessage(Mensagem, "Status", SemProgresso, -1)
    If LogTxt <> "" Then Call RegistraLog(pMessage:=Mensagem, pFileOutPut:=LogTxt, bPrintDatetime:=IncluirDataHoraLog, bIncluirUsuario:=False, bOnlyTime:=True)
End Sub
