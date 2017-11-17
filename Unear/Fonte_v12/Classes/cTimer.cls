VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mEnabled As Boolean
Private mObj As Object
Private mStatusJobObject As Object
Private mStartTime As Date
Private mEndTime As Date
Private mcJobs As Collection
Private cLog As VBA.Collection
Private cLogTimer As VBA.Collection

Private cHandlers As Collection

Public Event OnTimer(msg As String)
Public Event OnStart()
Public Event OnFinish()
Public Event OnJobStart()

Public Enum eUnitTimer
    Segs = 1
    Mins = 2
    Hours = 3
End Enum

Private m_sLastStatusTimerMessage As String
Private m_TimerID As Variant


Sub StopTimer()
    mEnabled = False
    Call Raise_Event("OnEnd", cLog)
    Set mBL_Schedule.JobScheduler = Nothing
End Sub

Private Function getIntervalByUnit(unit As eUnitTimer) As String
    getIntervalByUnit = VBA.Choose(unit, "s", "n", "h")
End Function

Function SetDuration(Optional dtStart As Date, Optional value, Optional unit As eUnitTimer)
    If Not VBA.IsMissing(dtStart) And dtStart <> 0 Then
        mStartTime = dtStart
    Else
        mStartTime = VBA.Now
    End If
    If Not VBA.IsMissing(value) Then mEndTime = getTimerFromDuration(mStartTime, value, unit)
End Function

Public Function getTimerFromDuration(dtStart As Date, value, unit As eUnitTimer)
    Dim sInterval As String
    sInterval = getIntervalByUnit(unit)
    getTimerFromDuration = VBA.DateAdd(sInterval, value, dtStart)
End Function

Public Property Get Enabled() As Boolean
    Enabled = mEnabled
End Property

Public Property Let Enabled(value As Boolean)
    mEnabled = value
End Property

Public Property Get StartTime() As Date
    StartTime = mStartTime
End Property

Public Property Get EndTime() As Date
    EndTime = mEndTime
End Property

Public Property Get StatusTimerObject() As Object
    Set StatusTimerObject = mObj
End Property
Public Property Set StatusTimerObject(value As Object)
    Set mObj = value
End Property

Public Property Get StatusJobObject() As Object
    Set StatusJobObject = mStatusJobObject
End Property

Public Property Set StatusJobObject(value As Object)
    Set mStatusJobObject = value
End Property
' ----------------------------------------------------------------
' Procedure Name: Run
' Purpose: Inicia a execução do timer
' Procedure Kind: Sub
' Procedure Access: Public
' Author: Adelson
' Date: 04/11/2017
' ----------------------------------------------------------------
Public Sub Run()
    '----------------------------------------------------------------------------------------------------
    On Error GoTo TratarErro: Const cstr_ProcedureName As String = "TFWCliente.cTimer.Run"
    '------------------------------------------------------------------------------------------------
    Dim job As cJob
    Dim msgJob As String
    Dim msgTimer As String
    Dim CountExec As Integer
    Dim QtdExecutados As Integer
    
    
    Dim bExecutou As Boolean
    Dim bReprogramou As Boolean
    Dim dtHoraInicio As Date
    Dim dtHoraFim As Date
    Dim dtHorarioReprogramado As Date
    
    Dim vInfo() As String
    
    Set cLog = New Collection
    Call addLog("Timer inicializado ! ", "Status_Timer")
    Call addLog(mStartTime, "Inicio") 'Inicio
    Call addLog(VBA.IIf(mEndTime = 0, "Infinito", mEndTime), "Fim") 'Inicio
    Call addLog(Me.Jobs.count, "QtdJobs") 'Inicio
    Call Raise_Event("OnStart", cLog)
    
    Do While (VBA.Now <= mEndTime Or mEndTime = 0) And mEnabled
        VBA.DoEvents
        msgTimer = "Inicio :" & VBA.Format(mStartTime, "hh:nn:ss") & " | Fim :" & VBA.Format(mEndTime, "hh:nn:ss") & " | Duração : " & PegaTempoDecorrido(mEndTime - mStartTime) & " | Decorrido : " & PegaTempoDecorrido(VBA.Now - mStartTime) & " | Restante : " & PegaTempoDecorrido(mEndTime - VBA.Now) & " | Qtd Jobs : " & Me.Jobs.count
        Call setStatus(mObj, msgTimer, "Caption")
        'Verifica se há Jobs a ser executado nesse monitoramento.
        If Me.Jobs.count > 0 Then
            For Each job In Me.Jobs
                msgJob = "Job : " & job.nome & " | Proxima Execução : " & job.StartTime & " | Tempo Restante : " & PegaTempoDecorrido(job.StartTime - VBA.Now)
                Set cLog = New Collection
                Call addLog(mStartTime, "Inicio") 'Inicio
                Call addLog(mEndTime, "Termino")   'Fim
                Call addLog("Em Andamento...", "Status_Timer")
                Call addLog(VBA.Format(VBA.Now - mStartTime, "hh:nn:ss"), "Tempo_Decorrido")  'Tempo Decorrido
                Call addLog(VBA.Format(mEndTime - VBA.Now, "hh:nn:ss"), "Tempo_Restante")   'Tempo Restante
                Call addLog(job.StartTime - VBA.Now, "Tempo_Restante_Job")  'Tempo Restante
                
                Call addLog(Me.Jobs.count, "Qtd_Jobs") 'Qtd de Jobs Configurados
                Call addLog(job.nome, "Nome_Job") 'Nome do Job
                Call addLog(job.StartTime, "Data_Programada_Iniciar") 'Horario programado pra rodar
                Call addLog(job.RunAction, "Rotina") 'Rotina
                
                If VBA.Now() >= job.StartTime And job.status = NotExecuted Then
                    job.status = NotExecuted
                    job.JobID = NovoCodigo()
                    Call addLog(job.JobTimer.TimerID, "TimerID")
                    Call addLog(job.JobID, "JobID")
                    bExecutou = False
                    dtHoraInicio = VBA.Now
                    'Executa a rotina programada
                    'Passa para a função chamada algumas informações que ajudarão o controle.
                    'A instancia do Timer, do Job e do log
                    '-----------------------------------------------------------
                    Call Access.Application.Run(job.RunAction, VBA.Array(job, cLog))
                    '-----------------------------------------------------------
                    dtHoraFim = VBA.Now()
                    
                    bExecutou = True
                    job.status = Executed
                    job.CountExecuted = job.CountExecuted + 1
                    CountExec = job.CountExecuted
                    
                    
                    'Se houve instrução programada para executar, reprograma o horario
                    '-----------------------------------------------------------
                    bReprogramou = False
                    If VBA.IsArray(job.DurationAfterFinish) Then
                        job.status = NotExecuted
                        'DurationAfterFinish : Array com os dados de qtd e intervalo a ser incrementado
                        dtHorarioReprogramado = getTimerFromDuration(VBA.Now, CInt(job.DurationAfterFinish(0)), CInt(job.DurationAfterFinish(1)))
                        job.StartTime = dtHorarioReprogramado
                        bReprogramou = True
                    End If
                
                    Call addLog(dtHoraInicio, "JOB_Dt_Inicio") 'hora em que o job foi iniciado
                    Call addLog(dtHoraFim, "JOB_Dt_Fim") 'hora em que o job foi finalizado
                    Call addLog(PegaTempoDecorrido(dtHoraFim - dtHoraInicio), "JOB_Duracao") 'hora em que o job foi finalizado
                    
                    If bExecutou Then
                        Call addLog("Job executado!", "JOB_Mensagem") 'Executou
                    Else
                        Call addLog("Não executou", "JOB_Mensagem") 'Não executou
                    End If
                    
                    If bReprogramou Then
                        Call addLog("Job Reprogramado!", "JOB_Reprogramado_Mensagem") 'Rotina
                    Else
                        Call addLog("Não Reprogramou", "JOB_Reprogramado_Mensagem") 'Rotina
                    End If
                    
                    Call addLog(CountExec, "JOB_Qtd_Execucao") 'Qtd de Jobs executados ate aqui
                    Call addLog(dtHorarioReprogramado, "JOB_Proxima_Execucao")
                    
                    msgTimer = msgJob & " | Prox : " & VBA.Format(job.StartTime, "hh:nn:ss") & " | Executados : " & CountExec
                    Call setStatus(StatusJobObject, msgTimer, "Caption")
                    QtdExecutados = QtdExecutados + 1
                    Call Raise_Event("OnTimer", cLog)
                End If
            Next job
        End If
        'Stop foi solicitado para interromper
        If getFlagStop() Then Exit Do
        m_sLastStatusTimerMessage = msgTimer & " <=> " & msgJob
    Loop
    
    Call setStatus(mObj, "Inicio :" & VBA.Format(mStartTime, "hh:nn:ss") & " | Fim :" & VBA.Format(mEndTime, "hh:nn:ss") & " | Timer Finalizado !", "Caption")
    
    Set cLog = New Collection
    Call addLog(mStartTime, "Inicio") 'Inicio
    Call addLog(Me.TimerID, "TimerID")
    Call addLog(VBA.Now, "Fim") 'Inicio
    Call addLog(Me.Jobs.count, "QtdJobs") 'Inicio
    Call addLog(QtdExecutados, "QtdExecutados") 'Inicio
    Call addLog(PegaTempoDecorrido(mEndTime - mStartTime), "DuracaoProgramada")  'Inicio
    Call addLog(PegaTempoDecorrido(VBA.Now() - mStartTime), "DuracaoTotal")  'Inicio
    Call addLog("Timer Finalizado !", "Status_Timer")
    
    Call Raise_Event("OnFinish", cLog)
    
    On Error GoTo 0
Fim:
    Exit Sub
TratarErro:
    If VBA.Err <> 0 Then Call VBA.MsgBox(cstr_ProcedureName & " - " & VBA.Erl() & ">" & VBA.Err.Description, VBA.vbCritical, "Erro")
    GoTo Fim:
    Resume
End Sub

Sub addLog(msgLog, Optional pKey As String)
    If cLog Is Nothing Then Set cLog = New VBA.Collection
    If pKey <> "" Then
        cLog.Add msgLog, pKey
    Else
        cLog.Add msgLog
    End If
End Sub

Public Property Get Jobs() As VBA.Collection
    If mcJobs Is Nothing Then Set mcJobs = New VBA.Collection
    Set Jobs = mcJobs
End Property

Public Function AddJobs(ByVal vNewJob As cJob)
    If mcJobs Is Nothing Then Set mcJobs = New VBA.Collection
    Call mcJobs.Add(vNewJob)
End Function

Private Sub Raise_Event(sEvent As String, ParamArray pArgs() As Variant)
    Dim sMethodCallback  As String
    sMethodCallback = getCallback(sEvent)
    Select Case sEvent
    Case "OnStart":
        If sMethodCallback <> "" Then Call Application.Run(sMethodCallback, pArgs(0))
        RaiseEvent OnStart:
    Case "OnFinish":
        If sMethodCallback <> "" Then
            Call Application.Run(sMethodCallback, pArgs(0))
        Else
            RaiseEvent OnFinish
        End If
    Case "OnJobStart":
        RaiseEvent OnJobStart
    Case "OnTimer":
        If sMethodCallback <> "" Then Call Application.Run(sMethodCallback, pArgs(0))
    End Select
End Sub

Function getFlagStop()
    Dim bRet As Boolean
    bRet = Application.Run("ReadFlagStop")
    getFlagStop = bRet
End Function

Private Function getCallback(sEvent As String)
    On Error Resume Next
    getCallback = cHandlers.item(sEvent)
End Function

Function AddEventHandle(sEvent As String, strCallback As String)
    If cHandlers Is Nothing Then Set cHandlers = New Collection
    Call cHandlers.Add(strCallback, sEvent)
End Function

Public Property Get LastStatusTimerMessage() As String
    LastStatusTimerMessage = m_sLastStatusTimerMessage
End Property

Public Property Let LastStatusTimerMessage(ByVal sNewValue As String)
    m_sLastStatusTimerMessage = sNewValue
End Property

Sub SetStatusObjects(ParamArray vStatus() As Variant)
    If VBA.IsArray(vStatus) Then
        If UBound(vStatus, 1) > -1 Then
            Set mObj = vStatus(0)
            Set mStatusJobObject = vStatus(1)
        End If
    End If
End Sub

Private Sub setStatus(obj As CommandButton, valor As String, Optional propriedade As String = "Caption")
    On Error Resume Next
    obj.Properties(propriedade).value = valor
End Sub

Public Property Get TimerID() As Variant
    TimerID = m_TimerID
End Property

Public Property Let TimerID(ByVal NewValue As Variant)
    m_TimerID = NewValue
End Property

Public Property Get ListaLogs() As VBA.Collection
    Set ListaLogs = cLog
End Property
