VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfStatusTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub btnDataBase_Click()
    Call AbrirSubFormulario("frmCaminhoBE", "Configuração Back-End")
End Sub

Private Sub AtivarTimer_AfterUpdate()
    If AtivarTimer.value = -1 Then
        Me.TimerInterval = 1000
        AtualizaMonitor
    Else
        Me.TimerInterval = 0
        lblStatus.Caption = "Desativado !"
        AtualizaMonitor
        lblStatusTimer.Caption = lblStatus.Caption
    End If
End Sub

Private Sub btnEditarAgendamentos_Click()
    Call AbrirSubFormulario("frmJobs", "Configurar Agendamentos")
End Sub

Private Sub btnJobs_Click()
    If btnJobs.Caption = "Iniciar" Then
        Call salvarIntervalo
        Me.TimerInterval = 1000
        Call SalvaFlagTimer("ativado", -1)
        Call IniciarMonitor
    Else
        Call SalvaFlagTimer("ativado", 0)
        Call IniciarMonitor
        Me.TimerInterval = 0
        AtualizaMonitor
    End If
End Sub

Private Sub cmbRecorrencia_AfterUpdate()
    salvarIntervalo
End Sub

Private Sub cmbValor_AfterUpdate()
    salvarIntervalo
End Sub

Sub salvarIntervalo()
    Call mBL_Schedule.DeterminarIntervalo(valor:=Access.Nz(Me.cmbValor.value), unidade:=Access.Nz(Me.cmbRecorrencia.value))
    RefreshIntervalo
    Me.Requery
End Sub

Private Sub Form_Load()
    Call Me.Reload
    Call salvarIntervalo
End Sub

Public Sub Reload()
    'Inicia ja executando
    Call SalvaFlagTimer("ativado", -1)
    Call SalvaFlagTimer("ConsiderarIntervalo", 0)
End Sub

Function RefreshIntervalo()
    Me.HoraInicio.value = JobScheduler.StartTime
    Me.HoraFim.value = JobScheduler.EndTime
End Function

Private Sub Form_Timer()
    If (VBA.Now() >= VBA.CDate(PegarHoraInicio()) And VBA.Now() <= VBA.CDate(PegarHoraFim())) And TimerAtivacao Then
        If TimerAtivacao = False Then
            Call SalvaFlagTimer("ativado", -1)
            'Se chegar no horario de fim e ainda esta ativado
            Call IniciarMonitor
        Else
            AtualizaMonitor
        End If
    Else
        AtualizaMonitor
    End If
End Sub

Sub AtualizaMonitor()
    Dim flag As Variant
'    flag = flagIniciar()
    flag = (Not mBL_Schedule.JobScheduler Is Nothing)
    'Me.imgErro.Visible = Not Me.imgAviso.Visible
    If flag Then
        btnJobs.Caption = "Parar"
        lblStatusTimer.Caption = mBL_Schedule.JobScheduler.LastStatusTimerMessage
        lblStatusTimer.BackColor = 5026082
        lblStatusTimer.FontBold = True
        lblStatusTimer.ForeColor = 16777215
        Me.cmbRecorrencia.Enabled = False
        Me.cmbValor.Enabled = False
        Me.btnEditarAgendamentos.Enabled = False
        
        On Error Resume Next
        If JobScheduler.ListaLogs("Tempo_Restante_Job") < VBA.TimeSerial(0, 0, 5) Then
            If VBA.Err = 0 Then
            lblStatusTimer.BackColor = vbRed
            lblStatusTimer.Caption = "Job em execução, aguarde...."
            End If
        End If
        
    Else
        btnJobs.Caption = "Iniciar"
        lblStatusTimer.BackColor = 16777215
        lblStatusTimer.FontBold = False
        lblStatusTimer.ForeColor = 0
        lblStatusTimer.Caption = "Agendamento desativado. Clique em 'Iniciar'"
        Me.cmbRecorrencia.Enabled = True
        Me.cmbValor.Enabled = True
        Me.btnEditarAgendamentos.Enabled = True
    End If
End Sub

Private Function flagIniciar()
    If (VBA.Now() >= VBA.CDate(PegarHoraInicio()) And VBA.Now() <= VBA.CDate(PegarHoraFim())) Then
        'Dentro de range
        If TimerAtivacao = False Then
            Call SalvaFlagTimer("ativado", -1)
            'Se chegar no horario de fim e ainda esta ativado
            Call IniciarMonitor
            flagIniciar = VBA.Array(-1, "Timer executando dentro do intervalo programado (Inicio/Fim)")
        End If
    Else
        Call SalvaFlagTimer("ativado", 0)
        Call IniciarMonitor
        flagIniciar = VBA.Array(0, "Timer Interrompido Fora do intervalo (Inicio/Fim)")
    End If
End Function

Sub IniciarMonitor()
    If TimerAtivacao Then
        btnJobs.Caption = "Parar"
        Call mBL_Schedule.IniciarSchedule(True)
        Call mBL_Schedule.JobScheduler.Run
    Else
        btnJobs.Caption = "Iniciar"
        If Not mBL_Schedule.JobScheduler Is Nothing Then Call mBL_Schedule.JobScheduler.StopTimer
    End If
End Sub



