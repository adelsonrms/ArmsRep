VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmJobs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnJobs_Click()
    If btnJobs.Caption = "Iniciar" Then
         btnJobs.Caption = "Parar"
        Call mBL_Schedule.IniciarSchedule(True)
    Else
        btnJobs.Caption = "Iniciar"
        If Not mBL_Schedule.JobScheduler Is Nothing Then
            Call mBL_Schedule.JobScheduler.StopTimer
        End If
    End If
End Sub

Sub AtualizaStatus()
    If Not mBL_Schedule.JobScheduler Is Nothing Then
        'Call mBL_Schedule.JobScheduler.SetStatusObjects(Me.btnStatusTimer, Me.btnStatusJob)
    Else
        'Me.btnStatusTimer.Caption = "Timer desligado ! Nenhum Job será executado"
    End If
End Sub

Private Sub Form_Load()
    Me.ConsiderarIntervalo.value = False
    AtualizaStatus
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    Me.ConsiderarIntervalo.value = 0
End Sub
