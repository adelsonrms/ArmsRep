Attribute VB_Name = "mBL_Schedule"
Public JobScheduler As cTimer

Sub StartScheduler(Optional bNew As Boolean)
    If bNew Then
        Set JobScheduler = New cTimer
    Else
        If JobScheduler Is Nothing Then Set JobScheduler = New cTimer
    End If
End Sub

Function DeterminarIntervalo(valor As Integer, unidade As eUnitTimer)
    Call StartScheduler
    If Nz(valor) <> 0 And Nz(unidade) <> 0 Then
        Call JobScheduler.SetDuration(VBA.Now, valor, unidade)
    End If
End Function

Sub IniciarSchedule(bHabilita As Boolean, ParamArray vStatus() As Variant)
    Dim job As cJob
    Dim dtStart As Date
    Dim rngJob As Object
    
    Call LimpaMarcacao
    'AdicionaBarraMonitor
    Call StartScheduler
    
    JobScheduler.TimerID = NovoCodigo()
    JobScheduler.Enabled = bHabilita
    
    'Registra o evento OnTimer para executar a função 'RegistraLogJob()' externa para capturar os logs
    Call JobScheduler.AddEventHandle("OnTimer", "RegistraLogJob")
    Call JobScheduler.AddEventHandle("OnStart", "RegistraLogTimer")
    Call JobScheduler.AddEventHandle("OnFinish", "RegistraLogTimerFinal")
    
    'Definie a duração de quanto tempo o timer irá rodar
    'A função SetDuration define a data/hora de inicio e define a data/hora a partir da data de inicio
    dtStart = JobScheduler.StartTime
    'Configura objetos para captura mensagens de log
    '------------------------------------------------------
    'Call JobScheduler.SetStatusObjects(vStatus(0), vStatus(1))
    
    Set rsJobs = CurrentDb.OpenRecordset("tblJobs")
    With rsJobs
        Do While Not .EOF
            If .Fields("ativado").value = -1 Then
                'Cria e programa um Job
                '------------------------------------------------------
                Set job = New cJob
                With job
                    Set .JobTimer = JobScheduler
                    .nome = rsJobs.Fields("Nome").value
                    .StartTime = JobScheduler.getTimerFromDuration(dtStart, rsJobs.Fields("Frequencia").value, rsJobs.Fields("Recorrencia").value)
                    .DurationAfterFinish = VBA.Array(rsJobs.Fields("Frequencia").value, rsJobs.Fields("Recorrencia").value) 'Será executado a cada 1 Min
                    .RunAction = rsJobs.Fields("Acao").value
                End With
                Call JobScheduler.AddJobs(job)
            End If
            .MoveNext
        Loop
    End With
    'Inicia o timer
    'Call JobScheduler.Run
    'Aguarda a finalização
End Sub

Sub AdicionaBarraMonitor()
    Dim cb As Object
    Dim ctr As Object ' CommandBarButton
    
    On Error Resume Next
    Call Application.VBE.CommandBars("barTimer").Delete
    Set cb = Application.VBE.CommandBars.Add("barTimer")
    cb.Visible = True
    cb.position = msoBarBottom
    
    Set ctr = cb.Controls.Add
    ctr.Style = msoButtonCaption

    Set ctr = cb.Controls.Add
    ctr.Style = msoButtonCaption
End Sub

Sub RegistraLogTimer(cLog As Collection)
    Dim sLog As String
    sLog = VBA.Left(cLog.item("Status_Timer") & VBA.Space(30), 30)
    sLog = sLog & VBA.Left("| Inicio : " & cLog.item("Inicio") & VBA.Space(35), 35)
    sLog = sLog & VBA.Left("| Termino : " & cLog.item("Fim") & VBA.Space(35), 35)
    sLog = sLog & "| Qtd de Jobs : " & cLog.item("QtdJobs")
    Call RegistraLog(sLog, "Jobs.txt")
End Sub

Sub RegistraLogTimerFinal(cLog As Collection)
    Dim sLog As String
    Dim strAnexo As String
    Dim sTimer As String
    
    sLog = VBA.Left(cLog.item("Status_Timer") & VBA.Space(30), 30)
    sLog = sLog & VBA.Left("| Timer ID : " & cLog.item("TimerID") & VBA.Space(35), 35)
    sLog = sLog & VBA.Left("| Inicio : " & cLog.item("Inicio") & VBA.Space(35), 35)
    sLog = sLog & VBA.Left("| Termino : " & cLog.item("Fim") & VBA.Space(35), 35)
    sLog = sLog & "| Qtd de Jobs : " & cLog.item("QtdJobs")
    Call RegistraLog(sLog, "Jobs.txt")
    
    Call EnviarLogTimer(cLog.item("TimerID"), cLog)
    
    Set mBL_Schedule.JobScheduler = Nothing
    
End Sub

Sub EnviarLogTimer(sTimer As String, Optional cVariaveis As Collection)
    Dim strAnexo As String
    Dim Resumo As Object
    
    Call AtualizaVariavelModelo(17, "DataDeEnvio", VBA.Date)
    Call AtualizaVariavelModelo(17, "TimerID", sTimer)
    
    If Not cVariaveis Is Nothing Then
        Call AtualizaVariavelModelo(17, "DataInicio", cVariaveis.item("Inicio"))
        Call AtualizaVariavelModelo(17, "DataFim", cVariaveis.item("Fim"))
        Call AtualizaVariavelModelo(17, "Duracao", cVariaveis.item("DuracaoTotal"))
        Call AtualizaVariavelModelo(17, "QtdJobs", cVariaveis.item("QtdExecutados"))
    End If
    
    Call Inicializar_Globais
    
    
    
    Set Resumo = Conexao.PegarRS("Pegar_ResultadoCampanhasPorTimer", sTimer)
    If Not Resumo Is Nothing Then
        If Not Resumo.EOF Then
            With Resumo
            
                Call AtualizaVariavelModelo(17, "QtdCampanhas", Access.Nz(!QtdCampanhas.value, 0))
                Call AtualizaVariavelModelo(17, "QtdSucesso", Access.Nz(!SUCESSO.value, 0))
                Call AtualizaVariavelModelo(17, "QtdErros", Access.Nz(!Erros.value, 0))
                Call AtualizaVariavelModelo(17, "QtdTestada", Access.Nz(!Testado.value, 0))
                Call AtualizaVariavelModelo(17, "QtdExecutada", Access.Nz(!Executado.value, 0))
                
                If Access.Nz(!Executado.value, 0) = Access.Nz(!QtdCampanhas.value, 0) Then
                    Call AtualizaVariavelModelo(17, "Mensagem", "Todas as campanhas criadas foram processadas com sucesso. Confira no site da plataforma as informações sobre as campanhas enviadas")
                ElseIf (Access.Nz(!Executado.value, 0) = Access.Nz(!QtdCampanhas.value, 0)) And Access.Nz(!Testado.value, 0) <> Access.Nz(!Executado.value, 0) Then
                    Call AtualizaVariavelModelo(17, "Mensagem", "Todas as campanhas criadas foram processadas porém algumas não foram registradas os testes. Confira detalhes no site")
                Else
                    Call AtualizaVariavelModelo(17, "Mensagem", "Algumas campanhas não foram criadas. Confira detalhes no site")
                End If
                
            End With
        End If
    End If
    
    If HabilitaEnvioEmail Then
        Call Inicializar_Globais
        strAnexo = ""
        strEventoID = ""
        Set rsAnexos = Conexao.PegarRS("Pegar_EmailsPendentes", sTimer)
        If Not rsAnexos.EOF Then
            Do While Not rsAnexos.EOF
                If FileExists(Nz(rsAnexos!ArquivoLog.value)) Then
                    strAnexo = strAnexo & ";" & Nz(rsAnexos!ArquivoLog.value)
                    strEventoID = strEventoID & ",'" & Nz(rsAnexos!EventoID.value) & "'"
                End If
                rsAnexos.MoveNext
            Loop
        End If
        
        If strEventoID <> "" Then
            strAnexo = VBA.Mid(strAnexo, 2)
            strEventoID = VBA.Mid(strEventoID, 2)
            Call EnviarEmail(17, strAnexo)
            Call CurrentDb.Execute("UPDATE tblLog SET EnviarEmail=0 WHERE EventoID IN(" & strEventoID & ")")
        End If
    End If
    
End Sub

Sub RegistraLogJob(cLog As Collection)
    Dim sLog As String
    sLog = VBA.Left("Job : " & cLog.item("Nome_Job") & VBA.Space(20), 20)
    sLog = sLog & VBA.Left("| Acao : " & cLog.item("Rotina") & VBA.Space(30), 30)
    sLog = sLog & VBA.Left("| Inicio : " & cLog.item("JOB_Dt_Inicio") & VBA.Space(30), 30)
    sLog = sLog & VBA.Left("| Fim : " & cLog.item("JOB_Dt_Fim") & VBA.Space(30), 30)
    sLog = sLog & VBA.Left("| Duracao : " & cLog.item("JOB_Duracao") & VBA.Space(25), 25)
    sLog = sLog & VBA.Left("| Proximo : " & cLog.item("JOB_Proxima_Execucao") & VBA.Space(35), 35)
    sLog = sLog & VBA.Left("| Qtd Execucoes : " & cLog.item("JOB_Qtd_Execucao") & VBA.Space(20), 20)
    sLog = sLog & "| Status : " & cLog.item("JOB_Mensagem")
    Call RegistraLog(sLog, "Jobs.txt")
End Sub

Function ReadFlagStop()
    ReadFlagStop = bHabilita
End Function


Function HabilitaEnvioEmail() As Boolean
    With CurrentDb.OpenRecordset("tblConfigTimer")
        HabilitaEnvioEmail = Nz(.Fields("HabilitarEmail").value, 0)
    End With
End Function

Function PegarHoraInicio()
    With CurrentDb.OpenRecordset("tblConfigTimer")
        PegarHoraInicio = VBA.Date + VBA.CDate(Nz(.Fields("HoraInicio").value, 0))
    End With
End Function

Function PegarHoraFim()
    With CurrentDb.OpenRecordset("tblConfigTimer")
        PegarHoraFim = VBA.Date + VBA.CDate(Nz(.Fields("HoraFim").value, 0))
    End With
End Function

Function TimerAtivacao() As Boolean
    With CurrentDb.OpenRecordset("tblConfigTimer")
        TimerAtivacao = Nz(.Fields("Ativado").value, 0)
    End With
End Function

Function ConsiderarIntervalo() As Boolean
    With CurrentDb.OpenRecordset("tblConfigTimer")
        ConsiderarIntervalo = Nz(.Fields("ConsiderarIntervalo").value, 0)
    End With
End Function

Sub SalvaFlagTimer(info As String, valor)
    With CurrentDb.OpenRecordset("tblConfigTimer")
        If .EOF Then .addNew Else .edit
        .Fields(info).value = valor
        .Update
    End With
End Sub
