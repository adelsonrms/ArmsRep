VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim vtm As Date
Private oLv As ListBox
Private m_itipoVisualizacao As Integer
Private m_dtstartTime As Date
Private m_idefualtPagina As Integer
Private m_lMaxProgress As Long
Private m_bUpdateOnlyMensagens As Boolean
Private datafutura As Date
Private Sub btnTestaProgresso_Click()
    Call IncremensaProgressoBar(txtQtdIteracoes.value)
    txtContador.value = Me.btnProgress_Frente.tag
End Sub

Sub IncremensaProgressoBar(Optional vlMaximo As Long = 0, Optional modoProgress As eModoProgress = eModoProgress.EmEndamento)
    Dim txProporcao As Double, dblAndamento As Integer
    Dim dblPercent As Double
    
    Dim valor As Long
    
    If modoProgress = INICIO Then
        Me.btnProgress_Contorno.Visible = False
        Me.btnProgress_Frente.Visible = False
        Me.lblTimeDecorrido.Visible = False
        
        Me.btnProgress_Contorno.tag = vlMaximo
        Me.btnProgress_Frente.tag = 0
        Me.lblTimeDecorrido.Caption = ""
        Me.lblTime.Caption = VBA.Now()
    Else
        If VBA.IsNumeric(Me.btnProgress_Contorno.tag) Then vlMaximo = VBA.CLng(Me.btnProgress_Contorno.tag)
    End If
    
    If modoProgress = Fim Then 'Reset
        Me.btnProgress_Frente.tag = 0
        Me.btnProgress_Frente.Width = 0
        lblProgress_Percent.Caption = ""
        Me.btnProgress_DeTotal.Caption = ""
        lblProgress_Percent.Visible = False
        Me.btnProgress_Frente.Visible = False
        Me.btnProgress_DeTotal.Visible = False
        btnProgress_Contorno.tag = 0
        Me.btnProgress_Contorno.Visible = False
        Me.cmdFechar.Visible = True
        
        If VBA.IsDate(Me.lblTime.Caption) Then
            Me.lblTimeDecorrido.Caption = "Duração total deste processamento  : " & PegaTempoDecorrido(VBA.Now() - VBA.CDate(Me.lblTime.Caption))
            Me.lblTimeDecorrido.Visible = True
        End If
    Else
        
        If Not VBA.IsNumeric(Me.btnProgress_Frente.tag) Then Me.btnProgress_Frente.tag = 0
        'Or modoProgress = inicio
        If modoProgress = EmEndamento Or modoProgress = EmEndamento_MensagemFixa Then
            Me.btnProgress_Frente.tag = VBA.IIf(Me.btnProgress_Frente.tag = "", 0, VBA.CInt(Me.btnProgress_Frente.tag)) + 1
        Else
            Me.btnProgress_Frente.tag = 0
        End If
        
        If VBA.IsNumeric(Me.btnProgress_Frente.tag) Then
            valor = VBA.CInt(Me.btnProgress_Frente.tag)
        End If
        
        If vlMaximo = 0 Then vlMaximo = btnProgress_Contorno.Width
        If valor > vlMaximo Then vlMaximo = valor
        
        If vlMaximo > 0 Then
            txProporcao = (btnProgress_Contorno.Width / vlMaximo)
            dblAndamento = (valor * txProporcao)
        End If
        
        Me.btnProgress_Frente.Width = dblAndamento
        dblPercent = dblAndamento / (btnProgress_Contorno.Width)
        lblProgress_Percent.Caption = VBA.Format(dblPercent, "0%")
        Me.btnProgress_DeTotal.Caption = valor & " de " & vlMaximo
        
        If dblPercent > 0.5 Then lblProgress_Percent.ForeColor = VBA.vbWhite
        
        valor = valor + 1
        If dblPercent < 0.01 Then
            Me.btnProgress_Frente.Visible = False
        Else
            If Not Me.btnProgress_Frente.Visible Then
                Me.btnProgress_Frente.Visible = True
                lblProgress_Percent.Visible = True
                Me.btnProgress_DeTotal.Visible = True
            End If
            If Not Me.btnProgress_Contorno.Visible Then Me.btnProgress_Contorno.Visible = True
        End If
    End If
End Sub

Private Sub cmdCancelar_Click()
    CancelProgress = True
    
End Sub

Private Sub cmdFechar_Click()
    On Error Resume Next
    bShowBntFechar = False
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub Comando77_Click()
    Call IncremensaProgressoBar(-1)
End Sub

Private Sub lblDet_Click(): Call Detalhe_Click: End Sub

Private Sub Detalhe_Click()
    On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveYes
End Sub

Sub salvaDefault()
    If m_idefualtPagina <> 0 Then Me.Guia.value = m_idefualtPagina
End Sub

Public Property Get UpdateOnlyMensagens() As Boolean
    UpdateOnlyMensagens = m_bUpdateOnlyMensagens
End Property

Public Property Let UpdateOnlyMensagens(ByVal bUpdateOnlyMensagens As Boolean)
    m_bUpdateOnlyMensagens = bUpdateOnlyMensagens
End Property

Private Sub Form_Timer()
    Dim contagemRegressiva As Integer
    contagemRegressiva = VBA.Format(datafutura - VBA.Now, "ss")
    If contagemRegressiva <= 0 Then
        Me.TimerInterval = 0
        Call DoCmd.Close(acForm, Me.Name, acSaveNo)
    Else
        Me.cmdFechar.Visible = True
        Me.cmdFechar.Caption = "Fechar (" & VBA.Format(contagemRegressiva, "00") & " seg)"
    End If
End Sub

Private Sub Form_Load()
    vtm = VBA.Time
    Call LoadProgressMensagems(Me.OpenArgs)
End Sub

Sub LoadProgressMensagems(strArgs)
    Dim arrInfo
    Dim strOK As String
    Dim vMsg
    
    If VBA.IsNull(strArgs) Then GoTo SetInfo
    If strArgs = "" Then GoTo SetInfo
    
    arrInfo = VBA.Split(strArgs, "|")
    
    Dim pMensagem As String, _
        pTitulo As String, _
        bShowBntCancelar As Boolean, _
        bShowBntFechar As Boolean, _
        Progress_Max As Long, _
        iconeStatus As eIconeStatusMessage
    Dim f, M
    
    If Me.lblDet.Caption = "*" Then Me.lblDet.Caption = ""
    
    strOK = arrInfo(0)
    
    If VBA.CVar(arrInfo(5)) = eModoProgress.EmEndamento Then
        vMsg = VBA.Split(Me.lblDet.Caption, vbNewLine)
        
        If Me.lblDet.Caption <> "" Then
            For Each M In vMsg
               f = f & VBA.Left(VBA.CStr(M) & VBA.String(55, "."), 55) & "...OK" & vbNewLine
            Next
        End If
    End If
    
    pMensagem = f & arrInfo(0)
    
    pTitulo = arrInfo(1)
    bShowBntCancelar = arrInfo(2)
    bShowBntFechar = arrInfo(3)
    Progress_Max = VBA.CVar(arrInfo(4))
    iconeStatus = VBA.CVar(arrInfo(6))
        
    Call defineIconeStatus(iconeStatus)
    
    If UBound(arrInfo) = 7 Then
        vtm = VBA.Time
        datafutura = VBA.DateAdd("s", VBA.CInt(arrInfo(7)), VBA.Now())
        Me.TimerInterval = 1000
    End If
SetInfo:
    
    Me.cmdCancelar.Visible = bShowBntCancelar
    Me.cmdFechar.Visible = bShowBntFechar
    Me.lblTitulo.Caption = pTitulo
    Me.lblDet.Caption = pMensagem
    On Error Resume Next
    If VBA.CVar(arrInfo(5)) <> "0" And VBA.CVar(arrInfo(5)) <> eModoProgress.EmEndamento_ProgressoFixo Then
        If VBA.IsArray(arrInfo) Then
            Call IncremensaProgressoBar(Progress_Max, VBA.CVar(arrInfo(5)))
        End If
    End If
End Sub

Private Sub defineIconeStatus(ByVal iconeStatus As eIconeStatusMessage)
    Select Case iconeStatus
    Case eIconeStatusMessage.Icone_Sucesso
        Me.imgOK.Visible = True
        Me.imgAndamento.Visible = Not Me.imgOK.Visible
        Me.imgErro.Visible = Not Me.imgOK.Visible
        Me.imgAviso.Visible = Not Me.imgOK.Visible
        With Me.imgOK
            .Top = Me.imgAndamento.Top
            .Left = Me.imgAndamento.Left
        End With
    Case eIconeStatusMessage.Icone_Erro
        Me.imgErro.Visible = True
        Me.imgAndamento.Visible = Not Me.imgErro.Visible
        Me.imgOK.Visible = Not Me.imgErro.Visible
        Me.imgAviso.Visible = Not Me.imgErro.Visible
        With Me.imgErro
            .Top = Me.imgAndamento.Top
            .Left = Me.imgAndamento.Left
        End With
    Case eIconeStatusMessage.Icone_Aviso
        Me.imgAviso.Visible = True
        Me.imgAndamento.Visible = Not Me.imgAviso.Visible
        Me.imgOK.Visible = Not Me.imgAviso.Visible
        Me.imgErro.Visible = Not Me.imgAviso.Visible
        With Me.imgAviso
            .Top = Me.imgAndamento.Top
            .Left = Me.imgAndamento.Left
        End With
    Case Else
        Me.imgAndamento.Visible = True
        Me.imgOK.Visible = Not Me.imgAndamento.Visible
        Me.imgErro.Visible = Not Me.imgAndamento.Visible
        Me.imgAviso.Visible = Not Me.imgAndamento.Visible
    End Select
End Sub


Public Property Get tipoVisualizacao() As Integer
    tipoVisualizacao = m_itipoVisualizacao
End Property

Public Property Let tipoVisualizacao(ByVal itipoVisualizacao As Integer)
    m_itipoVisualizacao = itipoVisualizacao
End Property

Private Sub lvEvents_ColumnClick(ByVal ColumnHeader As Object)
    Debug.Print ColumnHeader.Text, ColumnHeader.Width
End Sub

Public Property Get StartTime() As Date
    StartTime = m_dtstartTime
End Property

Public Property Let StartTime(ByVal dtstartTime As Date)
    m_dtstartTime = dtstartTime
End Property

Public Property Get defualtPagina() As Integer
    defualtPagina = m_idefualtPagina
End Property

Public Property Let defualtPagina(ByVal idefualtPagina As Integer)
    m_idefualtPagina = idefualtPagina
End Property

Public Property Get MaxProgress() As Long
    MaxProgress = m_lMaxProgress
End Property

Public Property Let MaxProgress(ByVal lMaxProgress As Long)
    m_lMaxProgress = lMaxProgress
End Property

Private Sub lblTitulo_Click()
    MsgBox Me.lblDet.Caption
End Sub
