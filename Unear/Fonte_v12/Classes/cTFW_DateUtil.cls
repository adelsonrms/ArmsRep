VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTFW_DateUtil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' MÓDULO           : TFWAPP.AuxDate
' TIPO             : Módulo
' DATA/HORA        : 11/03/2015 10:24
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Contem funções uteis para uso de datas
'---------------------------------------------------------------------------------------

' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora       Autor     Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      11/03/2015 10:24 10:24 - Adelson   Criação/Atualização do procedimento
'---------------------------------------------------------------------------------------

'Retorna o proximo dia util após uma data
Function DiaUtil(Verdata As Date) As Boolean
    On Error Resume Next
    DiaUtil = Not (VBA.Weekday(Verdata) = 1 Or VBA.Weekday(Verdata) = 7 Or Not VBA.IsNull(DLookup("data", "tblFeriados", "data = #" & VBA.Format(Verdata, "mm/dd/yyyy") & "#")))
End Function

'Retorna o proximo dia util após uma data
Function ProximoDiaUtil(pData As Date, Optional offset As Integer = 1) As Date
    On Error Resume Next
    Dim pProxData As Date
    Dim i As Integer
    
    If offset < 0 Then
    
        pData = VBA.DateAdd("d", -1, pData)
        
        Do Until (Not (VBA.Weekday(pData) = 1 Or _
                   VBA.Weekday(pData) = 7 Or _
                   Not VBA.IsNull(DLookup("data", "tblFeriados", "data = #" & VBA.Format(pData, "mm/dd/yyyy") & "#"))))
                VBA.DoEvents
                pData = VBA.DateAdd("d", -1, pData)
        Loop
        pProxData = pData
    Else
    
        pProxData = VBA.DateAdd("d", 1, pData)
        
        Do Until (Not (VBA.Weekday(pProxData) = 1 Or _
                       VBA.Weekday(pProxData) = 7 Or _
                       Not VBA.IsNull(DLookup("data", "tblFeriados", "data = #" & VBA.Format(pProxData, "mm/dd/yyyy") & "#"))))
            VBA.DoEvents
            pProxData = VBA.DateAdd("d", 1, pProxData)
        Loop
    End If
    
    ProximoDiaUtil = pProxData
End Function

'Retorna o ultimo dia util de um mes
Function UltimoDiaUtilMes(Optional dtBase) As Date
    Dim iDay   As Integer
    Dim dtDay  As Date
    Dim dtLastDay As Date, i As Integer
    Dim dtRef  As Date
    If VBA.IsMissing(dtBase) Then dtRef = Date Else dtRef = VBA.DateValue(dtBase)
    dtLastDay = VBA.DateAdd("d", -1, VBA.DateSerial(VBA.Year(dtRef), VBA.Month(dtRef) + 1, 1))
    For i = VBA.Day(dtLastDay) To 1 Step -1
        dtDay = VBA.DateSerial(VBA.Year(dtRef), VBA.Month(dtLastDay), i)
        If VBA.Weekday(dtDay, VBA.vbMonday) < 6 Then
            UltimoDiaUtilMes = dtDay
            Exit Function
        End If
    Next
End Function

'Retorna o ultimo dia util do ano
Function PrimeiroDiaDoAno(Optional dtBase)
    Dim dtRef  As Date
    If VBA.IsMissing(dtBase) Then dtRef = VBA.Date Else dtRef = VBA.DateValue(dtBase)
    dtRef = VBA.DateSerial(VBA.Year(dtRef), 1, 1)
    PrimeiroDiaDoAno = dtRef
End Function

Function VoltaMes(Optional dtBase)
    Dim dtRef  As Date
    If VBA.IsMissing(dtBase) Then dtRef = VBA.Date Else dtRef = VBA.DateValue(dtBase)
    dtRef = VBA.DateAdd("m", -1, dtRef)
    VoltaMes = dtRef
End Function

Function VoltaAno(Optional dtBase)
    Dim dtRef  As Date
    If VBA.IsMissing(dtBase) Then dtRef = VBA.Date Else dtRef = VBA.DateValue(dtBase)
    dtRef = VBA.DateSerial(VBA.Year(dtRef) - 1, VBA.Month(dtRef), VBA.Day(dtRef))
    VoltaAno = dtRef
End Function


Function VoltaSemana(Optional dtBase)
    Dim dtRef  As Date
    If VBA.IsMissing(dtBase) Then dtRef = VBA.Date Else dtRef = VBA.DateValue(dtBase)
    dtRef = VBA.DateAdd("d", -7, dtRef)
    VoltaSemana = dtRef
End Function

Function AvancaSemana(Optional dtBase)
    Dim dtRef  As Date
    If VBA.IsMissing(dtBase) Then dtRef = VBA.Date Else dtRef = VBA.DateValue(dtBase)
    dtRef = VBA.DateAdd("d", 7, dtRef)
    AvancaSemana = dtRef
End Function


'Retorna o ultimo dia util do ano
Function PrimeiroDiaDoMes(Optional dtBase)
    Dim dtRef  As Date
    If VBA.IsMissing(dtBase) Then dtRef = VBA.Date Else dtRef = VBA.DateValue(dtBase)
    dtRef = VBA.DateSerial(VBA.Year(dtRef), VBA.Month(dtRef), 1)
    PrimeiroDiaDoMes = dtRef
End Function

'Retorna o ultimo dia util do ano
Function UltimoDiaDoMes(Optional dtBase)
    Dim dtRef  As Date
    If VBA.IsMissing(dtBase) Then dtRef = VBA.Date Else dtRef = VBA.DateValue(dtBase)
    dtRef = VBA.DateAdd("m", 1, dtRef) + 1 '
    dtRef = VBA.DateSerial(VBA.Year(dtRef), VBA.Month(dtRef), 1) - 1
    UltimoDiaDoMes = dtRef
End Function

'Retorna o ultimo dia util do ano
Function UltimoDiaAno(Optional dtBase)
    Dim dtRef  As Date
    If VBA.IsMissing(dtBase) Then dtRef = Date Else dtRef = VBA.DateValue(dtBase)
    dtRef = VBA.DateSerial(VBA.Year(dtRef), 12, 31)
    UltimoDiaAno = dtRef
End Function

'Recupera o primeiro dia da semana (começando em Segunda) de uma dataBase
Function PrimeiroDiaDaSemana(Optional dtBase)
    Dim dtRef  As Date
    If VBA.IsMissing(dtBase) Then dtRef = VBA.Date Else dtRef = VBA.DateValue(dtBase)
    dtRef = VBA.DateAdd("d", -(VBA.Weekday(dtRef, vbMonday) - 1), dtRef)
    PrimeiroDiaDaSemana = dtRef
End Function

'Recupera o primeiro dia da semana (começando em Segunda) de uma dataBase
Function UltimoDiaDaSemana(Optional dtBase)
    Dim dtRef  As Date
    If VBA.IsMissing(dtBase) Then dtRef = VBA.Date Else dtRef = VBA.DateValue(dtBase)
    dtRef = PrimeiroDiaDaSemana(AvancaSemana(dtRef)) - 3
    UltimoDiaDaSemana = dtRef
End Function

'Retorna o ultimo dia util do ano anterior
Function UltimoDiaAnoAnterior(Optional dtBase)
    Dim dtRef  As Date
    If VBA.IsMissing(dtBase) Then dtRef = Date Else dtRef = VBA.DateValue(dtBase)
    dtRef = VBA.DateSerial(VBA.Year(dtRef) - 1, 12, 31)
    UltimoDiaAnoAnterior = dtRef
End Function

'Primeiro dia util do ano
Function PrimeiroDiaUtilAno(Optional dtBase)
    Dim dtRef  As Date
    If VBA.IsMissing(dtBase) Then dtRef = Date Else dtRef = VBA.DateValue(dtBase)
    dtRef = VBA.DateSerial(VBA.Year(dtRef), 1, 1)
    PrimeiroDiaUtilAno = dtRef
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : AuxDate.ContarDiasUteis()
' TIPO             : Function
' DATA/HORA        : 11/03/2015 10:26
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Retorna a quantidade de dias uteis dentro de um intervalo de datas
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      11/03/2015 10:26    Adelson         Criação/Atualização do procedimento
'---------------------------------------------------------------------------------------
Function ContarDiasUteis(dataInicio, Optional dataFim) As Long
    Dim pProxData As Date
    Dim count  As Long

10  On Error GoTo ContarDiasUteis_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
20  Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Function AuxDate.ContarDiasUteis()"
    '----------------------------------------------------------------------------------------------------

30  If VBA.IsMissing(dataFim) Then dataFim = Date
40  If Not VBA.IsDate(dataFim) Then Exit Function
50  dataFim = VBA.DateValue(dataFim)
60  dataInicio = VBA.DateValue(dataInicio)
70  pProxData = dataInicio

80  Do Until pProxData > dataFim
90      VBA.DoEvents
100     If (Not (VBA.Weekday(pProxData) = 1 Or _
                 VBA.Weekday(pProxData) = 7 Or _
                 Not VBA.IsNull(DLookup("data", "tblFeriados", "data = #" & VBA.Format(pProxData, "mm/dd/yyyy") & "#")))) Then
110         count = count + 1
120     End If
130     pProxData = VBA.DateAdd("d", 1, pProxData)
140 Loop
150 ContarDiasUteis = count

Fim:
160 On Error GoTo 0
170 Exit Function

ContarDiasUteis_Error:
180 If VBA.Err <> 0 Then
190     lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
200     Debug.Print "Erro ao determinar a quantidade de dias uteis entre as datas. VBA.Error > " & strErrorMessagem
210 End If
    GoTo Fim:
    'Debug Mode
220 Resume
End Function

