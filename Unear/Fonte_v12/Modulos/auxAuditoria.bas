Attribute VB_Name = "auxAuditoria"
Option Compare Database
Public cOldValues As Variant
Public cNewValues As Variant


Sub inserirLogAlteracoes(ByVal formulario As Form, _
                         Optional ByVal rowIDTabela As Long, _
                         Optional registro As Variant, _
                         Optional pTipoAcao As String = "Alteração", _
                         Optional pCodAlteracao As String, Optional vOldValues As Variant)

    Dim rs     As Object ' Recordset
    Dim arrDados As Variant
    Dim contador As Integer

    On Error GoTo inserirLogAlteracoes_Error

    Set rs = CodeDb.OpenRecordset("tblLogAlteracoes")
    arrDados = verificarCamposAlterados2(vOldValues, pegarValores(formulario))
    
    If VBA.IsArray(arrDados) And UBound(arrDados) > 0 Or pTipoAcao = "Exclusão" Then
        With rs
            For contador = 0 To UBound(arrDados, 2)
                .addNew
                .Fields("NomeTabela") = formulario.RecordSource
                .Fields("IDRegistro") = registro
                .Fields("rowIDTabela") = rowIDTabela
                If pTipoAcao <> "Exclusão" Then
                    .Fields("origem") = formulario.Name & "." & arrDados(0, contador)
                    .Fields("NomeCampo") = arrDados(0, contador)
                    .Fields("ValorAntigo") = arrDados(1, contador)
                    .Fields("ValorNovo") = arrDados(2, contador)
                Else
                    .Fields("origem") = formulario.Name
                End If
                .Fields("TipoAcao") = pTipoAcao
                .Fields("UsuarioAlteracao") = VBA.Environ("ComputerName") & "\" & VBA.Environ("Username")
                .Fields("codMarcacao") = pCodAlteracao
                .Update
            Next contador
        End With
    End If

    On Error GoTo 0
    Exit Sub

inserirLogAlteracoes_Error:
    If VBA.Err <> 0 Then Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxForm.inserirLogAlteracoes()", VBA.Erl)
    Exit Sub

End Sub

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : AuxForm.verificarCamposAlterados()
' TIPO             : Function
' DATA/HORA        : 02/02/2015 17:49
' CONSULTOR        : (Jonathan)/ TECNUN - Jonathan Rocha Figueroa Dantas
' CONTATO          :  jonathan@tecun.com.br
' DESCRIÇÃO        : Esta rotina tem como objetivo devolver um array com os
'                    campos que tiveram seus valores alterados dentro de um formulario acoplado....
'---------------------------------------------------------------------------------------
Function verificarCamposAlterados(ByVal formulario As Form, ParamArray NomeControles() As Variant) As Variant

          Dim controle As Control
          Dim arrDados() As Variant
          Dim contador As Integer

10        On Error GoTo verificarCamposAlterados_Error

20        ReDim arrDados(0 To 2, 0 To 100)

30        If VBA.IsArray(NomeControles) And UBound(NomeControles) > 0 Then
40            For contador = 0 To UBound(NomeControles)
50                If formulario.Controls(NomeControles(contador)).OldValue <> formulario.Controls(NomeControles(contador)).value Then
60                    arrDados(0, contador) = NomeControles(contador)
70                    arrDados(1, contador) = formulario.Controls(NomeControles(contador)).OldValue
80                    arrDados(2, contador) = formulario.Controls(NomeControles(contador)).value
90                End If
100           Next
110           ReDim Preserve arrDados(0 To 2, 0 To contador)
120           verificarCamposAlterados = arrDados
130           Exit Function
140       End If
150       For Each controle In formulario.Controls
160           If controle.controlType = acComboBox Or controle.controlType = acTextBox Or controle.controlType = acCheckBox Then
170               If Not VBA.IsArray(controle.OldValue) Then
180                   If controle.OldValue <> controle.value Then
190                       arrDados(0, contador) = controle.Name
200                       arrDados(1, contador) = controle.OldValue
210                       arrDados(2, contador) = controle.value
220                       contador = contador + 1
230                   End If
240               End If
250           End If
260       Next controle

270       If contador <> 0 Then
280           ReDim Preserve arrDados(0 To 2, 0 To contador - 1)
290           verificarCamposAlterados = arrDados
300       Else
310           ReDim arrDados(0 To 0, 0 To 0)
320           verificarCamposAlterados = arrDados
330           Exit Function
340       End If


350       On Error GoTo 0
360       Exit Function

verificarCamposAlterados_Error:
370       If VBA.Err <> 0 Then Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxForm.verificarCamposAlterados()", VBA.Erl)
380       Exit Function
390       Resume
End Function

Function verificarCamposAlterados2(oldValues As Variant, newValues As Variant) As Variant
    Dim i As Integer, arrDados As Variant, contador As Integer
10  ReDim arrDados(0 To 2, 0 To UBound(oldValues, 2))
20  For i = 0 To UBound(oldValues, 2)
30      If Nz(oldValues(1, i)) <> Nz(newValues(1, i)) Then
40          arrDados(0, contador) = newValues(0, i)
50          arrDados(1, contador) = oldValues(1, i)
60          arrDados(2, contador) = newValues(1, i)
            contador = contador + 1
70      End If
80  Next
90  If contador <> 0 Then
100     ReDim Preserve arrDados(0 To 2, 0 To contador - 1)
110     verificarCamposAlterados2 = arrDados
120 Else
130     ReDim arrDados(0 To 0, 0 To 0)
140     verificarCamposAlterados2 = arrDados
150     Exit Function
160 End If
End Function

Sub sinalizaAlterados(oldValues As Variant, frm As Object, Optional prop As String = "Value")
    Dim i As Integer, arrDados As Variant, contador As Integer
    Dim newValues As Variant
    newValues = pegarValores(frm, prop)
20  For i = 0 To UBound(oldValues, 2)
        frm.Controls(oldValues(0, i)).BorderColor = VBA.vbBlack
30      If Nz(oldValues(1, i)) <> Nz(newValues(1, i)) Then frm.Controls(oldValues(0, i)).BorderColor = VBA.vbRed
80  Next
End Sub

Sub LimparFiltros(fParent As Object)
    Dim ctr As Object
10  For Each ctr In fParent.Controls
        Select Case ctr.controlType
        Case acComboBox, acTextBox: ctr.value = VBA.vbNullString
        Case acCheckBox: ctr.value = 0
        End Select
90  Next ctr
End Sub

Function contarControlesComValues(frm As Object) As Long
    Dim ctr As Object, contador As Integer
10  For Each ctr In frm.Controls
20      If ctr.controlType = acComboBox Or _
           ctr.controlType = acTextBox Or _
           ctr.controlType = acCheckBox Then
60              contador = contador + 1
80      End If
90  Next ctr
    contarControlesComValues = contador
End Function

Function pegarValores(ParentControls As Object, Optional propValue As String = "Value")
    Dim controle As Control
    Dim arrDados() As Variant
    Dim contador As Integer
    Dim i As Integer

10  ReDim arrDados(0 To 1, 0 To ParentControls.Controls.count)

20  For Each controle In ParentControls.Controls
30      If controle.controlType = acComboBox Or _
           controle.controlType = acTextBox Or _
           controle.controlType = acCheckBox Then
40          If Not VBA.IsArray(controle.value) Then
50              arrDados(0, contador) = controle.Name
60              On Error Resume Next
                If Screen.ActiveControl.Name = controle.Name Then
                    arrDados(1, contador) = Nz(controle.Text)
                Else
                    If Not controle.controlType = acCheckBox Then
                    If controle.InputMask <> "" Then
                        If controle.InputMask = "!\(99"") ""!9900\-0000;;_" Then
                            arrDados(1, contador) = VBA.Format(Nz(controle.value), "(0) 0000-0000")
                        ElseIf controle.InputMask = "!\(99"") ""!99000\-0000;;_" Then
                            arrDados(1, contador) = VBA.Format(Nz(controle.value), "(0) 0 0000-0000")
                        Else
                            arrDados(1, contador) = VBA.Format(Nz(controle.value), controle.InputMask)
                        End If
                    Else
                        arrDados(1, contador) = Nz(controle.value)
                    End If
                    Else
                        arrDados(1, contador) = Nz(controle.value)
                    End If
                End If
                If arrDados(1, contador) <> "" Then contador = contador + 1
160         End If
170     End If
180 Next controle

190 If contador <> 0 Then
200     ReDim Preserve arrDados(0 To 1, 0 To contador - 1)
210     pegarValores = arrDados
220 Else
230     ReDim arrDados(0 To 0, 0 To 0)
240     pegarValores = VBA.vbEmpty
250     Exit Function
260 End If
End Function

Function pegarValoresRecortset(source As Variant, Optional ChaveCampo As String = "ID", Optional ChaveValor)
    Dim controle As Object
    Dim arrDados() As Variant
    Dim contador As Integer
    Dim i As Integer
    Dim rsSource As Object
    Dim strMultValores As String
    Dim rs2 As Object ' Recordset2

10  Set rsSource = CodeDb.OpenRecordset("SELECT * FROM " & source & " WHERE " & ChaveCampo & " = " & ChaveValor)
20  contador = 0

30  Do While Not rsSource.EOF
40      ReDim arrDados(0 To 1, 0 To rsSource.Fields.count)
50      For Each controle In rsSource.Fields
60          arrDados(0, contador) = controle.Name
70          If VBA.IsObject(controle.value) Then
80              Set rs2 = controle.value
90              If Not rs2.EOF Then
100                 strMultValores = ""
110                 Do While Not rs2.EOF
120                     Select Case controle.Name
                        Case "codComarca"
130                         strMultValores = strMultValores & ", " & Nz(PegarDescricaoDeCodigo("tblComarca", "nomeComarca", "codComarca", rs2(0).value))
140                     End Select
150                     rs2.MoveNext
160                 Loop
170                 arrDados(1, contador) = Mid(strMultValores, 2)
180             End If
190         Else
200             arrDados(1, contador) = Nz(controle.value)
210         End If

220         Select Case controle.Name
            Case "TelFixo", "TelTrabalho", "TelCel"
230             arrDados(1, contador) = FormataTelefone(arrDados(1, contador))
240         Case "cpf_cnpj", "cpf_cnpj_favorecido"
250             arrDados(1, contador) = FormataCPFBNPJ(CStr(arrDados(1, contador)))
260         Case "EnderecoCEP"
270             arrDados(1, contador) = FormataCEP(CStr(arrDados(1, contador)))
280         End Select
290         If arrDados(1, contador) = "" Then arrDados(1, contador) = " "
300         If arrDados(1, contador) <> "" Then contador = contador + 1
310     Next
320     rsSource.MoveNext
330 Loop

340 If contador <> 0 Then
350     ReDim Preserve arrDados(0 To 1, 0 To contador - 1)
360     pegarValoresRecortset = arrDados
370 Else
380     ReDim arrDados(0 To 0, 0 To 0)
390     pegarValoresRecortset = arrDados
400     Exit Function
410 End If
End Function


'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : Sessao.RegistraLog()
' TIPO             : Sub
' DATA/HORA        : 29/04/2014 15:59
' CONSULTOR        : (ADELSON)/ TECNUN - Adelson Rosendo Marques da Silva
' CONTATO          : adelson@tecnun.com.br
' DESCRIÇÃO        : Registra uma entrada no log de eventos
'---------------------------------------------------------------------------------------
'REVISÃO
' 27/09/2015 01:44 - ARMS   Inclusão de parametro para escolher incluir ou nao o calculo de tempo no log
'---------------------------------------------------------------------------------------
Public Sub RegistraLog(pMessage As String, _
          Optional pFileOutPut As String = "", _
          Optional bStartNewFile As Boolean = False, _
          Optional bPrintDatetime As Boolean = True, _
          Optional bDebugPrint As Boolean, _
          Optional bCloseHifen As Boolean = False, _
          Optional pCaracter As String = "-", _
          Optional bShowFileInNotePad As Boolean = False, _
          Optional bIncluirCalculoTermpo As Boolean, _
          Optional bIncluirUsuario As Boolean = True, _
          Optional bOnlyTime As Boolean = False)
        
        On Error Resume Next 'Não precisamos tratar erros aqui
        
        Dim NumFileLog As Integer
        Dim msgLog As String
        Dim FileNameLog As String
        Dim sTime As String
        Dim sPathLog As String
          
10      sPathLog = CurrentProject.Path & "\Log\"
          
20      If VBA.Dir(sPathLog, VBA.vbDirectory) = "" Then Call VBA.MkDir(sPathLog)

30      If pFileOutPut <> "" Then
40            If VBA.InStr(pFileOutPut, "\") = 0 Then
50                FileNameLog = sPathLog & pFileOutPut
60            Else
70                FileNameLog = pFileOutPut
80            End If
90      Else
100           FileNameLog = sPathLog & CurrentProject.Name & ".txt"
110     End If

120     If bIncluirCalculoTermpo Then
130           sTime = PegaTempoDecorrido(VBA.Now - dtFirstTime)
140           If sTime = "00 segs" Then
150               sTime = " @ START " & sTime
160               dtLastTime = dtFirstTime
170           End If
180           msgLog = _
                  VBA.IIf(bPrintDatetime, VBA.Now & VBA.vbTab & _
                  VBA.IIf(bIncluirUsuario, ChaveUsuario & VBA.vbTab, "") & _
                  VBA.Right(Space(16) & PegaTempoDecorrido(VBA.Now - dtLastTime), 16) & " | " & Right(Space(16) & sTime, 16) & VBA.vbTab, "") & pMessage
190     Else
              msgLog = _
                  VBA.IIf(bPrintDatetime, VBA.IIf(bOnlyTime, VBA.Time, VBA.Now) & VBA.Space(1), "") & _
                  VBA.IIf(bIncluirUsuario, ChaveUsuario & VBA.Space(1), "") & _
                  pMessage

210     End If

220     NumFileLog = openRegistraLog(FileNameLog, bStartNewFile, bPrintDatetime)

230     If bCloseHifen Then Print #NumFileLog, String(300, pCaracter)
240     Print #NumFileLog, msgLog
250     If bCloseHifen Then Print #NumFileLog, String(300, pCaracter)
260     Close #NumFileLog

270     If bDebugPrint Then Debug.Print msgLog

280     dtLastTime = Time

290     If bShowFileInNotePad Then openTextView FileNameLog

End Sub

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : Sessao.openRegistraLog()
' TIPO             : Function
' DATA/HORA        : 29/04/2014 15:58
' CONSULTOR        : (ADELSON)/ TECNUN - Adelson Rosendo Marques da Silva
' CONTATO          : adelson@tecnun.com.br
' DESCRIÇÃO        : Abre um novo arquivo de logo
'---------------------------------------------------------------------------------------
Private Function openRegistraLog(pathLog As String, Optional bClear As Boolean = False, Optional PrintDatetime As Boolean = True) As Integer
    On Error Resume Next
    Dim NumFile As Integer
    NumFile = VBA.FreeFile()

    'Reinicia o log caso esteja cheio
    If VBA.Dir(pathLog) <> "" Then
        If VBA.FileSystem.FileLen(pathLog) > 300000 Then
            Name pathLog As pathLog & "_" & VBA.Format(VBA.Now, "yyyymmddhhnnss") & ".txt"
            bClear = True
        End If
    End If

    If VBA.Dir(pathLog) = "" Or bClear Then
        Open pathLog For Output As #NumFile
        Print #NumFile, "Ferramenta : " & CurrentDb.Name
        Print #NumFile, VBA.String(120, "-")
        Print #NumFile, VBA.Left("DATA / HORA" & VBA.String(VBA.Len(VBA.Now), " "), VBA.Len(VBA.Now)) & VBA.vbTab & "MAQUINA \ USUARIO"
        Print #NumFile, VBA.String(120, "-")
    Else
        Open pathLog For Append As #NumFile
    End If
    openRegistraLog = NumFile
End Function

'Formata o tempo decorrido de um processo
Function PegaTempoDecorrido(dtTimeToFormat As Date)
    Dim ret    As String
    If dtTimeToFormat = VBA.Now Then PegaTempoDecorrido = "(Start)"
    If VBA.Hour(dtTimeToFormat) > 0 Then ret = VBA.Format(VBA.Hour(dtTimeToFormat), "00") & " hrs"
    If VBA.Hour(dtTimeToFormat) = 0 And VBA.Minute(dtTimeToFormat) > 0 Then ret = ret & VBA.IIf(ret <> "", " e ", "") & VBA.Format(VBA.Minute(dtTimeToFormat), "00") & " min"
    If VBA.Hour(dtTimeToFormat) = 0 And VBA.Minute(dtTimeToFormat) > 0 And VBA.Second(dtTimeToFormat) Then ret = ret & VBA.IIf(ret <> "", " e ", "") & VBA.Format(VBA.Second(dtTimeToFormat), "00") & " segs"
    If VBA.Hour(dtTimeToFormat) = 0 And VBA.Minute(dtTimeToFormat) = 0 And VBA.Second(dtTimeToFormat) Then ret = VBA.Format(VBA.Second(dtTimeToFormat), "00") & " segs"
    If ret = "" Then ret = "00 segs"
    PegaTempoDecorrido = ret
End Function

'Exibe o conteudo texto no Notepad
Private Sub openTextView(source As String)
    If VBA.Dir(source) <> "" Then
        VBA.Shell "notepad.exe " & source, VBA.vbMaximizedFocus
    End If
End Sub

'Formata o tamanho de um arquivo em bytes
Function formatFileSize(size As Long, Optional bIncluirUnidade As Boolean = True)
    Dim bUnidade As String
    Dim lngSize As Double

    Select Case size
        Case Is >= 1073741824: lngSize = VBA.Round(((size / 1024) / 1024) / 1024, 2): bUnidade = VBA.IIf(bIncluirUnidade, " GB", "")
        Case Is >= 1048576: lngSize = VBA.Round((size / 1024) / 1024, 2): bUnidade = VBA.IIf(bIncluirUnidade, " MB", "")
        Case Is >= 1024 And size < (size * 1024): lngSize = VBA.Round(size / 1024, 2): bUnidade = VBA.IIf(bIncluirUnidade, " KB", "")
        Case Is < 1024: lngSize = size: bUnidade = VBA.IIf(bIncluirUnidade, " Bytes", "")
    End Select

    formatFileSize = lngSize
    If bIncluirUnidade Then formatFileSize = Right(Space(8) & lngSize & bUnidade, 12)

End Function

Function FileSize(pFilePath As String) As Long
    If FileExists(pFilePath) Then FileSize = FileSystem.FileLen(pFilePath)
End Function

Function FileDateTime(pFilePath As String) As Date
    If FileExists(pFilePath) Then FileDateTime = FileSystem.FileDateTime(pFilePath)
End Function

Function getHostPath() As Object
    Dim oApp   As Object
    Set oApp = Access.Application
    If oApp.Name = "Microsoft Excel" Then
        Set getHostPath = oApp.ThisWorkbook
    Else
        Set getHostPath = oApp.CurrentProject
    End If
End Function

