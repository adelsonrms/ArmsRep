'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : VBEditor.ConverterParaHTML()
' TIPO             : Function
' DATA/HORA        : 12/09/2014 00:13
' CONSULTOR        : Adelson Rosendo Marques da Silva
' CONTATO          : adelsons@gmail.com.br
' DESCRIÇÃO        : Converte um código para HTML
'---------------------------------------------------------------------------------------
Function ConverterParaHTML(strCod As String) As String
    '----------------------------------------------------------------------------------------------------
    On Error GoTo TratarErro: Const cstr_ProcedureName As String = "VBAProject.VBCodeConvert.ConverterParaHTML"
    '------------------------------------------------------------------------------------------------
    Dim rngPalavraChave As Range
    Dim Linha As String
    Dim NovaLinha As String
    Dim LinhaFinal As String
    Dim Linhas
    Dim sPalavra As String
    Dim vPalavras
    Dim iPal As Long
    Dim iLin As Long
    Dim Operador As Range
    
    Linhas = VBA.Split(strCod, vbNewLine)
    LinhaFinal = ""

    For iLin = LBound(Linhas) To UBound(Linhas)
        NovaPalavra = ""
        Linha = ""
        Linha = Linhas(iLin)
        LinhaSaidaOperador = ""
        LinhaComOperador = ""
        textomeio = ""
        textoantes = ""
        If Linha <> "" Then
            vPalavras = VBA.Split(Linha, Chr(32))
            LinhaComOperador = Linha
            For iPal = LBound(vPalavras) To UBound(vPalavras)
                sPalavra = vPalavras(iPal)
                Set rngPalavraChave = shtVBAKeyWord.Range("PalavraChave").Find(sPalavra, , , 1)
                'Se encontrar a palavra
                If Not rngPalavraChave Is Nothing Then
                    sPalavra = "<font color=blue>" & sPalavra & " </font>"
                End If
                NovaPalavra = NovaPalavra & sPalavra & " "
            Next
            'Procura Operadores e formata de vermelho
            Fim = 1
            Espaco = ""
            Do While Fim > 0
                If Left(Linha, 1) <> "" And Left(Linha, 1) <> " " Then FimEspaco = 0: Exit Do
                If Mid(Linha, Fim, 1) <> " " Then FimEspaco = 0: Exit Do
                Fim = Fim + 1
                Espaco = Espaco & "&nbsp;"
            Loop

            LinhaFinal = LinhaFinal & Espaco & NovaPalavra & " " & "<br>" & vbNewLine
        End If
    Next
    ConverterParaHTML = "<font face=""Courier New"" size=2>" & LinhaFinal & "</font>"
    
    On Error GoTo 0
Fim:
    Exit Function
TratarErro:
    If VBA.Err <> 0 Then Call VBA.Err.Raise(VBA.Rnd(), cstr_ProcedureName & " - " & VBA.Erl() & ">" & VBA.Err.Description & VBA.vbCrLf)
    GoTo Fim:
    Resume
End Function
'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : VBEditor.LinhaFormatada()
' TIPO             : Function
' DATA/HORA        : 12/09/2014 00:00
' CONSULTOR        : Adelson Rosendo Marques da Silva
' CONTATO          : adelsons@gmail.com.br
' DESCRIÇÃO        : Formata uma linha de código
'---------------------------------------------------------------------------------------
Private Function LinhaFormatada(strCod As String)
    '----------------------------------------------------------------------------------------------------
    On Error GoTo TratarErro: Const cstr_ProcedureName As String = "VBAProject.VBEditor.LinhaFormatada"
    '------------------------------------------------------------------------------------------------
    PosDepois = 1
    LinhaFormatada = ""
    Linhas = VBA.Split(strCod, vbNewLine)
    LinhaFinal = ""

    For iLin = LBound(Linhas) To UBound(Linhas)
        NovaPalavra = ""
        Linha = Linhas(iLin)
        LinhaSaidaOperador = ""
        LinhaComOperador = ""
        textomeio = ""
        textoantes = ""
        LinhaComOperador = Linha
        PosDepois = 1
        If Linha <> "" Then
            For Each Operador In shtVBAKeyWord.Range("Operadores")
                l = InStr(PosDepois, LinhaComOperador, Chr(Operador.Value))
                If l > 0 Then
                    textoantes = VBA.Trim(Mid(LinhaComOperador, 1, l - 1))
                    textodepois = VBA.Trim(Mid(LinhaComOperador, l + 1, Len(LinhaComOperador)))
                    If textodepois = "" Then
                        textodepois = Mid(LinhaComOperador, l, Len(LinhaComOperador))
                        textomeio = Mid(LinhaComOperador, l + 1, (Len(LinhaComOperador) - ((Len(textoantes)) - (Len(textodepois) - 1))))
                        textomeio = "<font color=red>" & textodepois & "</font>"
                        PosDepois = ((Len(textoantes)) + 1) + (Len(textomeio))
                        LinhaComOperador = textoantes & textomeio
                    Else
                        textomeio = Mid(LinhaComOperador, l, 1)
                        textomeio = "<font color=red>" & textomeio & "</font>"
                        PosDepois = ((Len(textoantes)) + 1) + (Len(textomeio))
                        LinhaComOperador = textoantes & textomeio & textodepois
                    End If
                End If
            Next
            LinhaFormatada = LinhaFormatada & LinhaComOperador & vbNewLine
        End If
    Next
    On Error GoTo 0
Fim:
    Exit Function
TratarErro:
    If VBA.Err <> 0 Then Call VBA.Err.Raise(VBA.Rnd(), cstr_ProcedureName & " - " & VBA.Erl() & ">" & VBA.Err.Description & VBA.vbCrLf)
    GoTo Fim:
    Resume
End Function

