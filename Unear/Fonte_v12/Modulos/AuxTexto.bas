Attribute VB_Name = "AuxTexto"
Option Compare Database
Option Explicit

'********************PROPRIEDADES********************

Private m_RegExp        As Object 'VBScript_RegExp_55.RegExp

'---------------------------------------------------------------------------------------
' Modulo....: Publicas / Módulo
' Rotina....: RemoverAcentos / Function
' Autor.....: Fernando Fernandes
' Contato...: fernando@tecnun.com.br
' Data......: 08/04/2015
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina para remover os acentos dos caracteres de uma string, com a opção
'             para a remoção da respectiva letra
'---------------------------------------------------------------------------------------
Public Function RemoverAcentos(ByVal Texto As String, Optional ByVal RemoverLetrasAcentuadas As Boolean = False)
On Error GoTo TratarErro
Dim dicAcentos  As Object
Dim keyAcentos  As Variant
Dim ComAcento   As Variant
Dim SemAcento   As Variant
Dim letra       As String
Dim i           As Long
    
    Set dicAcentos = VBA.CreateObject("Scripting.Dictionary")
    
    ComAcento = Array("à", "á", "â", "ã", "ä", "è", "é", "ê", "ë", "ì", "í", "î", "ï", "ò", "ó", "ô", "õ", "ö", _
                      "ù", "ú", "û", "ü", "À", "Á", "Â", "Ã", "Ä", "È", "É", "Ê", "Ë", "Ì", "Í", "Î", _
                      "Ò", "Ó", "Ô", "Õ", "Ö", "Ù", "Ú", "Û", "Ü", "ç", "Ç", "ñ", "Ñ")
    SemAcento = Array("a", "a", "a", "a", "a", "e", "e", "e", "e", "i", "i", "i", "i", "o", "o", "o", "o", "o", _
                      "u", "u", "u", "u", "A", "A", "A", "A", "A", "E", "E", "E", "E", "I", "I", "I", _
                      "O", "O", "O", "O", "O", "U", "U", "U", "U", "c", "C", "n", "N")
                      
    For i = LBound(ComAcento, 1) To UBound(ComAcento, 1)
        dicAcentos.Add ComAcento(i), SemAcento(i)
    Next i
    
    For Each keyAcentos In dicAcentos.Keys
        If RemoverLetrasAcentuadas Then
            Texto = VBA.Replace(Texto, VBA.CStr(keyAcentos), VBA.vbNullString, 1, -1, VBA.vbBinaryCompare)
        Else
            Texto = VBA.Replace(Texto, VBA.CStr(keyAcentos), dicAcentos(keyAcentos), 1, -1, VBA.vbBinaryCompare)
        End If
    Next keyAcentos
    
    RemoverAcentos = Texto
    
On Error GoTo 0
Exit Function
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "Publicas.RemoverAcentos()", Erl)
End Function

'---------------------------------------------------------------------------------------
' Modulo....: Publicas / Módulo
' Rotina....: RemoverQuebrasDeLinha / Function
' Autor.....: Fernando Fernandes
' Contato...: fenrando@tecnun.com.br
' Data......: 08/04/2015
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina para remover os caracteres de quebra de linha de uma string
'---------------------------------------------------------------------------------------
Public Function RemoverQuebrasDeLinha(ByVal Texto As String) As String
On Error GoTo TratarErro
Dim arrQuebras  As Variant
Dim cnt         As Long

    arrQuebras = Array(Chr(10), VBA.Chr(13), VBA.vbNewLine, VBA.vbCrLf, VBA.vbCr, VBA.vbLf)
    For cnt = LBound(arrQuebras, 1) To UBound(arrQuebras, 1) Step 1
        Texto = VBA.Trim(VBA.Replace(Texto, arrQuebras(cnt), VBA.vbNullString))
    Next cnt
    RemoverQuebrasDeLinha = Texto
    
On Error GoTo 0
Exit Function
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "Importacao.RemoverQuebrasDeLinha", Erl)
End Function

Public Function Converter_Datas(ByVal Texto As String, ByVal dtRef As Date) As String
On Error GoTo TratarErro
Dim Aux         As String

    Aux = VBA.Replace(Texto, "[MES]", VBA.UCase(VBA.Format(dtRef, "MMMM")))
    Aux = VBA.Replace(Aux, "[ANO]", VBA.Format(dtRef, "YYYY"))
    Converter_Datas = Aux
On Error GoTo 0
Exit Function
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxTexto.Converter_Datas", Erl)
End Function

Public Function Converter_Mes_Em_Numero(ByVal sMes As String) As Byte
On Error Resume Next 'Resume next necessario, pois o texto pode não ser um mês válido
Dim dicMeses        As Object
Dim arrMesesBRAAbb  As Variant
Dim arrMesesBRA     As Variant
Dim arrMesesENGAbb  As Variant
Dim arrMesesENG     As Variant
Dim cntMeses        As Long

    Set dicMeses = VBA.CreateObject("Scripting.Dictionary")
    

    arrMesesBRAAbb = Array("JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ")
    arrMesesBRA = Array("JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", _
                        "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO")
                        
    arrMesesENGAbb = Array("JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC")
    arrMesesENG = Array("JANUARY", "FEBRUARY", "MARCH", "APRIL", "MAY", "JUNE", _
                        "JULY", "AUGUST", "SEPTEMBER", "OCTOBER", "NOVEMBER", "DECEMBER")

    For cntMeses = LBound(arrMesesBRA, 1) To UBound(arrMesesBRA, 1)
    
        If Not dicMeses.Exists(arrMesesBRAAbb) Then dicMeses.Add arrMesesBRAAbb(cntMeses), cntMeses + 1
        If Not dicMeses.Exists(arrMesesBRA) Then dicMeses.Add arrMesesBRA(cntMeses), cntMeses + 1
        
        If Not dicMeses.Exists(arrMesesENGAbb) Then dicMeses.Add arrMesesENGAbb(cntMeses), cntMeses + 1
        If Not dicMeses.Exists(arrMesesENG) Then dicMeses.Add arrMesesENG(cntMeses), cntMeses + 1

    Next cntMeses
    sMes = VBA.UCase(sMes)
    If dicMeses.Exists(sMes) Then Converter_Mes_Em_Numero = dicMeses(sMes)
    Call RemoverObjetosMemoria(arrMesesBRAAbb, arrMesesBRA, arrMesesENGAbb, arrMesesENG, dicMeses)
    
End Function

Public Function RemoverCaracteres(ByVal Aux As String, Optional ByVal Substituto As String = VBA.vbNullString) As String
Dim arrCaracteres       As Variant
Dim contador            As Byte

    arrCaracteres = Array("/", "?", "<", ">", "\", ":", "*", "|", "", """", ".")
    
    For contador = 0 To UBound(arrCaracteres) Step 1
        Aux = VBA.Replace(Aux, arrCaracteres(contador), Substituto)
    Next contador
    RemoverCaracteres = Aux
End Function


Private Property Get RegExp() As Object ' VBScript_RegExp_55.RegExp
    Set RegExp = m_RegExp
End Property
Private Property Set RegExp(ByVal valor As Object) ' VBScript_RegExp_55.RegExp)
    Set m_RegExp = valor
End Property

'---------------------------------------------------------------------------------------
' Modulo....: Publicas / Módulo
' Rotina....: IsLinhaMatch() / Function
' Autor.....: Jefferson
' Contato...: jefferson@tecnun.com.br
' Data......: 09/11/2012
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina para verificar se uma linha Corresponde a um padrao, usando regex
'---------------------------------------------------------------------------------------
Public Function IsLinhaMatch(ByVal linha As String, ParamArray Padroes() As Variant) As Boolean
    Dim Resultado As Boolean
    Dim contador As Byte

On Error GoTo TrataErro
                                          'New VBScript_RegExp_55.RegExp
    If RegExp Is Nothing Then Set RegExp = VBA.CreateObject("VBScript.RegExp")

    With RegExp
        Padroes = AuxArray.Acertar_Array_Parametros(Padroes)
        For contador = 0 To UBound(Padroes) Step 1
            If Not Padroes(contador) = VBA.vbNullString Then
                .Pattern = Padroes(contador)
                If .test(linha) Then
                    Resultado = True
                    Exit For
                End If
            End If
        Next contador
    End With
    IsLinhaMatch = Resultado
Exit Function
TrataErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxTexto.IsLinhaMatch()", Erl)
End Function

Public Function PegarTexto_Regex(ByVal linha As String, ParamArray Padroes() As Variant) As String
    On Error GoTo TratarErro
    Dim TextoAux As String
    Dim contador As Integer
    
        If RegExp Is Nothing Then Set RegExp = VBA.CreateObject("VBScript.RegExp")
        With RegExp
            For contador = 0 To UBound(Padroes) Step 1
                If Not Padroes(contador) = VBA.vbNullString Then
                    .Pattern = Padroes(contador)
                    If .test(linha) Then
                        TextoAux = VBA.Trim(.Execute(linha).item(0).value)
                    End If
                End If
            Next contador
        End With
        PegarTexto_Regex = TextoAux
    
    On Error GoTo 0
    Exit Function
TratarErro:
        Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "Publicas.PegarTexto_Regex", Erl)
    Resume
End Function

Function RegexValidaData()
    Dim separadores As String
    Dim nomeMeses As Variant
    Dim Regex_31_Dias As String
    Dim Regex_29_30_Dias As String
    Dim RegEx_29_Fev As String
    Dim RegEx_Demais As String
    Dim Regex_Ano As String
    Dim RegEx_Dia_Mes_Ano As String
    Dim RegEx_Mes As String
    Dim RegEx_Mes_Ano As String
    Dim Regex_Ano_4_Digitos As String
    
    nomeMeses = RegexNomeMesesDatas
    separadores = "(\/|-|\.)?"
    
    Regex_31_Dias = "(?:(?:31" & separadores & "(?:0[13578]|1?[02]|(?:" & nomeMeses(0) & "))))"
    Regex_29_30_Dias = "(?:(?:29|30)" & separadores & "(?:0?[1,3-9]|1[0-2]|(?:" & nomeMeses(1) & ")))"
    RegEx_29_Fev = "(?:29" & separadores & "(?:0?2|(?:fev)))"
    RegEx_Demais = "(?:0?[1-9]|1\d|2[0-8])" & separadores & "(?:(?:0?[1-9]|(?:" & nomeMeses(2) & "))|(?:1[0-2]|(?:" & nomeMeses(3) & ")))"
    Regex_Ano = "((?:(?:1[6-9]|[2-9]\d)?\d{2})|\d{2})" 'Ano com 2 ou 4 digitos
    Regex_Ano_4_Digitos = "((?:(?:1[6-9]|[2-9]\d)\d{2}))"
    Regex_Ano_4_Digitos = "((?:(?:1[6-9]|20)\d{2}))" 'Anos validos : de 1600 a 2099
    
    RegEx_Mes = "(?:(?:0?[1-9]|(?:" & nomeMeses(3) & "))|(?:1[0-2]|(?:" & nomeMeses(4) & ")))"
    
    RegEx_Dia_Mes_Ano = "(" & Regex_31_Dias & "|" & Regex_29_30_Dias & "|" & RegEx_29_Fev & "|" & RegEx_Demais & ")" & separadores & Regex_Ano
    RegEx_Mes_Ano = RegEx_Mes & separadores & Regex_Ano
    
    RegexValidaData = Array(RegEx_Mes & separadores & Regex_Ano, _
                            RegEx_Mes & separadores & Regex_Ano_4_Digitos, _
                            Regex_Ano_4_Digitos & separadores & RegEx_Mes, _
                            RegEx_Dia_Mes_Ano)
    
End Function

Function RegEx_Padrao_Ano_Mes()
    RegEx_Padrao_Ano_Mes = RegexValidaData(2)
End Function

Function RegEx_Padrao_Mes_Ano()
    RegEx_Padrao_Mes_Ano = RegexValidaData(0)
End Function

Function RegEx_Padrao_Dia_Mes_Ano()
    RegEx_Padrao_Dia_Mes_Ano = RegexValidaData(1)
End Function

Function RegexNomeMesesDatas()
    Dim mes_31 As String, mes_29_30 As String, mes_29_fev As String, todas_datas As String
    Dim todas_datas_jan_set As String
    Dim todas_datas_out_dez As String
    Dim Mes As Integer
    
    For Mes = 1 To 12
        If InStr("13578", Mes) > 0 Then mes_31 = mes_31 & "|" & LCase(VBA.MonthName(Mes, True))
        If Mes <> 2 Then mes_29_30 = mes_29_30 & "|" & VBA.LCase(VBA.MonthName(Mes, True))
        If Mes = 2 Then mes_29_fev = mes_29_fev & "|" & VBA.LCase(VBA.MonthName(Mes, True))
        If Mes < 10 Then todas_datas_jan_set = todas_datas_jan_set & "|" & VBA.LCase(VBA.MonthName(Mes, True))
        If Mes > 10 Then todas_datas_out_dez = todas_datas_out_dez & "|" & VBA.LCase(VBA.MonthName(Mes, True))
    Next
    mes_31 = Mid(mes_31, 2)
    mes_29_30 = Mid(mes_29_30, 2)
    mes_29_fev = Mid(mes_29_fev, 2)
    todas_datas_jan_set = Mid(todas_datas_jan_set, 2)
    todas_datas_out_dez = Mid(todas_datas_out_dez, 2)
    
    RegexNomeMesesDatas = Array(mes_31, mes_29_30, mes_29_fev, todas_datas_jan_set, todas_datas_out_dez)
End Function


Function DeterminaDataPorRegex(pTexto As String)
    Dim dt As Date, sData As String
    'If IsLinhaMatch(pTexto, RegEx_Padrao_Ano_Mes) Then
        sData = RemoverCaracteres(PegarTexto_Regex(pTexto, RegEx_Padrao_Ano_Mes))
        If sData <> "" Then
            If Len(sData) = 6 Then
                sData = VBA.DateSerial(VBA.Left(sData, 4), Right(sData, 2), 1)
            End If
        End If
        
        sData = RemoverCaracteres(PegarTexto_Regex(pTexto, RegEx_Padrao_Mes_Ano))
        If sData <> "" Then
            If Len(sData) = 6 Then
                sData = VBA.DateSerial(Right(sData, 4), VBA.Left(sData, 2), 1)
            End If
        End If
        If sData <> "" Then
            If VBA.IsDate(sData) Then DeterminaDataPorRegex = StrConv(VBA.Format(sData, "   MMMM/YYYY"), VBA.vbProperCase) Else DeterminaDataPorRegex = "Periodo não identificado"
        End If
End Function


Function TestaRegExp(padrao As String, Texto As String)
    Dim objRegExp As Object ' RegExp
    Dim objMatch As Object ' Match
    Dim colMatch As Object ' MatchCollection
    Dim valor As String
    'cria um objeto expressão regular
    Set objRegExp = VBA.CreateObject("VBScript.RegExp")
    
    'define o padrão - Pattern
    objRegExp.Pattern = padrao
    'define IgnoreCase
    objRegExp.IgnoreCase = True
    'define a propriedade global
    objRegExp.Global = True
    'verifica se a string pode ser comparada
    If (objRegExp.test(Texto) = True) Then
        'obtem as coincidencias
        Set colMatch = objRegExp.Execute(Texto)   'executa a busca

        For Each objMatch In colMatch
            valor = valor & " padrao encontrado na posição "
            valor = valor & objMatch.FirstIndex & ". o valor é '"
            valor = valor & objMatch.value & " '." & VBA.vbCrLf
        Next
    Else
        valor = "Comparação falhou !"
    End If
    TestaRegExp = valor
End Function
