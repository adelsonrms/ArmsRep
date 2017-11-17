Attribute VB_Name = "AuxDictionary"
Option Explicit
Option Private Module

'---------------------------------------------------------------------------------------
' Rotina....: DropDictionary() / Sub
' Contato...: fernando.fernandes@outlook.com.br
' Autor.....: Fernando Fernandes
' Empresa...: www.Planilhando.Com.Br
' Data......: 02/19/2014 (mdy)
' Descrição.: Routine that drops a whole dictionary into a worksheet, in the chosen orientation, given the worksheet and a start range
'Scripting.Dictionary
'---------------------------------------------------------------------------------------
Public Sub DropDictionary(ByRef dic As Object, _
                          ByVal Orientation As String, _
                          ByRef wsh As Object, ByVal lRow As Long, lCol As Long)
                          
On Error GoTo TreatError
Dim mtz As Variant
Dim key As Variant
Dim cnt As Long

    cnt = 1
    Select Case Orientation
        Case "H"
            ReDim mtz(1 To 1, 1 To dic.count)
            For Each key In dic
                mtz(1, cnt) = dic.item(key)
                cnt = cnt + 1
            Next key
        Case "V"
            ReDim mtz(1 To dic.count, 1 To 1)
            For Each key In dic
                mtz(cnt, 1) = dic.item(key)
                cnt = cnt + 1
            Next key
    End Select
    
    If VBA.IsArray(mtz) Then
        Call DropArray(wsh, lRow, lCol, mtz)
    End If
    
On Error GoTo 0
Exit Sub
TreatError:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxDictionary.DropDictionary()", Erl, True)
End Sub

'---------------------------------------------------------------------------------------
' Modulo....: auxArray / Módulo
' Rotina....: GetDictionaryFromArray() / Function
' Autor.....: Fernando Reis
' Contato...: fernando.fernandes@outlook.com.br
' Data......: 02/18/2014 (mdy)
' Empresa...: www.Planilhando.com.br
' Descrição.: This routine gets all the content of an array column and loads it into a dictionary
'---------------------------------------------------------------------------------------
Public Function GetDictionaryFromArray(ByRef ArrayToFilter As Variant, _
                                       ByRef dicOriginal As Object, _
                                       ByVal Header As XlYesNoGuess, _
                                       ByVal WhichField As Long, _
                                       ByVal Append As Boolean) As Object 'Scripting.Dictionary
On Error GoTo TreatError
Dim arrAux          As Variant
Dim dicAux          As Object
Dim lngContador     As Long
    Set dicAux = VBA.CreateObject("Scripting.Dictionary")
    If Append Then Set dicAux = dicOriginal
    If VBA.IsArray(ArrayToFilter) Then
        For lngContador = LBound(ArrayToFilter, 1) To UBound(ArrayToFilter, 1) Step 1
            If Header = xlYes And lngContador = LBound(ArrayToFilter, 1) Then GoTo Proximo
            If VBA.Trim(ArrayToFilter(lngContador, WhichField)) <> VBA.vbNullString And ArrayToFilter(lngContador, WhichField) <> 0 Then
                If dicAux.count = 0 Then
                    Call dicAux.Add(ArrayToFilter(lngContador, WhichField), lngContador)
                ElseIf Not dicAux.Exists(ArrayToFilter(lngContador, WhichField)) Then
                    Call dicAux.Add(ArrayToFilter(lngContador, WhichField), lngContador)
                End If
            End If
Proximo:
        Next lngContador
        Set GetDictionaryFromArray = dicAux
    Else
        Set GetDictionaryFromArray = VBA.IIf(Append, dicAux, Nothing)
        
    End If
    Call RemoverObjetosMemoria(dicAux, arrAux, lngContador)
    
On Error GoTo 0
Exit Function
TreatError:
   Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxDictionary.GetDictionaryFromArray()")
End Function

'---------------------------------------------------------------------------------------
' Modulo....: auxArray / Módulo
' Rotina....: GetDictionaryFromRange() / Function
' Autor.....: Fernando Reis
' Contato...: fernando.fernandes@outlook.com.br
' Data......: 01/04/2013 (mdy)
' Revisão...: 10/06/2014 (mdy)
' Empresa...: Planilhando
' Descrição.: This routine gets all the content of a range and loads it into a dictionary
'---------------------------------------------------------------------------------------
Public Function GetDictionaryFromRange(ByVal rngData As Object, _
                                       Optional ByVal KeyColumn As Long = 1, _
                                       Optional ByVal ItemColumn As Long = 2, _
                                       Optional ByVal blUseSecondColumn As Boolean = False) As Object 'Scripting.Dictionary
On Error GoTo TreatError
Dim arrAux          As Variant
Dim dicAux          As Object
Dim lngContador     As Long
    Set dicAux = VBA.CreateObject("Scripting.Dictionary")
    If Not rngData.Rows.count = 0 Then
        arrAux = AuxArray.GetArrayFromRange(rngData)
        If rngData.Columns.count = 1 And Not blUseSecondColumn Then
        
            For lngContador = 1 To UBound(arrAux) Step 1
            
                If VBA.Trim(arrAux(lngContador, KeyColumn)) <> VBA.vbNullString And arrAux(lngContador, KeyColumn) <> 0 Then
                
                    If dicAux.count = 0 Then
                        Call dicAux.Add(arrAux(lngContador, KeyColumn), lngContador)
                        
                    ElseIf Not dicAux.Exists(arrAux(lngContador, KeyColumn)) Then
                        Call dicAux.Add(arrAux(lngContador, KeyColumn), lngContador)
                        
                    End If
                    
                End If
                
            Next lngContador
            
        ElseIf blUseSecondColumn And rngData.Columns.count >= ItemColumn Then
        
            For lngContador = 1 To UBound(arrAux) Step 1
            
                If VBA.Trim(arrAux(lngContador, KeyColumn)) <> VBA.vbNullString And arrAux(lngContador, KeyColumn) <> 0 Then
                
                    If dicAux.count = 0 Then
                        Call dicAux.Add(arrAux(lngContador, KeyColumn), arrAux(lngContador, ItemColumn))
                        
                    ElseIf Not dicAux.Exists(arrAux(lngContador, KeyColumn)) Then
                        Call dicAux.Add(arrAux(lngContador, KeyColumn), arrAux(lngContador, ItemColumn))
                        
                    End If
                    
                End If
                
            Next lngContador
        Else
            Set GetDictionaryFromRange = Nothing
            
        End If
        
        Set GetDictionaryFromRange = dicAux
    Else
        Set GetDictionaryFromRange = Nothing
    End If
    Call RemoverObjetosMemoria(dicAux, arrAux)
    
On Error GoTo 0
Exit Function
TreatError:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxDictionary.GetDictionaryFromRange()")
End Function

'---------------------------------------------------------------------------------------
' Modulo....: AuxDictionary / Módulo
' Rotina....: GetDictionaryFromRecordSet() / Function
' Autor.....: Fernando Reis
' Contato...: fernando.fernandes@outlook.com.br
' Data......: 01/04/2013 (mdy)
' Revisão...: 10/06/2014 (mdy)
' Empresa...: Planilhando
' Descrição.: This routine gets all the content of a range and loads it into a dictionary
'---------------------------------------------------------------------------------------
Public Function GetDictionaryFromRecordSet(ByVal rs As Object, _
                                           Optional ByVal KeyField As String, _
                                           Optional ByVal ItemField As String, _
                                           Optional ByVal blUseSecondColumn As Boolean = False) As Object 'Scripting.Dictionary
On Error GoTo TreatError
Dim arrAux          As Variant
Dim dicAux          As Object
Dim lngContador     As Long
Set dicAux = VBA.CreateObject("Scripting.Dictionary")
''    If Not rs Is Nothing Then
''        If rs.RecordCount > 0 Then
''            arrAux = Conexao.PegarArray(rs)
''            If rngData.Columns.Count = 1 And Not blUseSecondColumn Then
''
''                For lngContador = 1 To UBound(arrAux) Step 1
''
''                    If VBA.Trim(arrAux(lngContador, KeyColumn)) <>VBA.vbNullString And arrAux(lngContador, KeyColumn) <> 0 Then
''
''                        If dicAux.Count = 0 Then
''                            Call dicAux.Add(arrAux(lngContador, KeyColumn), lngContador)
''
''                        ElseIf Not dicAux.Exists(arrAux(lngContador, KeyColumn)) Then
''                            Call dicAux.Add(arrAux(lngContador, KeyColumn), lngContador)
''
''                        End If
''
''                    End If
''
''                Next lngContador
''
''            ElseIf blUseSecondColumn And rngData.Columns.Count >= ItemColumn Then
''
''                For lngContador = 1 To UBound(arrAux) Step 1
''
''                    If VBA.Trim(arrAux(lngContador, KeyColumn)) <>VBA.vbNullString And arrAux(lngContador, KeyColumn) <> 0 Then
''
''                        If dicAux.Count = 0 Then
''                            Call dicAux.Add(arrAux(lngContador, KeyColumn), arrAux(lngContador, ItemColumn))
''
''                        ElseIf Not dicAux.Exists(arrAux(lngContador, KeyColumn)) Then
''                            Call dicAux.Add(arrAux(lngContador, KeyColumn), arrAux(lngContador, ItemColumn))
''
''                        End If
''
''                    End If
''
''                Next lngContador
''            Else
''                Set GetDictionaryFromRange = Nothing
''
''            End If
''        End If
''        Set GetDictionaryFromRange = dicAux
''    Else
''        Set GetDictionaryFromRange = Nothing
''    End If
''    Call RemoverObjetosMemoria(dicAux, arrAux)
    
On Error GoTo 0
Exit Function
TreatError:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxDictionary.GetDictionaryFromRecordSet()")
End Function

'---------------------------------------------------------------------------------------
' Modulo....: Publicas / Módulo
' Rotina....: CriarDicionario / Sub
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 26/08/2013
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina para criar um dicionario de acordo com as informações passadas no array
'---------------------------------------------------------------------------------------
Public Function CriarDicionario(ByVal Parametros As Variant) As Object 'Scripting.Dictionary
On Error GoTo TratarErro
Dim dicAux          As Object
Dim ContLinhas      As Long
Set dicAux = VBA.CreateObject("Scripting.Dictionary")
    For ContLinhas = LBound(Parametros, 1) To UBound(Parametros, 1) Step 1
        If Not dicAux.Exists(Parametros(ContLinhas)) Then
            Call dicAux.Add(Parametros(ContLinhas), Empty)
        End If
    Next ContLinhas
    If dicAux.count > 0 Then Set CriarDicionario = dicAux
On Error GoTo 0
Exit Function
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxDictionary.CriarDicionario", Erl)
End Function
