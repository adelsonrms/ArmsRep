Attribute VB_Name = "AuxArray"
Option Compare Database
Option Explicit

Public Enum FilterArrayAction
    Keep = 0
    Remove = 1
End Enum

Public Enum FiltrarMatrizAcao
    Manter = 0
    Apagar = 1
End Enum

Public Enum CellProperty
    value = 0
    Formula = 1
    formular1c1 = 2
End Enum

Public Enum TypeOfTrim
    VisualBasicForApplication = 1
    WorksheetFunction = 2
    Caractere160 = 3
End Enum

Public Enum Orientation
    Vertical = 1
    Horizontal = 2
End Enum

Public Enum position
    BeforeFirst = 1
    AfterLast = 2
End Enum

'---------------------------------------------------------------------------------------
' Rotina....: Transpose() / Function
' Contato...: fernando@tecnun.com.br
' Autor.....: Jefferson Dantas
' Revisão...: Fernando Fernandes
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: This routine transposes any uni or bidimensional array
'---------------------------------------------------------------------------------------
Public Function Transpose(ByVal Matriz As Variant) As Variant
On Error GoTo TreatError
Dim lngContador     As Long
Dim lngContador1    As Long
Dim arrAux          As Variant

    If VBA.IsArray(Matriz) Then
        Select Case NumberOfDimensions(Matriz)
            Case 1
                arrAux = Matriz

            Case 2
'creating auxiliary array with inverted dimensions
                ReDim arrAux(LBound(Matriz, 2) To UBound(Matriz, 2), LBound(Matriz, 1) To UBound(Matriz, 1))

                For lngContador = LBound(Matriz, 2) To UBound(Matriz, 2) Step 1
                    For lngContador1 = LBound(Matriz, 1) To UBound(Matriz, 1) Step 1
                        arrAux(lngContador, lngContador1) = Matriz(lngContador1, lngContador)
                    Next lngContador1
                Next lngContador
        End Select
    End If

    Transpose = arrAux
On Error GoTo 0
Exit Function
TreatError:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "auxArray.Transpose()", Erl)
End Function

'---------------------------------------------------------------------------------------
' Modulo....: auxarray / Módulo
' Rotina....: GetArrayFromRange() / Function
' Autor.....: Fernando Fernandes
' Contato...: fernando.fernandes@outlook.com.br
' Data......: 12/19/2012 (mdy)
' Empresa...: Planilhando
' Descrição.: This routine creates an array from a given range of 1 or more cells.
'---------------------------------------------------------------------------------------
Public Function GetArrayFromRange(ByRef rng As Object, _
                                  Optional WhichProperty As CellProperty = CellProperty.value, _
                                  Optional WhichRow As Long = 0) As Variant
On Error GoTo TreatError
Dim arrArray(1 To 1, 1 To 1) As Variant

    With rng
        If .Cells.count = 1 Then
            Select Case WhichProperty
                Case CellProperty.value
                    arrArray(1, 1) = .value
                    GetArrayFromRange = arrArray
                Case CellProperty.Formula
                    arrArray(1, 1) = .Formula
                    GetArrayFromRange = arrArray
                Case CellProperty.formular1c1
                    arrArray(1, 1) = .formular1c1
                    GetArrayFromRange = arrArray
            End Select
        Else
            If WhichRow = 0 Then
                Select Case WhichProperty
                    Case CellProperty.value
                        GetArrayFromRange = .value
                    Case CellProperty.Formula
                        GetArrayFromRange = .Formula
                    Case CellProperty.formular1c1
                        GetArrayFromRange = .formular1c1
                End Select
            Else
                With .Rows(WhichRow)
                    Select Case WhichProperty
                        Case CellProperty.value
                            GetArrayFromRange = .value
                        Case CellProperty.Formula
                            GetArrayFromRange = .Formula
                        Case CellProperty.formular1c1
                            GetArrayFromRange = .formular1c1
                    End Select
                End With
            End If
        End If
    End With
On Error GoTo 0
Exit Function
TreatError:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "auxArray.GetArrayFromRange()", Erl)
End Function

'---------------------------------------------------------------------------------------
' Rotina....: FilterArray() / Function
' Contato...: fernando.fernandes@outlook.com.br
' Autor.....: Fernando Fernandes
' Empresa...: Planilhando
' Descrição.: This routine filters the content of any bidimensional array, with the given criterias
'             arrOriginal is the complete array
'             FilterAction (enum):
'                          Keep   => Creates a brand new array, keeping only the chosen criterias in the chosen field
'                          Remove => Creates a brand new array, removing the chosen criterias from the chosen field
'             Header determines if the array has or has not headers, to decide if the first row will be removed or not
'             lColumn is the column number within the array, the field where the filter will be based on
'---------------------------------------------------------------------------------------
Public Function FilterArray(ByVal arrOriginal As Variant, _
                            ByVal FilterAction As FilterArrayAction, _
                            ByVal Header As XlYesNoGuess, _
                            ByVal lColumn As Long, _
                            ParamArray Criterias() As Variant) As Variant
On Error GoTo TreatError
Dim arrFinal            As Variant
Dim cntOriginalArray    As Long
Dim cntFinalArray       As Long
Dim cntCriterias        As Long

    If VBA.IsArray(arrOriginal) Then
        If lColumn <= UBound(arrOriginal, 2) Then
            cntFinalArray = LBound(arrOriginal, 1)

'creating auxiliary array with same dimensions
            ReDim arrFinal(LBound(arrOriginal, 1) To UBound(arrOriginal, 1), LBound(arrOriginal, 2) To UBound(arrOriginal, 2))

            If Header = xlYes Then
                Call CopyArrayRow(arrOriginal, LBound(arrOriginal, 1), arrFinal, LBound(arrFinal, 1))
                cntFinalArray = cntFinalArray + 1
            End If

            For cntOriginalArray = cntFinalArray To UBound(arrOriginal, 1) Step 1

                For cntCriterias = LBound(Criterias, 1) To UBound(Criterias, 1) Step 1

                    If (FilterAction = Keep And arrOriginal(cntOriginalArray, lColumn) Like Criterias(cntCriterias)) Or _
                       (FilterAction = Remove And Not arrOriginal(cntOriginalArray, lColumn) Like Criterias(cntCriterias)) Then
                        Call CopyArrayRow(arrOriginal, cntOriginalArray, arrFinal, cntFinalArray)
                        cntFinalArray = cntFinalArray + 1
                    End If

                Next cntCriterias

            Next cntOriginalArray
            cntFinalArray = cntFinalArray - 1

            If cntFinalArray >= 0 Then Call ResizeArray(arrFinal, cntFinalArray)

        End If
    End If

    FilterArray = arrFinal

On Error GoTo 0
Exit Function
TreatError:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "auxArray.FilterArray()", Erl, True)
End Function

'---------------------------------------------------------------------------------------
' Rotina....: FilterArray() / Function
' Contato...: fernando.fernandes@outlook.com.br
' Autor.....: Fernando Fernandes
' Empresa...: Planilhando
' Descrição.: This routine filters the content of any bidimensional array, with the given criterias
'---------------------------------------------------------------------------------------
Public Function FiltrarMatriz(ByVal arrOriginal As Variant, _
                              ByVal FilterAction As FiltrarMatrizAcao, _
                              ByVal Header As XlYesNoGuess, _
                              ByVal lColumn As Long, _
                              ParamArray Criterias() As Variant) As Variant
On Error GoTo TreatError
Dim arrFinal            As Variant
Dim cntOriginalArray    As Long
Dim cntFinalArray       As Long
Dim cntCriterias        As Long

    If VBA.IsArray(arrOriginal) Then
        If lColumn <= UBound(arrOriginal, 2) Then
            cntFinalArray = LBound(arrOriginal, 1)

'creating auxiliary array with same dimensions
            ReDim arrFinal(LBound(arrOriginal, 1) To UBound(arrOriginal, 1), LBound(arrOriginal, 2) To UBound(arrOriginal, 2))

            If Header = xlYes Then
                Call CopyArrayRow(arrOriginal, LBound(arrOriginal, 1), arrFinal, LBound(arrFinal, 1))
                cntFinalArray = cntFinalArray + 1
            End If

            For cntOriginalArray = cntFinalArray To UBound(arrOriginal, 1) Step 1

                For cntCriterias = LBound(Criterias, 1) To UBound(Criterias, 1) Step 1

                    If (FilterAction = Manter And arrOriginal(cntOriginalArray, lColumn) Like Criterias(cntCriterias)) Or _
                       (FilterAction = Apagar And Not arrOriginal(cntOriginalArray, lColumn) Like Criterias(cntCriterias)) Then
                        Call CopyArrayRow(arrOriginal, cntOriginalArray, arrFinal, cntFinalArray)
                        cntFinalArray = cntFinalArray + 1
                    End If

                Next cntCriterias

            Next cntOriginalArray
            cntFinalArray = cntFinalArray - 1

            If cntFinalArray >= 0 Then Call ResizeArray(arrFinal, cntFinalArray)

        End If
    End If

    FiltrarMatriz = arrFinal

On Error GoTo 0
Exit Function
TreatError:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "auxArray.FilterArray()", Erl, True)
End Function
'---------------------------------------------------------------------------------------
' Rotina......: DropArray() / Sub
' Contato.....: fernando.fernandes@outlook.com.br
' Autor.......: Fernando Fernandes
' Date........: Nov/15th/2013 - Original Drop Array for usual Excel arrays, starting with index 1
' Review Date.: Oct/6th/2014  - Adapted Drop Array for usual Access arrays, considering arrays starting with index 0
' Review Date.: Mar/19th/2015 - Adapted Drop Array for limiting the number of rows os columns to drop from the array
' Empresa.....: www.Planilhando.Com.Br
' Descrição...: Routine that drops a part of or a whole array into a worksheet, given the worksheet and a start range (with row and column indexes
'               plus optional arguments to limit the number of rows and/or colummns to drop, from the array
'---------------------------------------------------------------------------------------
Public Sub DropArray(ByRef wsh As Object, ByVal lRow As Long, ByVal lCol As Long, _
                     ByRef arr As Variant, _
                     Optional numRows As Long = 0, _
                     Optional numCols As Long = 0)
On Error GoTo TreatError
Dim FinalRow As Long
Dim FinalCol As Long

    With wsh

        If LBound(arr, 1) = 1 And LBound(arr, 2) = 1 Then
            If numRows = 0 Then FinalRow = lRow + UBound(arr, 1) - 1 Else FinalRow = numRows
            If numCols = 0 Then FinalCol = lCol + UBound(arr, 2) - 1 Else FinalCol = numCols


        ElseIf LBound(arr, 1) = 0 And LBound(arr, 2) = 0 Then
            If numRows = 0 Then FinalRow = lRow + UBound(arr, 1) Else FinalRow = lRow + numRows - 1
            If numCols = 0 Then FinalCol = lCol + UBound(arr, 2) Else FinalCol = lCol + numCols - 1

        End If

        .Range(.Cells(lRow, lCol), .Cells(FinalRow, FinalCol)).value = arr

    End With

On Error GoTo 0
Exit Sub
TreatError:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "auxArray.DropArray()", Erl, True)
End Sub

'---------------------------------------------------------------------------------------
' Rotina......: NumberOfDimensions() / Sub
' Contato.....: fernando.fernandes@outlook.com.br
' Autor.......: Fernando Fernandes
' Date........: Feb/18th/2020
' Observação..: Eu sei é futuro, mas uma rotina que fala de dimensões, tinha que falar de viagem no tempo
' Empresa.....: www.Planilhando.Com.Br
' Descrição...: Retorna o número total de dimensões de uma matriz
'---------------------------------------------------------------------------------------
Public Function NumberOfDimensions(ByVal arr As Variant) As Long
On Error GoTo TreatError
Dim cnt As Long

    cnt = 1
    Do Until VBA.Err.Number <> 0
        If LBound(arr, cnt) >= 0 Then NumberOfDimensions = cnt
        cnt = cnt + 1
    Loop

On Error GoTo 0
Exit Function
TreatError:
    Exit Function
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxArray.NumberOfDimensions()", Erl, True)
End Function

'---------------------------------------------------------------------------------------
' Rotina....: GetRowFromArray() / Sub
' Contato...: Fernando.Fernandes@Outlook.com.br
' Autor.....: Fernando Fernandes
' Empresa...: www.Planilhando.Com.Br
' Descrição.: Brings a given row from an array, returns the array of that row
'---------------------------------------------------------------------------------------
Public Function GetRowFromArray(ByVal arr As Variant, _
                                ByVal RowNumber) As Variant
On Error GoTo TreatError
Dim lColumns            As Long
Dim arrAux              As Variant

    If VBA.IsArray(arr) Then
        ReDim arrAux(LBound(arr, 1) To LBound(arr, 1), LBound(arr, 2) To UBound(arr, 2))

        For lColumns = LBound(arrAux, 2) To UBound(arrAux, 2) Step 1
            arrAux(LBound(arr, 1), lColumns) = arr(RowNumber, lColumns)
        Next

    End If

    GetRowFromArray = arrAux

On Error GoTo 0
Exit Function
TreatError:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "auxArray.GetRowFromArray()", Erl, True)
End Function

'---------------------------------------------------------------------------------------
' Modulo....: auxArray / Módulo
' Rotina....: GetArrayFromDictionary() / Function
' Autor.....: Fernando Fernandes / fernando.fernandes@outlook.com.br
' Revisão...: Jefferson Dantas / jefferson@tecnun.com.br
' Data......: 02/18/2014 (mdy)
' Empresa...: Planilhando
' Descrição.: This routine gets all the content of a range and loads it into a dictionary
'             dic is the dictionary which wwill become an array
'             UseItems is the boolean to decide if the item will be broken into many columns, separated by pipe "|"
'---------------------------------------------------------------------------------------
Public Function GetArrayFromDictionary(ByVal dic As Object, _
                                       Optional ByVal UseItems As Boolean = False) As Variant
On Error GoTo TratarErro
Dim key     As Variant
Dim mtz     As Variant
Dim Aux     As Variant
Dim cnt     As Long
Dim cntCol  As Long

    cnt = 0
    If Not dic Is Nothing Then
        If dic.count > 0 Then
            For Each key In dic.Keys
                Aux = VBA.Split(dic(key), "|")
                If UseItems Then
                    If Not VBA.IsArray(mtz) Then
                        ReDim mtz(dic.count - 1, UBound(Aux) + 1)
                    ElseIf (UBound(Aux) + 1) > UBound(mtz, 2) Then
                        ReDim Preserve mtz(dic.count - 1, UBound(Aux) + 1)
                    End If

                    For cntCol = LBound(Aux) To UBound(Aux) Step 1
                        mtz(cnt, cntCol + 1) = Aux(cntCol)
                    Next cntCol
                Else
                    ReDim mtz(dic.count - 1, 0)
                End If
                mtz(cnt, 0) = key
                cnt = cnt + 1
            Next key
        End If
    End If
Fim:
    If VBA.IsArray(mtz) Then GetArrayFromDictionary = mtz
On Error GoTo 0
Exit Function
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxArray.GetArrayFromDictionary", Erl)
    GoTo Fim
End Function

'---------------------------------------------------------------------------------------
' Rotina....: ResizeArray() / Sub
' Contato...: fernando.fernandes@outlook.com.br
' Autor.....: Jefferson Dantas
' Revisão...: Fernando Fernandes
' Empresa...: Planilhando
' Descrição.: This routine resizes any bidimensional array, to a new number of rows, keeping the contents
'             Redim Preserve
'---------------------------------------------------------------------------------------
Public Sub ResizeArray(ByRef mtz As Variant, ByVal NewSize As Long)
On Error GoTo TreatError

Dim FirstElementRow As Long, LastElementRow As Long
Dim FirstElementCol As Long, LastElementCol As Long

    FirstElementRow = LBound(mtz, 1): FirstElementCol = LBound(mtz, 2)

    LastElementRow = UBound(mtz, 1):  LastElementCol = UBound(mtz, 2)

    mtz = Transpose(mtz)
    ReDim Preserve mtz(FirstElementCol To LastElementCol, FirstElementRow To NewSize)
    mtz = Transpose(mtz)

On Error GoTo 0
Exit Sub
TreatError:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "auxArray.ResizeArray()", Erl, True)
End Sub
'
''---------------------------------------------------------------------------------------
'' Rotina....: Transpose() / Function
'' Contato...: fernando.fernandes@outlook.com.br
'' Autor.....: Jefferson Dantas
'' Revisão...: Fernando Fernandes
'' Empresa...: Planilhando
'' Descrição.: This routine transposes any bidimensional array
''---------------------------------------------------------------------------------------
'Public Function Transpose(ByVal Matriz As Variant) As Variant
'On Error GoTo TreatError
'Dim lngContador     As Long
'Dim lngContador1    As Long
'Dim arrAux          As Variant
'
'    If VBA.IsArray(Matriz) Then
'        Select Case NumberOfDimensions(Matriz)
'            Case 1
'                Matriz = Matriz
'
'            Case 2
''creating auxiliary array with inverted dimensions
'                ReDim arrAux(LBound(Matriz, 2) To UBound(Matriz, 2), LBound(Matriz, 1) To UBound(Matriz, 1))
'
'                For lngContador = LBound(Matriz, 2) To UBound(Matriz, 2) Step 1
'                    For lngContador1 = LBound(Matriz, 1) To UBound(Matriz, 1) Step 1
'                        arrAux(lngContador, lngContador1) = Matriz(lngContador1, lngContador)
'                    Next lngContador1
'                Next lngContador
'        End Select
'    End If
'
'    Transpose = arrAux
'On Error GoTo 0
'Exit Function
'TreatError:
'    Call xlExceptions.TreatError(VBA.Err.description, VBA.Err.Number, "auxArray.Transpose()", Erl, True)
'End Function
'

''---------------------------------------------------------------------------------------
'' Rotina....: TrimArray() / Sub
'' Contato...: fernando.fernandes@outlook.com.br
'' Autor.....: Fernando Fernandes
'' Empresa...: www.Planilhando.Com.Br
'' Descrição.: Routine that trims all data in an array
''---------------------------------------------------------------------------------------
'Public Sub TrimArray(ByRef ArrayToTrim As Variant, TrimType As TypeOfTrim)
'On Error GoTo TreatError
'Dim lRow    As Long
'Dim lCol    As Long
'
'    If VBA.IsArray(ArrayToTrim) Then
'        Select Case TrimType
'            Case TypeOfTrim.VisualBasicForApplication
'                For lRow = LBound(ArrayToTrim, 1) To UBound(ArrayToTrim, 1)
'                    For lCol = LBound(ArrayToTrim, 2) To UBound(ArrayToTrim, 2)
'                        If Not VBA.IsNumeric(ArrayToTrim(lRow, lCol)) And Not VBA.IsDate(ArrayToTrim(lRow, lCol)) Then
'                            ArrayToTrim(lRow, lCol) = VBA.Trim(ArrayToTrim(lRow, lCol))
'                        End If
'                    Next lCol
'                Next lRow
'
'            Case TypeOfTrim.WorksheetFunction
'                Call StartWSF
'                For lRow = LBound(ArrayToTrim, 1) To UBound(ArrayToTrim, 1)
'                    For lCol = LBound(ArrayToTrim, 2) To UBound(ArrayToTrim, 2)
'                        If Not VBA.IsNumeric(ArrayToTrim(lRow, lCol)) And Not VBA.IsDate(ArrayToTrim(lRow, lCol)) Then
'                            ArrayToTrim(lRow, lCol) = WSF.Trim(ArrayToTrim(lRow, lCol))
'                        End If
'                    Next lCol
'                Next lRow
'
'            Case TypeOfTrim.Caractere160
'                For lRow = LBound(ArrayToTrim, 1) To UBound(ArrayToTrim, 1)
'                    For lCol = LBound(ArrayToTrim, 2) To UBound(ArrayToTrim, 2)
'                        If Not VBA.IsNumeric(ArrayToTrim(lRow, lCol)) And Not VBA.IsDate(ArrayToTrim(lRow, lCol)) Then
'                            ArrayToTrim(lRow, lCol) = VBA.Replace(ArrayToTrim(lRow, lCol), VBA.Chr(160),VBA.vbNullString)
'                        End If
'                    Next lCol
'                Next lRow
'
'        End Select
'    End If
'
'On Error GoTo 0
'Exit Sub
'TreatError:
'    Call xlExceptions.TreatError(VBA.Err.description, VBA.Err.Number, "auxArray.TrimArray()", Erl, True)
'End Sub

''---------------------------------------------------------------------------------------
'' Rotina....: TextArrayFields() / Sub
'' Contato...: fernando.fernandes@outlook.com.br
'' Autor.....: Fernando Fernandes
'' Empresa...: www.Planilhando.Com.Br
'' Date......: 04/17/2014 (mdy)
'' Descrição.: Forces given or all columns in an array to be text (add apostrohpy
''             when value starts with =, + or -
''---------------------------------------------------------------------------------------
'Public Sub TextArrayFields(ByRef ArrayToText As Variant, ParamArray Fields() As Variant)
'On Error GoTo TreatError
'Dim lRow        As Long
'Dim lCol        As Long
'Dim Col2Text    As Long
'
'    Call StartWSF
'    If VBA.IsArray(ArrayToText) Then
'
'        If Fields(0) = "All" Then
'
'            For lCol = LBound(ArrayToText, 2) To UBound(ArrayToText, 2)
'
'                For lRow = LBound(ArrayToText, 1) To UBound(ArrayToText, 1)
'
'                    If VBA.Left(ArrayToText(lRow, lCol), 1) = "=" Or VBA.Left(ArrayToText(lRow, lCol), 1) = "-" Or VBA.Left(ArrayToText(lRow, lCol), 1) = "+" Then
'                        ArrayToText(lRow, lCol) = "'" & xlWSF.Trim(ArrayToText(lRow, lCol))
'                    End If
'
'                Next lRow
'
'            Next lCol
'
'        Else
'
'            For lCol = LBound(Fields, 1) To UBound(Fields, 1)
'
'                Col2Text = Fields(lCol)
'                For lRow = LBound(ArrayToText, 1) To UBound(ArrayToText, 1)
'
'                    If VBA.Left(ArrayToText(lRow, Col2Text), 1) = "=" Or VBA.Left(ArrayToText(lRow, Col2Text), 1) = "-" Or VBA.Left(ArrayToText(lRow, Col2Text), 1) = "+" Then
'                        ArrayToText(lRow, Col2Text) = "'" & xlWSF.Trim(ArrayToText(lRow, Col2Text))
'                    End If
'
'                Next lRow
'
'            Next lCol
'
'        End If
'
'    End If
'
'On Error GoTo 0
'Exit Sub
'TreatError:
'    Call xlExceptions.TreatError(VBA.Err.description, VBA.Err.Number, "auxArray.TrimArray()", Erl, True)
'End Sub
'
'
'---------------------------------------------------------------------------------------
' Rotina....: GetColFromArray() / Sub
' Contato...: Fernando.Fernandes@Outlook.com.br
' Autor.....: Fernando Fernandes
' Empresa...: www.Planilhando.Com.Br
' Descrição.: Brings a given column from an array, returns the array of that column
'---------------------------------------------------------------------------------------
Public Function GetColFromArray(ByVal arr As Variant, _
                                ByVal ColNumber) As Variant
On Error GoTo TreatError
Dim lRows               As Long
Dim arrAux              As Variant

    If VBA.IsArray(arr) Then
        ReDim arrAux(LBound(arr, 1) To UBound(arr, 1), 1 To 1)

        For lRows = LBound(arrAux, 1) To UBound(arrAux, 1) Step 1
            arrAux(lRows, 1) = arr(lRows, ColNumber)
        Next

    End If

    GetColFromArray = arrAux

On Error GoTo 0
Exit Function
TreatError:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "auxArray.GetColFromArray()", Erl, True)
End Function
'
'---------------------------------------------------------------------------------------
' Rotina....: AppendArrays() / Function
' Contato...: Fernando.Fernandes@Outlook.com.br
' Autor.....: Fernando Fernandes
' Ad........: www.Planilhando.Com.Br
' Date......: 02/18/2014
' Descr.....: Routine that will append array 1 under array one. Both have to be the same dimensions
'---------------------------------------------------------------------------------------
Public Function AppendArrays(ByRef ArrayUp As Variant, _
                             ByRef ArrayDown As Variant, _
                             Optional ByVal Direction As Orientation = Orientation.Vertical)
On Error GoTo TreatError
Dim lRowStart   As Long
Dim lRowEnd     As Long
Dim lRow        As Long
Dim lColStart   As Long
Dim lColEnd     As Long
Dim lCol        As Long
Dim arrAux      As Variant
Dim ArrayLeft   As Variant
Dim ArrayRight  As Variant

    If VBA.IsArray(ArrayUp) And VBA.IsArray(ArrayDown) Then

        Select Case Direction
            Case Orientation.Vertical
                If UBound(ArrayUp, 2) = UBound(ArrayDown, 2) Then
                    lRowStart = UBound(ArrayUp, 1) + 1
                    lRowEnd = UBound(ArrayUp, 1) + UBound(ArrayDown, 1)
                    arrAux = ArrayUp

                    Call ResizeArray(arrAux, lRowEnd)

                    For lRow = lRowStart To lRowEnd
                        For lCol = LBound(ArrayUp, 2) To UBound(ArrayUp, 2)
                            arrAux(lRow, lCol) = ArrayDown(lRow - UBound(ArrayUp, 1), lCol)
                        Next lCol
                    Next lRow
                End If

            Case Orientation.Horizontal
                ArrayLeft = ArrayUp
                ArrayRight = ArrayDown

                If UBound(ArrayLeft, 1) = UBound(ArrayRight, 1) Then

                    lColStart = UBound(ArrayLeft, 2) + 1
                    lColEnd = UBound(ArrayLeft, 2) + UBound(ArrayRight, 2)
                    arrAux = ArrayLeft

                    ReDim Preserve arrAux(LBound(ArrayUp, 1) To UBound(ArrayUp, 1), 1 To lColEnd)

                    For lCol = lColStart To lColEnd
                        For lRow = LBound(ArrayLeft, 1) To UBound(ArrayLeft, 1)
                            arrAux(lRow, lCol) = ArrayRight(lRow, lCol - UBound(ArrayLeft, 2))
                        Next lRow
                    Next lCol

                End If

        End Select
        If VBA.IsArray(arrAux) Then AppendArrays = arrAux
    End If

On Error GoTo 0
Exit Function
TreatError:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "auxArray.TrimArray()", Erl, True)
End Function
'
'---------------------------------------------------------------------------------------
' Rotina....: SumArrays() / Sub
' Contato...: Fernando.Fernandes@Outlook.com.br
' Autor.....: Fernando Fernandes
' Ad........: www.Planilhando.Com.Br
' Date......: 06/24/2014
' Descr.....: Routine that will sum array 1 with array 2. Both have to be the same dimensions
'---------------------------------------------------------------------------------------
Public Function SumArrays(ByVal arr1 As Variant, ByVal arr2 As Variant, _
                          Optional ByVal Orientation As Orientation = Vertical) As Variant
On Error GoTo TreatError
Dim lRow        As Long
Dim lCol        As Long

    If VBA.IsArray(arr1) And VBA.IsArray(arr2) Then

        If UBound(arr1, 1) = UBound(arr2, 1) And UBound(arr1, 2) = UBound(arr2, 2) Then

            Select Case Orientation
                Case Horizontal
                    For lCol = LBound(arr1, 2) To UBound(arr1, 2)
                        arr1(1, lCol) = arr1(1, lCol) + arr2(1, lCol)
                    Next lCol

                Case Vertical
                    For lRow = LBound(arr1, 1) To UBound(arr1, 1)
                        arr1(lRow, 1) = VBA.IIf(Not VBA.IsNumeric(arr1(lRow, 1)), 0, arr1(lRow, 1)) + VBA.IIf(Not VBA.IsNumeric(arr2(lRow, 1)), 0, arr2(lRow, 1))
                    Next lRow

            End Select

            SumArrays = arr1

        End If
    End If

On Error GoTo 0
Exit Function
TreatError:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "auxArray.SumArrays()", Erl, True)
End Function
'
''---------------------------------------------------------------------------------------
'' Rotina....: UpdateArray() / Sub
'' Contato...: Fernando.Fernandes@Outlook.com.br
'' Autor.....: Fernando Fernandes
'' Ad........: www.Planilhando.Com.Br
'' Date......: 07/15/2014
'' Descr.....:
''---------------------------------------------------------------------------------------
'Public Sub UpdateArray(ByRef arrFull As Variant, _
'                       ByRef arrUnidimensional As Variant, _
'                       ByRef Orientation As String, _
'                       ByRef Index As Long)
'On Error GoTo TreatError
'Dim cnt As Long
'
'    Select Case Orientation
'        Case "V"
'            For cnt = LBound(arrFull, 1) To UBound(arrFull) Step 1
'                If Not arrFull(cnt, Index) = arrUnidimensional(cnt, 1) Then arrFull(cnt, Index) = arrUnidimensional(cnt, 1)
'            Next
'        Case "H"
'            For cnt = LBound(arrFull, 1) To UBound(arrFull) Step 1
'                If Not arrFull(Index, cnt) = arrUnidimensional(1, cnt) Then arrFull(Index, cnt) = arrUnidimensional(1, cnt)
'            Next
'
'    End Select
'
'On Error GoTo 0
'Exit Sub
'TreatError:
'    Call xlExceptions.TreatError(VBA.Err.description, VBA.Err.Number, "auxArray.UpdateArray()", Erl, True)
'End Sub
'

'---------------------------------------------------------------------------------------
' Rotina....: CopyArrayRow() / Sub
' Contato...: Fernando.Fernandes@Outlook.com.br
' Autor.....: Fernando Fernandes
' Ad........: www.Planilhando.Com.Br
' Date......: 07/28/2014
' Descr.....: Creates a replica of a row from one array in a row in another array
'---------------------------------------------------------------------------------------
Public Function CopyArrayRow(ByRef arrFrom As Variant, _
                             ByVal rowFrom As Long, _
                             ByRef arrTo As Variant, _
                             ByRef rowTo As Long) As Variant
On Error GoTo TreatError
Dim cnt     As Long
Dim cntFrom As Long
Dim cntTo   As Long

    Select Case NumberOfDimensions(arrFrom)
        Case 1
            If UBound(arrFrom, 1) - LBound(arrFrom, 1) = UBound(arrTo, 2) - LBound(arrTo, 2) Then
                cntFrom = LBound(arrFrom, 1)
                cntTo = LBound(arrTo, 2)

                For cnt = LBound(arrFrom, 1) To UBound(arrFrom, 1) Step 1

                    arrTo(rowTo, cntTo) = arrFrom(cntFrom)

                    cntFrom = cntFrom + 1
                    cntTo = cntTo + 1

                Next cnt
            End If

        Case 2

            If UBound(arrFrom, 2) - LBound(arrFrom, 2) = UBound(arrTo, 2) - LBound(arrTo, 2) Then
                cntFrom = LBound(arrFrom, 2)
                cntTo = LBound(arrTo, 2)

                For cnt = LBound(arrFrom, 2) To UBound(arrFrom, 2) Step 1

                    arrTo(rowTo, cntTo) = arrFrom(rowFrom, cntFrom)

                    cntFrom = cntFrom + 1
                    cntTo = cntTo + 1

                Next cnt
            End If

    End Select
    CopyArrayRow = arrTo

On Error GoTo 0
Exit Function
TreatError:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "auxArray.CopyArrayRow()", Erl, True)
End Function
'

''''---------------------------------------------------------------------------------------
'''' Rotina....: CopyArrayRow() / Sub
'''' Contato...: Fernando.Fernandes@Outlook.com.br
'''' Autor.....: Fernando Fernandes
'''' Ad........: www.Planilhando.Com.Br
'''' Date......: 07/28/2014
'''' Descr.....: Creates a replica of a row from one array in a row in another array
''''---------------------------------------------------------------------------------------
'''Public Sub CopyArrayRow(ByRef arrFrom As Variant, _
'''                        ByVal rowFrom As Long, _
'''                        ByRef arrTo As Variant, _
'''                        ByRef rowTo As Long)
'''On Error GoTo TreatError
'''Dim cnt     As Long
'''Dim cntFrom As Long
'''Dim cntTo   As Long
'''
'''    Select Case NumberOfDimensions(arrFrom)
'''        Case 1
'''            If UBound(arrFrom, 1) - LBound(arrFrom, 1) = UBound(arrTo, 2) - LBound(arrTo, 2) Then
'''                cntFrom = LBound(arrFrom, 1)
'''                cntTo = LBound(arrTo, 2)
'''
'''                For cnt = LBound(arrFrom, 1) To UBound(arrFrom, 1) Step 1
'''
'''                    arrTo(rowTo, cntTo) = arrFrom(cntFrom)
'''
'''                    cntFrom = cntFrom + 1
'''                    cntTo = cntTo + 1
'''
'''                Next cnt
'''            End If
'''        Case 2
'''
'''            If UBound(arrFrom, 2) - LBound(arrFrom, 2) = UBound(arrTo, 2) - LBound(arrTo, 2) Then
'''                cntFrom = LBound(arrFrom, 1)
'''                cntTo = LBound(arrTo, 2)
'''
'''                For cnt = LBound(arrFrom, 2) To UBound(arrFrom, 2) Step 1
'''
'''                    arrTo(rowTo, cntTo) = arrFrom(rowFrom, cntFrom)
'''
'''                    cntFrom = cntFrom + 1
'''                    cntTo = cntTo + 1
'''
'''                Next cnt
'''            End If
'''    End Select
'''
'''On Error GoTo 0
'''Exit Sub
'''TreatError:
'''    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "auxArray.CopyArrayRow()", Erl, True)
'''End Sub
'
''---------------------------------------------------------------------------------------
'' Rotina....: CopyArrayCol() / Function
'' Contato...: Fernando.Fernandes@Outlook.com.br
'' Autor.....: Fernando Fernandes
'' Ad........: www.Planilhando.Com.Br
'' Date......: 08/14/2014
'' Descr.....: Creates a replica of a column from one array in a column in another array
''---------------------------------------------------------------------------------------
'Public Function CopyArrayCol(ByRef arrFrom As Variant, _
'                             ByVal colFrom As Long, _
'                             ByRef arrTo As Variant, _
'                             ByRef colTo As Long) As Variant
'On Error GoTo TreatError
'Dim cnt As Long
'Dim cntFrom As Long
'Dim cntTo   As Long
'
'    If UBound(arrFrom, 1) - LBound(arrFrom, 1) = UBound(arrTo, 1) - LBound(arrTo, 1) Then
'        Select Case NumberOfDimensions(arrFrom)
'            Case 2
'                For cnt = LBound(arrFrom, 1) To UBound(arrFrom, 1)
'                    arrTo(cnt, colTo) = arrFrom(cnt, colFrom)
'                Next cnt
'
'            Case 1
'
'                For cntFrom = LBound(arrFrom, 1) To UBound(arrFrom, 1)
'                    cntTo = VBA.vba.iif(cntTo = 0, LBound(arrTo, 1), cntTo + 1)
'                    arrTo(cntTo, colTo) = arrFrom(cntFrom)
'
'                Next cntFrom
'
'        End Select
'        CopyArrayCol = arrTo
'    End If
'
'On Error GoTo 0
'Exit Function
'TreatError:
'    Call xlExceptions.TreatError(VBA.Err.description, VBA.Err.Number, "auxArray.CopyArrayCol()", Erl, True)
'End Function
'
'---------------------------------------------------------------------------------------
' Rotina....: RemoveRowFromArray() / Function
' Contato...: Fernando.Fernandes@Outlook.com.br
' Autor.....: Fernando Fernandes
' Ad........: www.Planilhando.Com.Br
' Date......: 08/14/2014
' Descr.....:
'---------------------------------------------------------------------------------------
Public Function RemoveRowFromArray(ByRef arr As Variant, _
                                   ByVal row As Long) As Variant
On Error GoTo TreatError
Dim CutOffArray As Variant
Dim cntSource   As Long
Dim cntDestiny  As Long
Dim lCol        As Long

    If VBA.IsArray(arr) Then

        ReDim CutOffArray(LBound(arr, 1) To UBound(arr, 1), LBound(arr, 2) To UBound(arr, 2))
        cntDestiny = LBound(arr, 1)

        For cntSource = LBound(arr, 1) To UBound(arr, 1)
            If cntSource <> row Then
                For lCol = LBound(arr, 2) To UBound(arr, 2)

                    CutOffArray(cntDestiny, lCol) = arr(cntSource, lCol)

                Next lCol
                cntDestiny = cntDestiny + 1
            End If
        Next cntSource
        Call ResizeArray(CutOffArray, cntDestiny - 1)
        RemoveRowFromArray = CutOffArray

    Else
        RemoveRowFromArray = arr
    End If

    Call RemoverObjetosMemoria(CutOffArray, cntSource, cntDestiny, lCol, arr)

On Error GoTo 0
Exit Function
TreatError:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "auxArray.RemoveRowFromArray()", Erl, True)

End Function
'
'---------------------------------------------------------------------------------------
' Rotina....: InsertColumnInArray() / Function
' Contato...: Fernando.Fernandes@Outlook.com.br
' Autor.....: Fernando Fernandes
' Ad........: www.Planilhando.Com.Br
' Date......: 08/14/2014
' Descr.....: Inserts a column in an array, with a title and content if informed
'---------------------------------------------------------------------------------------
Public Function InsertColumnInArray(ByRef arr As Variant, _
                                    Optional ByVal position As position = position.AfterLast, _
                                    Optional ByVal Content As String = VBA.vbNullString, _
                                    Optional ByVal Title As String) As Variant
On Error GoTo TreatError
Dim arrColumn   As Variant
Dim arrAux      As Variant
Dim cnt         As Long
Dim Location    As Long

    ReDim arrColumn(LBound(arr, 1) To UBound(arr, 1), LBound(arr, 2) To LBound(arr, 2))

    If position = BeforeFirst Then
        arrAux = AppendArrays(arrColumn, arr, Orientation.Horizontal)
        Location = 1
    ElseIf position = AfterLast Then
        arrAux = AppendArrays(arr, arrColumn, Orientation.Horizontal)
        Location = UBound(arrAux, 2)
    End If

    If Content <> VBA.vbNullString Then
        For cnt = LBound(arrAux, 1) To UBound(arrAux, 1)
            arrAux(cnt, Location) = Content
        Next cnt
    End If
    If Title <> VBA.vbNullString Then arrAux(LBound(arrAux, 1), Location) = Title

    InsertColumnInArray = arrAux

    Call RemoverObjetosMemoria(arrColumn, arrAux)

On Error GoTo 0
Exit Function
TreatError:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "auxArray.InsertColumnInArray()", Erl, True)
End Function

'---------------------------------------------------------------------------------------
' Modulo....: auxArray / Módulo
' Rotina....: AjustarArray / Sub
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 13/01/2014
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina que corrige o array inserindo o valor da linha anterior na linha
'             vazia, isso acontece nos casos em que a planilha contem celulas de linhas
'             mescladas.
'---------------------------------------------------------------------------------------
Public Function AjustarArray(ByRef arrDados As Variant) As Variant
On Error GoTo TratarErro
Dim ContLinha       As Long
Dim ContCol         As Integer
Dim AuxDado         As Variant

    For ContCol = LBound(arrDados, 2) To UBound(arrDados, 2) Step 1
        For ContLinha = LBound(arrDados, 1) To UBound(arrDados, 1) Step 1
            If VBA.IsEmpty(arrDados(ContLinha, ContCol)) Then
                arrDados(ContLinha, ContCol) = AuxDado
            Else
                AuxDado = arrDados(ContLinha, ContCol)
            End If
        Next ContLinha
    Next ContCol

    AjustarArray = arrDados

On Error GoTo 0
Exit Function
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "auxArray.AjustarArray()", Erl)
End Function

Public Function ArrayBidimensional(ByVal arrParametros As Variant) As Boolean
On Error Resume Next 'Resume next necessario, pois o erro pode acontecer
Dim blResultado     As Boolean
Dim Var             As Variant
    Var = arrParametros(0)(0, 0)
    If VBA.Err.Number = 0 Then
        blResultado = True
    Else
        blResultado = False
    End If
    ArrayBidimensional = blResultado
End Function

'---------------------------------------------------------------------------------------
' Modulo    : ConexaoDB / Módulo de classe
' Rotina    : ContemMaisUbounds() / Function
' Autor     : Jefferson
' Data      : 07/11/2012 - 15:54
' Proposta  : Função para verificar se o arrParametros contém mais de uma dimensao (ubound)
'---------------------------------------------------------------------------------------
Public Function ContemMaisUbounds(ByVal arrParametros As Variant) As Boolean
On Error Resume Next 'Resume next necessario, pois o erro pode acontecer
Dim blResultado     As Boolean

    If UBound(arrParametros(0)) >= 0 Then
        If VBA.Err.Number = 0 Then blResultado = True
    Else
        blResultado = False
    End If
    ContemMaisUbounds = blResultado
End Function
'
'---------------------------------------------------------------------------------------
' Modulo    : AuxArray / Módulo
' Rotina    : AcertarArrayParamentros() / Function
'Autor:       Jefferson
' Data      : 07/11/2012 - 15:55
' Proposta  : Função para acertar o array caso ele tenha mais de uma dimensao
'---------------------------------------------------------------------------------------
Public Function Acertar_Array_Parametros(ByVal arrParametros As Variant) As Variant
On Error GoTo TratarErro
Dim arrAux          As Variant
Dim btContador      As Byte

    If ContemMaisUbounds(arrParametros) Then
        If ArrayBidimensional(arrParametros) Then
            ReDim arrAux(UBound(arrParametros(0), 2))
            For btContador = 0 To UBound(arrParametros(0), 2) Step 1
                arrAux(btContador) = arrParametros(0)(1, btContador)
            Next btContador
        Else
            ReDim arrAux(UBound(arrParametros(0)))
            For btContador = 0 To UBound(arrParametros(0)) Step 1
                arrAux(btContador) = arrParametros(0)(btContador)
            Next btContador
        End If
    Else
        arrAux = arrParametros
    End If
    Acertar_Array_Parametros = arrAux

Exit Function
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxArray.Acertar_Array_Parametros()", Erl)
Resume
End Function

'---------------------------------------------------------------------------------------
' Modulo    : AuxArray / Módulo
' Rotina    : CriarMatriz() / Function
' Autor     : Fernando Fernandes
' Data      : 14/04*2015 - 15:55
' Proposta  : Função para criar uma matriz, dados o índice do primeiro elemento (0 ou 1)
'             e o tamanho, com o índice dos últimos respectivos elementos
'---------------------------------------------------------------------------------------
Public Function CriarMatrizBidimensional(ByVal PrimeiroIndice As Byte, _
                                         ByVal NumLinhas As Long, _
                                         ByVal NumColunas As Long) As Variant
On Error GoTo TratarErro
Dim AuxMtz  As Variant

    ReDim AuxMtz(PrimeiroIndice To NumLinhas, PrimeiroIndice To NumColunas)
    CriarMatrizBidimensional = AuxMtz

On Error GoTo 0
Exit Function
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxArray.CriarMatrizBidimensional()", Erl)
    Exit Function
Resume Next
End Function


'---------------------------------------------------------------------------------------
' Rotina....: JuntarMatrizes() / Function
' Contato...: Fernando.Fernandes@Outlook.com.br
' Autor.....: Fernando Fernandes
' Ad........: www.Planilhando.Com.Br
' Date......: 02/18/2014
' Descr.....: Routine that will append array 1 under array one. Both have to be the same dimensions
'---------------------------------------------------------------------------------------
Public Function JuntarMatrizes(ByRef ArrayUp As Variant, _
                               ByRef ArrayDown As Variant, _
                               Optional ByVal Direction As Orientation = Orientation.Vertical)
On Error GoTo TreatError
Dim lRowStart   As Long
Dim lRowEnd     As Long
Dim lRow        As Long
Dim lColStart   As Long
Dim lColEnd     As Long
Dim lCol        As Long
Dim arrAux      As Variant
Dim ArrayLeft   As Variant
Dim ArrayRight  As Variant

    If VBA.IsArray(ArrayUp) And VBA.IsArray(ArrayDown) Then

        Select Case Direction
            Case Orientation.Vertical
                If UBound(ArrayUp, 2) = UBound(ArrayDown, 2) Then
                    lRowStart = UBound(ArrayUp, 1) + 1
                    lRowEnd = UBound(ArrayUp, 1) + UBound(ArrayDown, 1)
                    arrAux = ArrayUp

                    Call ResizeArray(arrAux, lRowEnd)

                    For lRow = lRowStart To lRowEnd
                        For lCol = LBound(ArrayUp, 2) To UBound(ArrayUp, 2)
                            arrAux(lRow, lCol) = ArrayDown(lRow - UBound(ArrayUp, 1), lCol)
                        Next lCol
                    Next lRow
                End If

            Case Orientation.Horizontal
                ArrayLeft = ArrayUp
                ArrayRight = ArrayDown

                If UBound(ArrayLeft, 1) = UBound(ArrayRight, 1) Then

                    lColStart = UBound(ArrayLeft, 2) + 1
                    lColEnd = UBound(ArrayLeft, 2) + UBound(ArrayRight, 2)
                    arrAux = ArrayLeft

                    ReDim Preserve arrAux(LBound(ArrayUp, 1) To UBound(ArrayUp, 1), LBound(ArrayUp, 2) To lColEnd)

                    For lCol = lColStart To lColEnd
                        For lRow = LBound(ArrayLeft, 1) To UBound(ArrayLeft, 1)
                            arrAux(lRow, lCol) = ArrayRight(lRow, lCol - UBound(ArrayLeft, 2))
                        Next lRow
                    Next lCol

                End If

        End Select
        If VBA.IsArray(arrAux) Then JuntarMatrizes = arrAux
    End If

On Error GoTo 0
Exit Function
TreatError:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "auxArray.JuntarMatrizes()", Erl, True)
End Function


'---------------------------------------------------------------------------------------
' Rotina....: TransporMatriz() / Function
' Contato...: fernando@tecnun.com.br
' Autor.....: Jefferson Dantas
' Revisão...: Fernando Fernandes
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: This routine transposes any uni or bidimensional array
'---------------------------------------------------------------------------------------
Public Function TransporMatriz(ByVal Matriz As Variant) As Variant
On Error GoTo TreatError
Dim lngContador     As Long
Dim lngContador1    As Long
Dim arrAux          As Variant

    If VBA.IsArray(Matriz) Then
        Select Case NumberOfDimensions(Matriz)
            Case 1
                Matriz = Matriz

            Case 2
                'creating auxiliary array with inverted dimensions
                ReDim arrAux(LBound(Matriz, 2) To UBound(Matriz, 2), LBound(Matriz, 1) To UBound(Matriz, 1))

                For lngContador = LBound(Matriz, 2) To UBound(Matriz, 2) Step 1
                    For lngContador1 = LBound(Matriz, 1) To UBound(Matriz, 1) Step 1
                        arrAux(lngContador, lngContador1) = Matriz(lngContador1, lngContador)
                    Next lngContador1
                Next lngContador
        End Select
    End If

    TransporMatriz = arrAux
On Error GoTo 0
Exit Function
TreatError:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "auxArray.TransporMatriz()", Erl)
End Function

'---------------------------------------------------------------------------------------
' Modulo    : AuxArray / Módulo
' Rotina    : PreencherValor() / Function
' Autor     : Fernando Fernandes
' Data      : 28/01/2016
' Proposta  : Função para preencher um campo de uma matriz com um valor específico
'---------------------------------------------------------------------------------------
Public Function PreencherValor(ByRef Matriz As Variant, ByVal campo As Long, ByVal valor As Variant, ByVal Header As XlYesNoGuess) As Variant
On Error GoTo TratarErro
Dim lin As Long
Dim ini As Long

    If VBA.IsArray(Matriz) Then
        ini = LBound(Matriz, 1) + VBA.IIf(Header = xlYes, 1, 0)
        Select Case NumberOfDimensions(Matriz)
            Case 1
                For lin = LBound(Matriz, 1) + 1 To UBound(Matriz, 1) Step 1
                    Matriz(lin) = valor
                Next lin
            Case 2
                For lin = LBound(Matriz, 1) + 1 To UBound(Matriz, 1) Step 1
                    Matriz(lin, campo) = valor
                Next lin
        End Select
    End If
    PreencherValor = Matriz

Fim:
On Error GoTo 0
Exit Function
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxArray.PreencherValor()", Erl)
    Exit Function
Resume Next
End Function

'---------------------------------------------------------------------------------------
' Modulo    : AuxArray / Módulo
' Rotina    : ConverterMatriz() / Function
' Autor     : Fernando Fernandes / Jeff Dantas
' Data      : 20/08/2015
' Proposta  : Função para mudar as dimensões de uma matriz. de 2 pra 1 ou de 1 pra duas.
'---------------------------------------------------------------------------------------
Public Function ConverterMatriz(ByVal ArrayToRearrange As Variant, _
                                Optional ByVal Orientation As Orientation = Vertical) As Variant
On Error GoTo TratarErro
Dim arrAux  As Variant
Dim cnt     As Long

    If VBA.IsArray(ArrayToRearrange) Then
        Select Case NumberOfDimensions(ArrayToRearrange)
            Case 1

                If Orientation = Horizontal Then
                    ReDim arrAux(LBound(ArrayToRearrange, 1) To LBound(ArrayToRearrange, 1), _
                                 LBound(ArrayToRearrange, 1) To UBound(ArrayToRearrange, 1))
                    For cnt = LBound(ArrayToRearrange, 1) To UBound(ArrayToRearrange, 1) Step 1
                        arrAux(LBound(arrAux, 1), cnt) = ArrayToRearrange(cnt)
                    Next cnt
                ElseIf Orientation = Vertical Then
                    ReDim arrAux(LBound(ArrayToRearrange, 1) To UBound(ArrayToRearrange, 1), _
                                 LBound(ArrayToRearrange, 1) To LBound(ArrayToRearrange, 1))
                    For cnt = LBound(ArrayToRearrange, 1) To UBound(ArrayToRearrange, 1) Step 1
                        arrAux(cnt, LBound(arrAux, 1)) = ArrayToRearrange(cnt)
                    Next cnt
                End If

            Case 2

                If LBound(ArrayToRearrange, 1) = UBound(ArrayToRearrange, 1) Then
                    ReDim arrAux(LBound(ArrayToRearrange, 2) To UBound(ArrayToRearrange, 2))
                    Orientation = Horizontal
                    For cnt = LBound(arrAux, 1) To UBound(arrAux, 1) Step 1
                        arrAux(cnt) = ArrayToRearrange(LBound(ArrayToRearrange, 1), cnt)
                    Next cnt
                ElseIf LBound(ArrayToRearrange, 2) = UBound(ArrayToRearrange, 2) Then
                    ReDim arrAux(LBound(ArrayToRearrange, 1) To UBound(ArrayToRearrange, 1))
                    Orientation = Vertical
                    For cnt = LBound(arrAux, 1) To UBound(arrAux, 1) Step 1
                        arrAux(cnt) = ArrayToRearrange(cnt, LBound(ArrayToRearrange, 1))
                    Next cnt
                End If

        End Select
        If VBA.IsArray(arrAux) Then ConverterMatriz = arrAux
    End If

Fim:
On Error GoTo 0

Exit Function
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxArray.RedimensionarMatriz()", Erl)
    Exit Function
Resume Next
End Function

'Função que converte o resultado de uma lista de um de um Recordset para String.
Public Function RsToString(obj, delimitador) As String
    On Error GoTo ErrConvert
    With obj
        .MoveLast
        .MoveFirst
        RsToString = VBA.Join(AuxArray.ConverterMatriz(AuxArray.Transpose(.GetRows(.RecordCount)), Horizontal), delimitador)
    End With
    Exit Function
ErrConvert:
    If VBA.Err <> 0 Then
        RsToString = "#Error > " & VBA.Err.Description
    End If
End Function


'---------------------------------------------------------------------------------------
' Rotina....: JuntarMatrizes2() / Function
' Contato...: Fernando.Fernandes@Outlook.com.br
' Autor.....: Fernando Fernandes
' Ad........: www.Planilhando.Com.Br
' Date......: 02/18/2014
' Descr.....: Routine that will append array 1 under array one. Both have to be the same dimensions
'---------------------------------------------------------------------------------------
Public Function JuntarMatrizes2(ByRef ArrayUp As Variant, _
                                ByRef ArrayDown As Variant, _
                                Optional ByVal Direction As Orientation = Orientation.Vertical)
On Error GoTo TreatError
Dim lRowStart   As Long
Dim lRowEnd     As Long
Dim lRow        As Long
Dim lColStart   As Long
Dim lColEnd     As Long
Dim lCol        As Long
Dim arrAux      As Variant
Dim ArrayLeft   As Variant
Dim ArrayRight  As Variant

Dim qtdUp       As Long
Dim qtdDown     As Long
Dim linOrigem   As Long

    If VBA.IsArray(ArrayUp) And VBA.IsArray(ArrayDown) Then

        Select Case Direction
            Case Orientation.Vertical
                If UBound(ArrayUp, 2) = UBound(ArrayDown, 2) Then
                    qtdUp = UBound(ArrayUp, 1) - LBound(ArrayUp, 1) + 1
                    qtdDown = UBound(ArrayDown, 1) - LBound(ArrayDown, 1) + 1

                    lRowStart = UBound(ArrayUp, 1) + 1
                    lRowEnd = UBound(ArrayUp, 1) + qtdDown

                    arrAux = ArrayUp

                    Call ResizeArray(arrAux, lRowEnd)
                    linOrigem = LBound(ArrayDown, 1)
                    For lRow = lRowStart To lRowEnd
                        For lCol = LBound(ArrayUp, 2) To UBound(ArrayUp, 2)
                            arrAux(lRow, lCol) = ArrayDown(linOrigem, lCol)
                        Next lCol
                        linOrigem = linOrigem + 1
                    Next lRow
                End If

        End Select
        If VBA.IsArray(arrAux) Then JuntarMatrizes2 = arrAux
    End If

On Error GoTo 0
Exit Function
TreatError:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "auxArray.JuntarMatrizes2()", Erl, True)
End Function



