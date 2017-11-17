VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTFW_Array"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' PROJETO/MODULO   : cTFW_Array
' TIPO             : Class Module
' DATA/HORA        : 14/07/2016 12:31
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Encapsula funções de manipulação de Matrizes (Array)
'                    As funções foram construidas pelo Fernando Fernandes
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão do Módulo
' **************************************************************************************
'   Versão    Data/Hora             Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      14/07/2016 12:31
'---------------------------------------------------------------------------------------
Option Compare Database
Private m_vMatriz As Variant
Private m_lRowCount As Long
Private m_lColumnCount As Long
Private m_lNewElement As Long
Private m_lElement As Long
Private m_lRowsCount As Long

Public Event OnError(ErrorMessage As String)

Function Limpar()
    Call NewArray(0)
End Function

'PENDENTE DE IMPLEMENTAÇÃO
Function RemoveElement(lIndex)
'----------------------------------------------------------------------------------------------------
 On Error GoTo ErrorHandler: Const cstr_ProcedureName As String = "cTFW_Array.RemoveElement()"
'----------------------------------------------------------------------------------------------------

    'Remover pelo index se for numerico
    If VBA.IsNumeric(lIndex) Then
    
    Else
        'Se o index for um texto, percorrer o array, identifica o index e remove por index
    End If
    
    Exit Function
ErrorHandler: If VBA.Err <> 0 Then: Call ColetaErro(VBA.Err.Number, VBA.Err.Description, cstr_ProcedureName)
End Function

Private Sub RemoveElementByIndex(lIndex As Long)
'----------------------------------------------------------------------------------------------------
 On Error GoTo ErrorHandler: Const cstr_ProcedureName As String = "cTFW_Array.RemoveElementByIndex()"
'----------------------------------------------------------------------------------------------------


Exit Sub
ErrorHandler: If VBA.Err <> 0 Then: Call ColetaErro(VBA.Err.Number, VBA.Err.Description, cstr_ProcedureName)
End Sub

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cTFW_Array.Matriz()
' TIPO             : Property
' DATA/HORA        : 14/07/2016 13:06
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Retorna a instancia do array ja populado
'---------------------------------------------------------------------------------------
Public Property Get Matriz() As Variant
    On Error GoTo Erro
    Matriz = m_vMatriz
    Exit Property
Erro:
    If VBA.Err = 9 Then
        Matriz = VBA.vbEmpty
    End If
End Property

Public Property Let Matriz(ByVal vMatriz As Variant)
    m_vMatriz = vMatriz
    Call Resize(UBound(m_vMatriz, 1))
End Property

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cTFW_Array.Clone()
' TIPO             : Function
' DATA/HORA        : 24/05/2017 18:06
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Cria uma copia do array na instancia atual e retorna uma nova instancia
'---------------------------------------------------------------------------------------
Function Clone() As cTFW_Array
'----------------------------------------------------------------------------------------------------
On Error GoTo ErrorHandler: Const cstr_ProcedureName As String = "cTFW_Array.Clone()"
'----------------------------------------------------------------------------------------------------
    Dim newArr As New cTFW_Array
    Set newArr = New cTFW_Array
    Call newArr.CopyFromArray(Me.Matriz)
    Set Clone = newArr
    
    Exit Function
ErrorHandler: If VBA.Err <> 0 Then: Call ColetaErro(VBA.Err.Number, VBA.Err.Description, cstr_ProcedureName)
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cTFW_Array.CopyFromArray()
' TIPO             : Function
' DATA/HORA        : 14/07/2016 13:06
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Carrega os dados do array atual a partir de uma fonte externa
'---------------------------------------------------------------------------------------
Function CarregarDe(FonteLista, Optional bTranspose As Boolean = False)
'----------------------------------------------------------------------------------------------------
On Error GoTo ErrorHandler: Const cstr_ProcedureName As String = "cTFW_Array.CopyFromArray()"
'----------------------------------------------------------------------------------------------------
    Dim mtz, iCol As Long
    
    If VBA.IsArray(FonteLista) Then
        Call CopyFromArray(FonteLista, bTranspose)
    ElseIf VBA.IsObject(FonteLista) Then
        If VBA.TypeName(FonteLista) = "Collection" Then
            Me.Limpar
            For iCol = 1 To FonteLista.count
               Me.AddElement FonteLista(iCol)
            Next iCol
        ElseIf VBA.TypeName(FonteLista) = "Range" Then
            Call CopyFromArray(FonteLista.value, bTranspose)
        End If
    End If
    CarregarDe = VBA.Err = 0
    Exit Function
ErrorHandler: If VBA.Err <> 0 Then: Call ColetaErro(VBA.Err.Number, VBA.Err.Description, cstr_ProcedureName)
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cTFW_Array.CopyFromArray()
' TIPO             : Function
' DATA/HORA        : 14/07/2016 13:06
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Carrega os dados do array atual a partir de uma fonte externa
'---------------------------------------------------------------------------------------
Public Function CopyFromArray(mtz, Optional bTranspose As Boolean = False)
'----------------------------------------------------------------------------------------------------
On Error GoTo ErrorHandler: Const cstr_ProcedureName As String = "cTFW_Array.CopyFromArray()"
'----------------------------------------------------------------------------------------------------
    Matriz = mtz
    If bTranspose Then Matriz = Transpose(mtz)
    Call Resize(UBound(m_vMatriz, 1))
    ResetDimencoes
    Exit Function
ErrorHandler: If VBA.Err <> 0 Then: Call ColetaErro(VBA.Err.Number, VBA.Err.Description, cstr_ProcedureName)
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cTFW_Array.AddElement()
' TIPO             : Function
' DATA/HORA        : 14/07/2016 13:06
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Adiciona um novo item (elemento) no array
'---------------------------------------------------------------------------------------
Public Function AddElement(Optional ByVal lNewElement)
'----------------------------------------------------------------------------------------------------
On Error GoTo ErrorHandler: Const cstr_ProcedureName As String = "cTFW_Array.AddElement()"
'----------------------------------------------------------------------------------------------------
    Dim lColumnIndex
    If NumberOfDimensions(m_vMatriz) = 2 Then
        m_vMatriz(m_lRowCount - 1, lColumnIndex - 1) = lNewElement
    Else
        If UBound(m_vMatriz, 1) > 0 Or m_vMatriz(m_lRowCount) <> VBA.vbEmpty Then
            Call Resize(m_lRowCount + 1)
            m_vMatriz(m_lRowCount) = lNewElement
        Else
            m_vMatriz(m_lRowCount) = lNewElement
        End If
    End If
    Call ResetDimencoes
    m_lRowsCount = m_lRowsCount + 1
    Exit Function
ErrorHandler: If VBA.Err <> 0 Then: Call ColetaErro(VBA.Err.Number, VBA.Err.Description, cstr_ProcedureName)
End Function
'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : ctfw_Array.Element()
' TIPO             : Property
' DATA/HORA        : 14/07/2016 13:06
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Recupera um item de um array na coordenada linha/coluna especificada.
'                    Caso não informa, recupera o ultimo item da  matriz
'---------------------------------------------------------------------------------------
Public Property Get element(Optional iRowIndex As Long = -1, Optional lColumnIndex = -1) As Variant
'----------------------------------------------------------------------------------------------------
 On Error GoTo ErrorHandler: Const cstr_ProcedureName As String = "cTFW_Array.Get Element()"
'----------------------------------------------------------------------------------------------------
    
    If iRowIndex = -1 Then iRowIndex = RowCount
    If lColumnIndex = -1 Then lColumnIndex = ColumnCount
    
    If NumberOfDimensions(m_vMatriz) = 2 Then
        element = m_vMatriz(iRowIndex, lColumnIndex)
    Else
        element = m_vMatriz(iRowIndex)
    End If

Fim:
 On Error GoTo 0
 Exit Property

ErrorHandler: If VBA.Err <> 0 Then: Call ColetaErro(VBA.Err.Number, VBA.Err.Description, cstr_ProcedureName)
End Property

Public Property Let element(iRowIndex As Long, lColumnIndex, ByVal lElement As Variant)
'----------------------------------------------------------------------------------------------------
 On Error GoTo ErrorHandler: Const cstr_ProcedureName As String = "cTFW_Array.Let Element()"
'----------------------------------------------------------------------------------------------------

    If NumberOfDimensions(m_vMatriz) = 2 Then
        m_vMatriz(iRowIndex, lColumnIndex) = lElement
    Else
        m_vMatriz(iRowIndex) = lElement
    End If
    Call ResetDimencoes
    
ErrorHandler: If VBA.Err <> 0 Then: Call ColetaErro(VBA.Err.Number, VBA.Err.Description, cstr_ProcedureName)
End Property

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cTFW_Array.ResetDimencoes()
' TIPO             : Function
' DATA/HORA        : 14/07/2016 13:06
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Redefine os valores dos tamanhos das dimensoes Largura e ALtura
'---------------------------------------------------------------------------------------
Private Sub ResetDimencoes()
'----------------------------------------------------------------------------------------------------
 On Error GoTo ErrorHandler: Const cstr_ProcedureName As String = "cTFW_Array.ResetDimencoes()"
'----------------------------------------------------------------------------------------------------
    
    m_lRowCount = UBound(m_vMatriz, 1)
    If NumberOfDimensions(m_vMatriz) = 2 Then m_lColumnCount = UBound(m_vMatriz, 2)
    m_lRowsCount = m_lRowCount
ErrorHandler: If VBA.Err <> 0 Then: Call ColetaErro(VBA.Err.Number, VBA.Err.Description, cstr_ProcedureName)
End Sub

'Inicializa a classe com uma matriz Bidimencional
Private Sub Class_Initialize()
    Call NewArray(0)
End Sub
'---------------------------------------------------------------------------------------
' Rotina....: ResizeArray() / Sub
' Contato...: fernando.fernandes@outlook.com.br
' Autor.....: Jefferson Dantas
' Revisão...: Fernando Fernandes
' Empresa...: Planilhando
' Descrição.: This routine resizes any bidimensional array, to a new number of rows, keeping the contents
'             Redim Preserve
'---------------------------------------------------------------------------------------
Public Sub Resize(ByVal NewSize As Long)
'----------------------------------------------------------------------------------------------------
 On Error GoTo ErrorHandler: Const cstr_ProcedureName As String = "cTFW_Array.Resize()"
'----------------------------------------------------------------------------------------------------

Dim mtz As Variant
Dim FirstElementRow As Long, LastElementRow As Long
Dim FirstElementCol As Long, LastElementCol As Long

    mtz = m_vMatriz
    FirstElementRow = LBound(mtz, 1): If NumberOfDimensions(mtz) = 2 Then FirstElementCol = LBound(mtz, 2)
    LastElementRow = UBound(mtz, 1):  If NumberOfDimensions(mtz) = 2 Then LastElementCol = UBound(mtz, 2)
    mtz = Transpose(mtz)
    
    If NumberOfDimensions(mtz) = 2 Then
        ReDim Preserve mtz(FirstElementCol To LastElementCol, FirstElementRow To NewSize)
    Else
        ReDim Preserve mtz(FirstElementRow To NewSize)
    End If
    mtz = Transpose(mtz)
    m_vMatriz = mtz
    Call ResetDimencoes
On Error GoTo 0
Exit Sub

ErrorHandler: If VBA.Err <> 0 Then: Call ColetaErro(VBA.Err.Number, VBA.Err.Description, cstr_ProcedureName)
End Sub

'---------------------------------------------------------------------------------------
' Modulo    : AuxArray / Módulo
' Rotina    : NewArray() / Function
' Autor     : Fernando Fernandes
' Data      : 14/04*2015 - 15:55
' Proposta  : Função para criar uma matriz, dados o índice do primeiro elemento (0 ou 1)
'             e o tamanho, com o índice dos últimos respectivos elementos
'---------------------------------------------------------------------------------------
Private Function NewArray(Optional ByVal PrimeiroIndice As Byte = 0)
'----------------------------------------------------------------------------------------------------
 On Error GoTo ErrorHandler: Const cstr_ProcedureName As String = "cTFW_Array.NewArray()"
'----------------------------------------------------------------------------------------------------
    
    Dim AuxMtz  As Variant
    If ColumnCount > 0 Then
        ReDim AuxMtz(PrimeiroIndice To RowCount, PrimeiroIndice To ColumnCount)
    Else
        ReDim AuxMtz(PrimeiroIndice To RowCount)
    End If
    m_vMatriz = AuxMtz
    Call ResetDimencoes
On Error GoTo 0
Exit Function
ErrorHandler: If VBA.Err <> 0 Then: Call ColetaErro(VBA.Err.Number, VBA.Err.Description, cstr_ProcedureName)
End Function

'---------------------------------------------------------------------------------------
' Rotina......: NumberOfDimensions() / Sub
' Contato.....: fernando.fernandes@outlook.com.br
' Autor.......: Fernando Fernandes
' Date........: Feb/18th/2020
' Observação..: Eu sei é futuro, mas uma rotina que fala de dimensões, tinha que falar de viagem no tempo
' Empresa.....: www.Planilhando.Com.Br
' Descrição...: Retorna o número total de dimensões de uma matriz
'---------------------------------------------------------------------------------------
Private Function NumberOfDimensions(ByVal arr As Variant) As Long
'----------------------------------------------------------------------------------------------------
 On Error GoTo ErrorHandler: Const cstr_ProcedureName As String = "cTFW_Array.NumberOfDimensions()"
'----------------------------------------------------------------------------------------------------
Dim cnt As Long

    cnt = 1
    Do Until VBA.Err.Number <> 0
        If LBound(arr, cnt) >= 0 Then NumberOfDimensions = cnt
        cnt = cnt + 1
    Loop

On Error GoTo 0
Exit Function
ErrorHandler:
Exit Function
End Function

'---------------------------------------------------------------------------------------
' Rotina....: Transpose() / Function
' Contato...: fernando@tecnun.com.br
' Autor.....: Jefferson Dantas
' Revisão...: Fernando Fernandes
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: This routine transposes any uni or bidimensional array
'---------------------------------------------------------------------------------------
Private Function Transpose(ByVal Matriz As Variant) As Variant
'----------------------------------------------------------------------------------------------------
 On Error GoTo ErrorHandler: Const cstr_ProcedureName As String = "cTFW_Array.Transpose()"
'----------------------------------------------------------------------------------------------------

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
ErrorHandler: If VBA.Err <> 0 Then: Call ColetaErro(VBA.Err.Number, VBA.Err.Description, cstr_ProcedureName)
End Function

Private Property Get RowCount() As Long
    RowCount = m_lRowCount
End Property

Private Property Let RowCount(ByVal lRowCount As Long)
    m_lRowCount = lRowCount
End Property

Private Property Get ColumnCount() As Long
    ColumnCount = m_lColumnCount
End Property

Private Property Let ColumnCount(ByVal lColumnCount As Long)
    m_lColumnCount = lColumnCount
End Property

Public Property Get RowsCount() As Long
    RowsCount = m_lRowsCount
End Property

'Coleta e trata erros ocorridos nessa classe
Function ColetaErro(lngErro As Long, strDescricao As String, Optional strProcedure As String)
    Dim strDescError As String
    strDescError = "Error : " & strProcedure & " > " & strDescricao
    Debug.Print strDescError
    RaiseEvent OnError(strDescError)
End Function

Public Sub QuickSort(vArray As Variant, inLow As Long, inHi As Long)

  Dim pivot   As Variant
  Dim tmpSwap As Variant
  Dim tmpLow  As Long
  Dim tmpHi   As Long

  tmpLow = inLow
  tmpHi = inHi

  pivot = vArray((inLow + inHi) \ 2)

  While (tmpLow <= tmpHi)

     While (vArray(tmpLow) < pivot And tmpLow < inHi)
        tmpLow = tmpLow + 1
     Wend

     While (pivot < vArray(tmpHi) And tmpHi > inLow)
        tmpHi = tmpHi - 1
     Wend

     If (tmpLow <= tmpHi) Then
        tmpSwap = vArray(tmpLow)
        vArray(tmpLow) = vArray(tmpHi)
        vArray(tmpHi) = tmpSwap
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
     End If

  Wend

  If (inLow < tmpHi) Then QuickSort vArray, inLow, tmpHi
  If (tmpLow < inHi) Then QuickSort vArray, tmpLow, inHi

End Sub

Private Sub QuickSort2(ByRef Field, ByVal LB As Long, ByVal UB As Long)
    Dim P1 As Long, P2 As Long, Ref As String, TEMP As String

    P1 = LB
    P2 = UB
    Ref = Field((P1 + P2) / 2)

    Do
        Do While (Field(P1) < Ref)
            P1 = P1 + 1
        Loop

        Do While (Field(P2) > Ref)
            P2 = P2 - 1
        Loop

        If P1 <= P2 Then
            TEMP = Field(P1)
            Field(P1) = Field(P2)
            Field(P2) = TEMP

            P1 = P1 + 1
            P2 = P2 - 1
        End If
    Loop Until (P1 > P2)

    If LB < P2 Then Call QuickSort(Field, LB, P2)
    If P1 < UB Then Call QuickSort(Field, P1, UB)
End Sub
Public Sub Sort(ByRef mtz)
    QuickSort2 mtz, LBound(mtz), UBound(mtz)
    'InsertionSort mtz, LBound(mtz), UBound(mtz)
End Sub
