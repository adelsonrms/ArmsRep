VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Private m_sTextString As String


Public Property Get TextString() As String
    TextString = ClearString(m_sTextString)
End Property

Public Property Let TextString(ByVal sTextString As String)
    m_sTextString = sTextString
End Property

Function Replace(OldText As String, newText As String) As cString
    m_sTextString = VBA.Replace(m_sTextString, OldText, newText)
    Set Replace = Me
End Function

Function Format(pFormat As String) As cString
    m_sTextString = VBA.Format(m_sTextString, pFormat)
    Set Format = Me
End Function

Function Contains(pValue As String) As Boolean
    Contains = VBA.InStr(m_sTextString, pValue) > 0
End Function

Function Equals(pValue As String) As Boolean
    Equals = Trim(m_sTextString) = pValue
End Function

Function ClearString(pString As String) As String
    Dim i As Integer
    Dim pOutput As String
    For i = 1 To VBA.Len(pString)
        If VBA.Asc(VBA.Mid(pString, i, 1)) > 31 And VBA.Asc(VBA.Mid(pString, i, 1)) < 121 Then
            pOutput = pOutput & VBA.Mid(pString, i, 1)
        End If
    Next i
    ClearString = VBA.Trim(pOutput)
End Function

Function RemoveCaracteres(ParamArray arrCaracteres() As Variant)
    Dim contador            As Byte, Aux As String
    Aux = m_sTextString
    For contador = 0 To UBound(arrCaracteres) Step 1
        Aux = VBA.Replace(Aux, arrCaracteres(contador), "")
    Next contador
    RemoveCaracteres = Aux
End Function


