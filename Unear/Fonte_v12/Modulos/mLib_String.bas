Attribute VB_Name = "mLib_String"
Option Compare Database

Function NewString(pString As String) As cString
    Dim cs As New cString
    cs.TextString = pString
    Set NewString = cs
End Function
