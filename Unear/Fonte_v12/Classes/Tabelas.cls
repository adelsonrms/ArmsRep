VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private m_Name          As String
Public Property Get Name() As String
    Name = m_Name
End Property
Public Property Let Name(ByVal valor As String)
    m_Name = valor
End Property

Public Function Abre() As Object
On Error GoTo TrataErro
    Set Abre = Access.CurrentDb.OpenRecordset(Me.Name)
Exit Function
TrataErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "Tabela.Abre()", Erl)
    Set Abre = Nothing
End Function





