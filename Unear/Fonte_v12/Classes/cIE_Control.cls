VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cIE_Control"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private m_Element As Object
Private m_IE_Tab As cIE_Tab
Private m_Name As String
Private m_ID As String
Private m_Value As String
Private m_TypeElement As String
Private m_Controls As Collection

Public Property Get element() As Object
    Set element = m_Element
End Property

Public Property Set element(ByVal vNewValue As Object)
    Set m_Element = vNewValue
End Property

Public Property Get IE_Tab() As cIE_Tab
    Set IE_Tab = m_IE_Tab
End Property

Public Property Set IE_Tab(ByVal vNewValue As cIE_Tab)
    Set m_IE_Tab = vNewValue
    Call RefreshChildren
End Property

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let Name(ByVal vNewValue As String)
    m_Name = vNewValue
End Property

Public Property Get ID() As String
    ID = m_ID
End Property

Public Property Let ID(ByVal vNewValue As String)
    m_ID = vNewValue
End Property

Public Property Get value() As String
    value = m_Value
End Property

Public Property Let value(ByVal vNewValue As String)
    m_Value = vNewValue
    Call sendValue
End Property

Public Property Get TypeElement() As String
    TypeElement = m_TypeElement
End Property

Public Property Let TypeElement(ByVal vNewValue As String)
    m_TypeElement = vNewValue
End Property

Function Click() As Boolean
    Click = m_IE_Tab.Click(m_Element)
End Function

Private Sub sendValue()
    If TypeElement = "INPUT" Then
        m_Element.value = m_Value
    End If
End Sub

'Private m_Controls As Collection
Public Property Get Controls() As Collection
    If m_Controls Is Nothing Then Set m_Controls = New Collection
    Set Controls = m_Controls
End Property

Public Property Set Controls(ByVal vNewValue As Collection)
    Set Controls = vNewValue
End Property

Private Function RefreshChildren()
    Set m_Controls = m_IE_Tab.GetControls(Me.element, False, False)
End Function
