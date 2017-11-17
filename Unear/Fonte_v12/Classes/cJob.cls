VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cJob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mNome As String
Private mStartTime As Date
Private mRunAction As String
Private mStatus As eJobStatus
Private mDurationAfterFinish As Variant
Private mCountExecuted As Long

Public Enum eJobStatus
    NotExecuted = 0
    Executed = 1
End Enum

Private m_objJobTimer As cTimer

Private m_sJobID As String


Public Property Get nome() As String
    nome = mNome
End Property
Public Property Let nome(value As String)
    mNome = value
End Property

Public Property Get JobTimer() As cTimer
    Set JobTimer = m_objJobTimer
End Property

Public Property Set JobTimer(ByVal objNewValue As cTimer)
    Set m_objJobTimer = objNewValue
End Property

Public Property Get StartTime() As Date
    StartTime = mStartTime
End Property

Public Property Let StartTime(value As Date)
    mStartTime = value
End Property

Public Property Get RunAction() As String
    RunAction = mRunAction
End Property

Public Property Let RunAction(value As String)
    mRunAction = value
End Property

Public Property Get status() As eJobStatus
    status = mStatus
End Property
Public Property Let status(value As eJobStatus)
    mStatus = value
End Property

Public Property Get DurationAfterFinish() As Variant
    DurationAfterFinish = mDurationAfterFinish
End Property

Public Property Let DurationAfterFinish(ByVal vNewValue As Variant)
    mDurationAfterFinish = vNewValue
End Property

Public Property Get CountExecuted() As Long
    CountExecuted = mCountExecuted
End Property
Public Property Let CountExecuted(value As Long)
    mCountExecuted = value
End Property

Public Property Get JobID() As String
    JobID = m_sJobID
End Property

Public Property Let JobID(ByVal sNewValue As String)
    m_sJobID = sNewValue
End Property
