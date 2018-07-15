VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type defEvent
    Second As Long
    Class As Long
    Signal As Long
    State As Boolean
End Type

Private mName As String
Private mSeconds As Collection
Private mEvents As Collection

Public Property Let Name(NewName As String)
    mName = NewName
End Property

Public Property Get Name() As String
    Name = mName
End Property

Public Function NewEvent(Second As Long, Class As Long, Signal As Long, State As Boolean) As clsEvent
Dim myEvent As clsEvent
    
    On Error GoTo SecondExists
    mSeconds.Add Second, CStr(Second)
Exists:
    On Error GoTo 0
    Set myEvent = New clsEvent
    myEvent.Second = Second
    myEvent.Class = Class
    myEvent.Signal = Signal
    myEvent.State = State
    mEvents.Add myEvent, myEvent.Key
    Set NewEvent = myEvent
    Set myEvent = Nothing
    Exit Function
SecondExists:
    Select Case Err.Number
    Case Is = 457   'exists
    Case Else
        MsgBox "Error " & Err.Number & vbCrLf _
        & Err.Description, vbExclamation + vbOKOnly, _
        "New Event Error"
    End Select
    Resume Exists
End Function
Private Sub Class_Initialize()
Debug.Print "Create clsProfile"
    'creates the collection when this class is created
    Set mSeconds = New Collection
    Set mEvents = New Collection
End Sub

Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mEvents = Nothing
    Set mSeconds = Nothing
Debug.Print "Destroy clsProfile"
End Sub
