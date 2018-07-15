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
'Private Type defEvent
'    Second As Long
'    Message As String
'    Signal As Long
'    Raised As Boolean
'End Type
Private mName As String
Private mMessage As String
Private mSeconds As Collection
Private mEvents As Collection
Private mRecallClass As clsEvent 'Not in use

Public Property Let Name(NewName As String)
    mName = NewName
End Property

Public Property Get Name() As String
    Name = mName
End Property

Public Property Let Message(NewMessage As String)
    mMessage = NewMessage
End Property

Public Property Get Message() As String
    Message = mMessage
End Property

Public Property Let RecallClass(NewRecallClass As String)
    mRecallClass = NewRecallClass
End Property

Public Property Get RecallClass() As String
    RecallClass = mRecallClass
End Property

'Public Property Let RecallEnabled(NewRecallEnabled As Boolean)
'    mRecallClass = NewRecallEnabled
'End Property

'Public Property Get RecallEnabled() As Boolean
'    RecallEnabled = mRecallEnabled
'End Property

Public Function NewEvent(Second As Long, Message As String, Signal As Long, Silent As Boolean, Raised As Boolean) As clsEvent
Dim MyEvent As clsEvent
'Add the Second into the Seconds Collection
    On Error GoTo SecondExists
    mSeconds.Add Message, CStr(Second)
Exists:
    On Error GoTo 0
   
'Add the Event into the Events Collection
    Set MyEvent = New clsEvent
    MyEvent.Second = Second
    MyEvent.Message = Message
    MyEvent.Signal = Signal
    MyEvent.Silent = Silent
    MyEvent.Raised = Raised
    On Error GoTo DuplicateEvent
    mEvents.Add MyEvent, MyEvent.key
    MyEvent.Index = mEvents.Count
    On Error GoTo 0
    Set NewEvent = MyEvent
    Set MyEvent = Nothing
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
DuplicateEvent:
    MsgBox "You cannot set the Time [" & MyEvent.Second * Multiplier _
    & "] for the same Signal [" & MyEvent.Signal & "]" & vbCrLf _
    & "This Event will be ignored", vbExclamation, "NewEvent"
End Function

'Check to see if this elapsed second is set up in
'the mSeconds collection. If it is we need to
'generate some events
Public Function IsEventDue(Second As Long) As Boolean
Dim data As String
    On Error GoTo KeyNotFound
'cant use if mseconds.item(cstr(second)) because
'0 seconds returns false !!!
    data = mSeconds.Item(CStr(Second))
    IsEventDue = True
KeyNotFound:
'Must see if we can turn off recall
#If False Then
    If frmMain.cmdRecall.BackColor = vbRed Then
        IsEventDue = True
    End If
#End If
End Function

Public Function GetEventDue(Second As Long) As String
Dim MyEvent As clsEvent

    For Each MyEvent In Myprofile
        If MyEvent.Second > Second Then
            GetEventDue = mSeconds.Item(CStr(MyEvent.Second))
            Exit Function
        End If
    Next MyEvent

KeyNotFound:
    GetEventDue = "Finish"
End Function
'Must change the procedure ID see Programmers Guide P489
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mEvents.[_NewEnum]
End Property

Public Property Get FirstEvent() As clsEvent
    If mEvents.Count > 0 Then
        Set FirstEvent = mEvents.Item(1)
    End If
End Property

Public Property Get LastEvent() As clsEvent
    If mEvents.Count > 0 Then
        Set LastEvent = mEvents.Item(mEvents.Count)
    End If
End Property

'Programmers Guide P489
#If False Then
Public Function Item(ByVal Index As Variant) As clsEvent
Attribute Item.VB_UserMemId = 0
    Set Item = mEvents.Item
End Function
#End If

Public Function NewCommand(Second As Long, Recall As String) As clsEvent
Dim MyEvent As clsEvent
'Add the Second into the Seconds Collection
    On Error GoTo CreateSecond
    Set MyEvent = mSeconds.Item(CStr(Second))
Exists:
    On Error GoTo 0
'Add the Event into the Events Collection
    
    Set MyEvent = Nothing
    Exit Function
CreateSecond:
    mSeconds.Add Message, CStr(Second)
    Select Case Err.Number
    Case Is = 457   'exists
    Case Else
        MsgBox "Error " & Err.Number & vbCrLf _
        & Err.Description, vbExclamation + vbOKOnly, _
        "New Event Error"
    End Select
    Resume Next
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
