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
vbwProfiler.vbwProcIn 97
vbwProfiler.vbwExecuteLine 1609
    mName = NewName
vbwProfiler.vbwProcOut 97
vbwProfiler.vbwExecuteLine 1610
End Property

Public Property Get Name() As String
vbwProfiler.vbwProcIn 98
vbwProfiler.vbwExecuteLine 1611
    Name = mName
vbwProfiler.vbwProcOut 98
vbwProfiler.vbwExecuteLine 1612
End Property

Public Property Let Message(NewMessage As String)
vbwProfiler.vbwProcIn 99
vbwProfiler.vbwExecuteLine 1613
    mMessage = NewMessage
vbwProfiler.vbwProcOut 99
vbwProfiler.vbwExecuteLine 1614
End Property

Public Property Get Message() As String
vbwProfiler.vbwProcIn 100
vbwProfiler.vbwExecuteLine 1615
    Message = mMessage
vbwProfiler.vbwProcOut 100
vbwProfiler.vbwExecuteLine 1616
End Property

Public Property Let RecallClass(NewRecallClass As String)
vbwProfiler.vbwProcIn 101
vbwProfiler.vbwExecuteLine 1617
    mRecallClass = NewRecallClass
vbwProfiler.vbwProcOut 101
vbwProfiler.vbwExecuteLine 1618
End Property

Public Property Get RecallClass() As String
vbwProfiler.vbwProcIn 102
vbwProfiler.vbwExecuteLine 1619
    RecallClass = mRecallClass
vbwProfiler.vbwProcOut 102
vbwProfiler.vbwExecuteLine 1620
End Property

'Public Property Let RecallEnabled(NewRecallEnabled As Boolean)
'    mRecallClass = NewRecallEnabled
'End Property

'Public Property Get RecallEnabled() As Boolean
'    RecallEnabled = mRecallEnabled
'End Property

Public Function NewEvent(Second As Long, Message As String, Signal As Long, Silent As Boolean, Raised As Boolean) As clsEvent
vbwProfiler.vbwProcIn 103
Dim MyEvent As clsEvent
'Add the Second into the Seconds Collection
vbwProfiler.vbwExecuteLine 1621
    On Error GoTo SecondExists
vbwProfiler.vbwExecuteLine 1622
    mSeconds.Add Message, CStr(Second)
Exists:
vbwProfiler.vbwExecuteLine 1623
    On Error GoTo 0

'Add the Event into the Events Collection
vbwProfiler.vbwExecuteLine 1624
    Set MyEvent = New clsEvent
vbwProfiler.vbwExecuteLine 1625
    MyEvent.Second = Second
vbwProfiler.vbwExecuteLine 1626
    MyEvent.Message = Message
vbwProfiler.vbwExecuteLine 1627
    MyEvent.Signal = Signal
vbwProfiler.vbwExecuteLine 1628
    MyEvent.Silent = Silent
vbwProfiler.vbwExecuteLine 1629
    MyEvent.Raised = Raised
vbwProfiler.vbwExecuteLine 1630
    On Error GoTo DuplicateEvent
vbwProfiler.vbwExecuteLine 1631
    mEvents.Add MyEvent, MyEvent.key
vbwProfiler.vbwExecuteLine 1632
    MyEvent.Index = mEvents.Count
vbwProfiler.vbwExecuteLine 1633
    On Error GoTo 0
vbwProfiler.vbwExecuteLine 1634
    Set NewEvent = MyEvent
vbwProfiler.vbwExecuteLine 1635
    Set MyEvent = Nothing
vbwProfiler.vbwProcOut 103
vbwProfiler.vbwExecuteLine 1636
    Exit Function
SecondExists:
vbwProfiler.vbwExecuteLine 1637
    Select Case Err.Number
'vbwLine 1638:    Case Is = 457 'exists
    Case Is = IIf(vbwProfiler.vbwExecuteLine(1638), VBWPROFILER_EMPTY, _
        457 )'exists
    Case Else
vbwProfiler.vbwExecuteLine 1639 'B
vbwProfiler.vbwExecuteLine 1640
        MsgBox "Error " & Err.Number & vbCrLf _
        & Err.Description, vbExclamation + vbOKOnly, _
        "New Event Error"
    End Select
vbwProfiler.vbwExecuteLine 1641 'B
vbwProfiler.vbwExecuteLine 1642
    Resume Exists
DuplicateEvent:
vbwProfiler.vbwExecuteLine 1643
    MsgBox "You cannot set the Time [" & MyEvent.Second * Multiplier _
    & "] for the same Signal [" & MyEvent.Signal & "]" & vbCrLf _
    & "This Event will be ignored", vbExclamation, "NewEvent"
vbwProfiler.vbwProcOut 103
vbwProfiler.vbwExecuteLine 1644
End Function

'Check to see if this elapsed second is set up in
'the mSeconds collection. If it is we need to
'generate some events
Public Function IsEventDue(Second As Long) As Boolean
vbwProfiler.vbwProcIn 104
Dim data As String
vbwProfiler.vbwExecuteLine 1645
    On Error GoTo KeyNotFound
'cant use if mseconds.item(cstr(second)) because
'0 seconds returns false !!!
vbwProfiler.vbwExecuteLine 1646
    data = mSeconds.Item(CStr(Second))
vbwProfiler.vbwExecuteLine 1647
    IsEventDue = True
KeyNotFound:
'Must see if we can turn off recall
#If False Then
    If frmMain.cmdRecall.BackColor = vbRed Then
        IsEventDue = True
    End If
#End If
vbwProfiler.vbwProcOut 104
vbwProfiler.vbwExecuteLine 1648
End Function

Public Function GetEventDue(Second As Long) As String
vbwProfiler.vbwProcIn 105
Dim MyEvent As clsEvent

vbwProfiler.vbwExecuteLine 1649
    For Each MyEvent In Myprofile
vbwProfiler.vbwExecuteLine 1650
        If MyEvent.Second > Second Then
vbwProfiler.vbwExecuteLine 1651
            GetEventDue = mSeconds.Item(CStr(MyEvent.Second))
vbwProfiler.vbwProcOut 105
vbwProfiler.vbwExecuteLine 1652
            Exit Function
        End If
vbwProfiler.vbwExecuteLine 1653 'B
vbwProfiler.vbwExecuteLine 1654
    Next MyEvent

KeyNotFound:
vbwProfiler.vbwExecuteLine 1655
    GetEventDue = "Finish"
vbwProfiler.vbwProcOut 105
vbwProfiler.vbwExecuteLine 1656
End Function
'Must change the procedure ID see Programmers Guide P489
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    'this property allows you to enumerate
    'this collection with the For...Each syntax
vbwProfiler.vbwProcIn 106
vbwProfiler.vbwExecuteLine 1657
    Set NewEnum = mEvents.[_NewEnum]
vbwProfiler.vbwProcOut 106
vbwProfiler.vbwExecuteLine 1658
End Property

Public Property Get FirstEvent() As clsEvent
vbwProfiler.vbwProcIn 107
vbwProfiler.vbwExecuteLine 1659
    If mEvents.Count > 0 Then
vbwProfiler.vbwExecuteLine 1660
        Set FirstEvent = mEvents.Item(1)
    End If
vbwProfiler.vbwExecuteLine 1661 'B
vbwProfiler.vbwProcOut 107
vbwProfiler.vbwExecuteLine 1662
End Property

Public Property Get LastEvent() As clsEvent
vbwProfiler.vbwProcIn 108
vbwProfiler.vbwExecuteLine 1663
    If mEvents.Count > 0 Then
vbwProfiler.vbwExecuteLine 1664
        Set LastEvent = mEvents.Item(mEvents.Count)
    End If
vbwProfiler.vbwExecuteLine 1665 'B
vbwProfiler.vbwProcOut 108
vbwProfiler.vbwExecuteLine 1666
End Property

'Programmers Guide P489
#If False Then
Public Function Item(ByVal Index As Variant) As clsEvent
Attribute Item.VB_UserMemId = 0
    Set Item = mEvents.Item
End Function
#End If

Public Function NewCommand(Second As Long, Recall As String) As clsEvent
vbwProfiler.vbwProcIn 109
Dim MyEvent As clsEvent
'Add the Second into the Seconds Collection
vbwProfiler.vbwExecuteLine 1667
    On Error GoTo CreateSecond
vbwProfiler.vbwExecuteLine 1668
    Set MyEvent = mSeconds.Item(CStr(Second))
Exists:
vbwProfiler.vbwExecuteLine 1669
    On Error GoTo 0
'Add the Event into the Events Collection

vbwProfiler.vbwExecuteLine 1670
    Set MyEvent = Nothing
vbwProfiler.vbwProcOut 109
vbwProfiler.vbwExecuteLine 1671
    Exit Function
CreateSecond:
vbwProfiler.vbwExecuteLine 1672
    mSeconds.Add Message, CStr(Second)
vbwProfiler.vbwExecuteLine 1673
    Select Case Err.Number
'vbwLine 1674:    Case Is = 457 'exists
    Case Is = IIf(vbwProfiler.vbwExecuteLine(1674), VBWPROFILER_EMPTY, _
        457 )'exists
    Case Else
vbwProfiler.vbwExecuteLine 1675 'B
vbwProfiler.vbwExecuteLine 1676
        MsgBox "Error " & Err.Number & vbCrLf _
        & Err.Description, vbExclamation + vbOKOnly, _
        "New Event Error"
    End Select
vbwProfiler.vbwExecuteLine 1677 'B
vbwProfiler.vbwExecuteLine 1678
    Resume Next
vbwProfiler.vbwProcOut 109
vbwProfiler.vbwExecuteLine 1679
End Function


Private Sub Class_Initialize()
vbwProfiler.vbwProcIn 110
vbwProfiler.vbwExecuteLine 1680
Debug.Print "Create clsProfile"
    'creates the collection when this class is created
vbwProfiler.vbwExecuteLine 1681
    Set mSeconds = New Collection
vbwProfiler.vbwExecuteLine 1682
    Set mEvents = New Collection
vbwProfiler.vbwProcOut 110
vbwProfiler.vbwExecuteLine 1683
End Sub

Private Sub Class_Terminate()
    'destroys collection when this class is terminated
vbwProfiler.vbwProcIn 111
vbwProfiler.vbwExecuteLine 1684
    Set mEvents = Nothing
vbwProfiler.vbwExecuteLine 1685
    Set mSeconds = Nothing
vbwProfiler.vbwExecuteLine 1686
Debug.Print "Destroy clsProfile"
vbwProfiler.vbwProcOut 111
vbwProfiler.vbwExecuteLine 1687
End Sub


