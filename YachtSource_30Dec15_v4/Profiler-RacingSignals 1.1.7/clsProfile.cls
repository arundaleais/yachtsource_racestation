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
vbwProfiler.vbwProcIn 86
vbwProfiler.vbwExecuteLine 1375
    mName = NewName
vbwProfiler.vbwProcOut 86
vbwProfiler.vbwExecuteLine 1376
End Property

Public Property Get Name() As String
vbwProfiler.vbwProcIn 87
vbwProfiler.vbwExecuteLine 1377
    Name = mName
vbwProfiler.vbwProcOut 87
vbwProfiler.vbwExecuteLine 1378
End Property

Public Property Let Message(NewMessage As String)
vbwProfiler.vbwProcIn 88
vbwProfiler.vbwExecuteLine 1379
    mMessage = NewMessage
vbwProfiler.vbwProcOut 88
vbwProfiler.vbwExecuteLine 1380
End Property

Public Property Get Message() As String
vbwProfiler.vbwProcIn 89
vbwProfiler.vbwExecuteLine 1381
    Message = mMessage
vbwProfiler.vbwProcOut 89
vbwProfiler.vbwExecuteLine 1382
End Property

Public Property Let RecallClass(NewRecallClass As String)
vbwProfiler.vbwProcIn 90
vbwProfiler.vbwExecuteLine 1383
    mRecallClass = NewRecallClass
vbwProfiler.vbwProcOut 90
vbwProfiler.vbwExecuteLine 1384
End Property

Public Property Get RecallClass() As String
vbwProfiler.vbwProcIn 91
vbwProfiler.vbwExecuteLine 1385
    RecallClass = mRecallClass
vbwProfiler.vbwProcOut 91
vbwProfiler.vbwExecuteLine 1386
End Property

'Public Property Let RecallEnabled(NewRecallEnabled As Boolean)
'    mRecallClass = NewRecallEnabled
'End Property

'Public Property Get RecallEnabled() As Boolean
'    RecallEnabled = mRecallEnabled
'End Property

Public Function NewEvent(Second As Long, Message As String, Signal As Long, Raised As Boolean) As clsEvent
vbwProfiler.vbwProcIn 92
Dim MyEvent As clsEvent
'Add the Second into the Seconds Collection
vbwProfiler.vbwExecuteLine 1387
    On Error GoTo SecondExists
vbwProfiler.vbwExecuteLine 1388
    mSeconds.Add Message, CStr(Second)
Exists:
vbwProfiler.vbwExecuteLine 1389
    On Error GoTo 0

'Add the Event into the Events Collection
vbwProfiler.vbwExecuteLine 1390
    Set MyEvent = New clsEvent
vbwProfiler.vbwExecuteLine 1391
    MyEvent.Second = Second
vbwProfiler.vbwExecuteLine 1392
    MyEvent.Message = Message
vbwProfiler.vbwExecuteLine 1393
    MyEvent.Signal = Signal
vbwProfiler.vbwExecuteLine 1394
    MyEvent.Raised = Raised
vbwProfiler.vbwExecuteLine 1395
    On Error GoTo DuplicateEvent
vbwProfiler.vbwExecuteLine 1396
    mEvents.Add MyEvent, MyEvent.key
vbwProfiler.vbwExecuteLine 1397
    MyEvent.Index = mEvents.Count
vbwProfiler.vbwExecuteLine 1398
    On Error GoTo 0
vbwProfiler.vbwExecuteLine 1399
    Set NewEvent = MyEvent
vbwProfiler.vbwExecuteLine 1400
    Set MyEvent = Nothing
vbwProfiler.vbwProcOut 92
vbwProfiler.vbwExecuteLine 1401
    Exit Function
SecondExists:
vbwProfiler.vbwExecuteLine 1402
    Select Case Err.Number
'vbwLine 1403:    Case Is = 457 'exists
    Case Is = IIf(vbwProfiler.vbwExecuteLine(1403), VBWPROFILER_EMPTY, _
        457 )'exists
    Case Else
vbwProfiler.vbwExecuteLine 1404 'B
vbwProfiler.vbwExecuteLine 1405
        MsgBox "Error " & Err.Number & vbCrLf _
        & Err.Description, vbExclamation + vbOKOnly, _
        "New Event Error"
    End Select
vbwProfiler.vbwExecuteLine 1406 'B
vbwProfiler.vbwExecuteLine 1407
    Resume Exists
DuplicateEvent:
vbwProfiler.vbwExecuteLine 1408
    MsgBox "You cannot set the Time [" & MyEvent.Second * Multiplier _
    & "] for the same Signal [" & MyEvent.Signal & "]" & vbCrLf _
    & "This Event will be ignored", vbExclamation, "NewEvent"
vbwProfiler.vbwProcOut 92
vbwProfiler.vbwExecuteLine 1409
End Function

'Check to see if this elapsed second is set up in
'the mSeconds collection. If it is we need to
'generate some events
Public Function IsEventDue(Second As Long) As Boolean
vbwProfiler.vbwProcIn 93
Dim data As String
vbwProfiler.vbwExecuteLine 1410
    On Error GoTo KeyNotFound
'cant use if mseconds.item(cstr(second)) because
'0 seconds returns false !!!
vbwProfiler.vbwExecuteLine 1411
    data = mSeconds.Item(CStr(Second))
vbwProfiler.vbwExecuteLine 1412
    IsEventDue = True
KeyNotFound:
'Must see if we can turn off recall
#If False Then
    If frmMain.cmdRecall.BackColor = vbRed Then
        IsEventDue = True
    End If
#End If
vbwProfiler.vbwProcOut 93
vbwProfiler.vbwExecuteLine 1413
End Function

Public Function GetEventDue(Second As Long) As String
vbwProfiler.vbwProcIn 94
Dim MyEvent As clsEvent

vbwProfiler.vbwExecuteLine 1414
    For Each MyEvent In Myprofile
vbwProfiler.vbwExecuteLine 1415
        If MyEvent.Second > Second Then
vbwProfiler.vbwExecuteLine 1416
            GetEventDue = mSeconds.Item(CStr(MyEvent.Second))
vbwProfiler.vbwProcOut 94
vbwProfiler.vbwExecuteLine 1417
            Exit Function
        End If
vbwProfiler.vbwExecuteLine 1418 'B
vbwProfiler.vbwExecuteLine 1419
    Next MyEvent

KeyNotFound:
vbwProfiler.vbwExecuteLine 1420
    GetEventDue = "Finish"
vbwProfiler.vbwProcOut 94
vbwProfiler.vbwExecuteLine 1421
End Function
'Must change the procedure ID see Programmers Guide P489
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    'this property allows you to enumerate
    'this collection with the For...Each syntax
vbwProfiler.vbwProcIn 95
vbwProfiler.vbwExecuteLine 1422
    Set NewEnum = mEvents.[_NewEnum]
vbwProfiler.vbwProcOut 95
vbwProfiler.vbwExecuteLine 1423
End Property

Public Property Get FirstEvent() As clsEvent
vbwProfiler.vbwProcIn 96
vbwProfiler.vbwExecuteLine 1424
    If mEvents.Count > 0 Then
vbwProfiler.vbwExecuteLine 1425
        Set FirstEvent = mEvents.Item(1)
    End If
vbwProfiler.vbwExecuteLine 1426 'B
vbwProfiler.vbwProcOut 96
vbwProfiler.vbwExecuteLine 1427
End Property

Public Property Get LastEvent() As clsEvent
vbwProfiler.vbwProcIn 97
vbwProfiler.vbwExecuteLine 1428
    If mEvents.Count > 0 Then
vbwProfiler.vbwExecuteLine 1429
        Set LastEvent = mEvents.Item(mEvents.Count)
    End If
vbwProfiler.vbwExecuteLine 1430 'B
vbwProfiler.vbwProcOut 97
vbwProfiler.vbwExecuteLine 1431
End Property

'Programmers Guide P489
#If False Then
Public Function Item(ByVal Index As Variant) As clsEvent
Attribute Item.VB_UserMemId = 0
    Set Item = mEvents.Item
End Function
#End If

Public Function NewRecall(Second As Long, Recall As String) As clsEvent
vbwProfiler.vbwProcIn 98
Dim MyEvent As clsEvent
'Add the Second into the Seconds Collection
vbwProfiler.vbwExecuteLine 1432
    On Error GoTo CreateSecond
vbwProfiler.vbwExecuteLine 1433
    Set MyEvent = mSeconds.Item(CStr(Second))
Exists:
vbwProfiler.vbwExecuteLine 1434
    On Error GoTo 0
'Add the Event into the Events Collection

vbwProfiler.vbwExecuteLine 1435
    Set MyEvent = Nothing
vbwProfiler.vbwProcOut 98
vbwProfiler.vbwExecuteLine 1436
    Exit Function
CreateSecond:
vbwProfiler.vbwExecuteLine 1437
    mSeconds.Add Message, CStr(Second)
vbwProfiler.vbwExecuteLine 1438
    Select Case Err.Number
'vbwLine 1439:    Case Is = 457 'exists
    Case Is = IIf(vbwProfiler.vbwExecuteLine(1439), VBWPROFILER_EMPTY, _
        457 )'exists
    Case Else
vbwProfiler.vbwExecuteLine 1440 'B
vbwProfiler.vbwExecuteLine 1441
        MsgBox "Error " & Err.Number & vbCrLf _
        & Err.Description, vbExclamation + vbOKOnly, _
        "New Event Error"
    End Select
vbwProfiler.vbwExecuteLine 1442 'B
vbwProfiler.vbwExecuteLine 1443
    Resume Next
vbwProfiler.vbwProcOut 98
vbwProfiler.vbwExecuteLine 1444
End Function


Private Sub Class_Initialize()
vbwProfiler.vbwProcIn 99
vbwProfiler.vbwExecuteLine 1445
Debug.Print "Create clsProfile"
    'creates the collection when this class is created
vbwProfiler.vbwExecuteLine 1446
    Set mSeconds = New Collection
vbwProfiler.vbwExecuteLine 1447
    Set mEvents = New Collection
vbwProfiler.vbwProcOut 99
vbwProfiler.vbwExecuteLine 1448
End Sub

Private Sub Class_Terminate()
    'destroys collection when this class is terminated
vbwProfiler.vbwProcIn 100
vbwProfiler.vbwExecuteLine 1449
    Set mEvents = Nothing
vbwProfiler.vbwExecuteLine 1450
    Set mSeconds = Nothing
vbwProfiler.vbwExecuteLine 1451
Debug.Print "Destroy clsProfile"
vbwProfiler.vbwProcOut 100
vbwProfiler.vbwExecuteLine 1452
End Sub


