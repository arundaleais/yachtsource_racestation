VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mIndex As Long
Private mSecond As Long
Private mMessage As String
Private mSignal As Long
Private mRaised As Boolean
Private mEnabled As Boolean

Public Property Let Second(NewSecond As Long)
vbwProfiler.vbwProcIn 73
vbwProfiler.vbwExecuteLine 1344
    mSecond = NewSecond
vbwProfiler.vbwProcOut 73
vbwProfiler.vbwExecuteLine 1345
End Property

Public Property Get Second() As Long
vbwProfiler.vbwProcIn 74
vbwProfiler.vbwExecuteLine 1346
    Second = mSecond
vbwProfiler.vbwProcOut 74
vbwProfiler.vbwExecuteLine 1347
End Property

Public Property Get Index() As Long
vbwProfiler.vbwProcIn 75
vbwProfiler.vbwExecuteLine 1348
    Index = mIndex
vbwProfiler.vbwProcOut 75
vbwProfiler.vbwExecuteLine 1349
End Property

Public Property Let Index(Item As Long)
vbwProfiler.vbwProcIn 76
vbwProfiler.vbwExecuteLine 1350
    mIndex = Item
vbwProfiler.vbwProcOut 76
vbwProfiler.vbwExecuteLine 1351
End Property

Public Property Let Message(NewMessage As String)
'Combine message is separate in .ini file
vbwProfiler.vbwProcIn 77
vbwProfiler.vbwExecuteLine 1352
    If mMessage <> "" Then
vbwProfiler.vbwExecuteLine 1353
         mMessage = mMessage & ", "
    End If
vbwProfiler.vbwExecuteLine 1354 'B
vbwProfiler.vbwExecuteLine 1355
    mMessage = mMessage & NewMessage
vbwProfiler.vbwProcOut 77
vbwProfiler.vbwExecuteLine 1356
End Property

Public Property Get Message() As String
vbwProfiler.vbwProcIn 78
vbwProfiler.vbwExecuteLine 1357
    Message = mMessage
vbwProfiler.vbwProcOut 78
vbwProfiler.vbwExecuteLine 1358
End Property

Public Property Let Signal(NewSignal As Long)
vbwProfiler.vbwProcIn 79
vbwProfiler.vbwExecuteLine 1359
    mSignal = NewSignal
vbwProfiler.vbwProcOut 79
vbwProfiler.vbwExecuteLine 1360
End Property

Public Property Get Signal() As Long
vbwProfiler.vbwProcIn 80
vbwProfiler.vbwExecuteLine 1361
    Signal = mSignal
vbwProfiler.vbwProcOut 80
vbwProfiler.vbwExecuteLine 1362
End Property

Public Property Let Raised(NewRaised As Boolean)
vbwProfiler.vbwProcIn 81
vbwProfiler.vbwExecuteLine 1363
If mRaised <> NewRaised Then
'MsgBox "signal " & mSignal & " Raised changing to " & NewRaised
End If
vbwProfiler.vbwExecuteLine 1364 'B
vbwProfiler.vbwExecuteLine 1365
    mRaised = NewRaised
vbwProfiler.vbwProcOut 81
vbwProfiler.vbwExecuteLine 1366
End Property

Public Property Get Raised() As Boolean
vbwProfiler.vbwProcIn 82
vbwProfiler.vbwExecuteLine 1367
    Raised = mRaised
vbwProfiler.vbwProcOut 82
vbwProfiler.vbwExecuteLine 1368
End Property

Public Property Let Enabled(NewEnabled As Boolean)
vbwProfiler.vbwProcIn 83
vbwProfiler.vbwExecuteLine 1369
    mEnabled = NewEnabled
vbwProfiler.vbwProcOut 83
vbwProfiler.vbwExecuteLine 1370
End Property

Public Property Get Enabled() As Boolean
vbwProfiler.vbwProcIn 84
vbwProfiler.vbwExecuteLine 1371
    Enabled = mEnabled
vbwProfiler.vbwProcOut 84
vbwProfiler.vbwExecuteLine 1372
End Property

Public Property Get key() As String
vbwProfiler.vbwProcIn 85
vbwProfiler.vbwExecuteLine 1373
    key = Format$(mSecond, "00000") & ":" _
    & Format$(mSignal, "000") & ":" & mRaised
vbwProfiler.vbwProcOut 85
vbwProfiler.vbwExecuteLine 1374
End Property


