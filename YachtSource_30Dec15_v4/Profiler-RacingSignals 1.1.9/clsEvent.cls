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
Private mSilent As Boolean

Public Property Let Second(NewSecond As Long)
vbwProfiler.vbwProcIn 84
vbwProfiler.vbwExecuteLine 1578
    mSecond = NewSecond
vbwProfiler.vbwProcOut 84
vbwProfiler.vbwExecuteLine 1579
End Property

Public Property Get Second() As Long
vbwProfiler.vbwProcIn 85
vbwProfiler.vbwExecuteLine 1580
    Second = mSecond
vbwProfiler.vbwProcOut 85
vbwProfiler.vbwExecuteLine 1581
End Property

Public Property Get Index() As Long
vbwProfiler.vbwProcIn 86
vbwProfiler.vbwExecuteLine 1582
    Index = mIndex
vbwProfiler.vbwProcOut 86
vbwProfiler.vbwExecuteLine 1583
End Property

Public Property Let Index(Item As Long)
vbwProfiler.vbwProcIn 87
vbwProfiler.vbwExecuteLine 1584
    mIndex = Item
vbwProfiler.vbwProcOut 87
vbwProfiler.vbwExecuteLine 1585
End Property

Public Property Let Message(NewMessage As String)
'Combine message is separate in .ini file
vbwProfiler.vbwProcIn 88
vbwProfiler.vbwExecuteLine 1586
    If mMessage <> "" Then
vbwProfiler.vbwExecuteLine 1587
         mMessage = mMessage & ", "
    End If
vbwProfiler.vbwExecuteLine 1588 'B
vbwProfiler.vbwExecuteLine 1589
    mMessage = mMessage & NewMessage
vbwProfiler.vbwProcOut 88
vbwProfiler.vbwExecuteLine 1590
End Property

Public Property Get Message() As String
vbwProfiler.vbwProcIn 89
vbwProfiler.vbwExecuteLine 1591
    Message = mMessage
vbwProfiler.vbwProcOut 89
vbwProfiler.vbwExecuteLine 1592
End Property

Public Property Let Signal(NewSignal As Long)
vbwProfiler.vbwProcIn 90
vbwProfiler.vbwExecuteLine 1593
    mSignal = NewSignal
vbwProfiler.vbwProcOut 90
vbwProfiler.vbwExecuteLine 1594
End Property

Public Property Get Signal() As Long
vbwProfiler.vbwProcIn 91
vbwProfiler.vbwExecuteLine 1595
    Signal = mSignal
vbwProfiler.vbwProcOut 91
vbwProfiler.vbwExecuteLine 1596
End Property

Public Property Let Raised(NewRaised As Boolean)
vbwProfiler.vbwProcIn 92
vbwProfiler.vbwExecuteLine 1597
If mRaised <> NewRaised Then
'MsgBox "signal " & mSignal & " Raised changing to " & NewRaised
End If
vbwProfiler.vbwExecuteLine 1598 'B
vbwProfiler.vbwExecuteLine 1599
    mRaised = NewRaised
vbwProfiler.vbwProcOut 92
vbwProfiler.vbwExecuteLine 1600
End Property

Public Property Get Raised() As Boolean
vbwProfiler.vbwProcIn 93
vbwProfiler.vbwExecuteLine 1601
    Raised = mRaised
vbwProfiler.vbwProcOut 93
vbwProfiler.vbwExecuteLine 1602
End Property

Public Property Let Silent(NewSilent As Boolean)
vbwProfiler.vbwProcIn 94
vbwProfiler.vbwExecuteLine 1603
    mSilent = NewSilent
vbwProfiler.vbwProcOut 94
vbwProfiler.vbwExecuteLine 1604
End Property

Public Property Get Silent() As Boolean
vbwProfiler.vbwProcIn 95
vbwProfiler.vbwExecuteLine 1605
    Silent = mSilent
vbwProfiler.vbwProcOut 95
vbwProfiler.vbwExecuteLine 1606
End Property

Public Property Get key() As String
vbwProfiler.vbwProcIn 96
vbwProfiler.vbwExecuteLine 1607
    key = Format$(mSecond, "00000") & ":" _
    & Format$(mSignal, "000") & ":" & mRaised
vbwProfiler.vbwProcOut 96
vbwProfiler.vbwExecuteLine 1608
End Property


