VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UiEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mIndex As Long
Private mSecond As Long
Private mSignal As Long
Private mEnabled As Boolean
Private mBackColor As Long
Private mFocus As Boolean

Public Property Let Second(NewSecond As Long)
vbwProfiler.vbwProcIn 93
vbwProfiler.vbwExecuteLine 1661
    mSecond = NewSecond
vbwProfiler.vbwProcOut 93
vbwProfiler.vbwExecuteLine 1662
End Property

Public Property Get Second() As Long
vbwProfiler.vbwProcIn 94
vbwProfiler.vbwExecuteLine 1663
    Second = mSecond
vbwProfiler.vbwProcOut 94
vbwProfiler.vbwExecuteLine 1664
End Property

Public Property Get Index() As Long
vbwProfiler.vbwProcIn 95
vbwProfiler.vbwExecuteLine 1665
    Index = mIndex
vbwProfiler.vbwProcOut 95
vbwProfiler.vbwExecuteLine 1666
End Property

Public Property Let Index(Item As Long)
vbwProfiler.vbwProcIn 96
vbwProfiler.vbwExecuteLine 1667
    mIndex = Item
vbwProfiler.vbwProcOut 96
vbwProfiler.vbwExecuteLine 1668
End Property

Public Property Let Signal(NewSignal As Long)
vbwProfiler.vbwProcIn 97
vbwProfiler.vbwExecuteLine 1669
    mSignal = NewSignal
vbwProfiler.vbwProcOut 97
vbwProfiler.vbwExecuteLine 1670
End Property

Public Property Get Signal() As Long
vbwProfiler.vbwProcIn 98
vbwProfiler.vbwExecuteLine 1671
    Signal = mSignal
vbwProfiler.vbwProcOut 98
vbwProfiler.vbwExecuteLine 1672
End Property

Public Property Get Enabled() As Long
vbwProfiler.vbwProcIn 99
vbwProfiler.vbwExecuteLine 1673
    Enabled = mEnabled
vbwProfiler.vbwProcOut 99
vbwProfiler.vbwExecuteLine 1674
End Property

Public Property Let Enabled(NewEnabled As Long)
vbwProfiler.vbwProcIn 100
vbwProfiler.vbwExecuteLine 1675
    mEnabled = NewEnabled
vbwProfiler.vbwProcOut 100
vbwProfiler.vbwExecuteLine 1676
End Property


Public Property Let BackColor(NewBackColor As Long)
vbwProfiler.vbwProcIn 101
vbwProfiler.vbwExecuteLine 1677
    mBackColor = NewBackColor
vbwProfiler.vbwProcOut 101
vbwProfiler.vbwExecuteLine 1678
End Property

Public Property Get BackColor() As Long
vbwProfiler.vbwProcIn 102
vbwProfiler.vbwExecuteLine 1679
    BackColor = mBackColor
vbwProfiler.vbwProcOut 102
vbwProfiler.vbwExecuteLine 1680
End Property


Public Property Let Focus(NewFocus As Long)
vbwProfiler.vbwProcIn 103
vbwProfiler.vbwExecuteLine 1681
    mFocus = NewFocus
vbwProfiler.vbwProcOut 103
vbwProfiler.vbwExecuteLine 1682
End Property

Public Property Get Focus() As Long
vbwProfiler.vbwProcIn 104
vbwProfiler.vbwExecuteLine 1683
    Focus = mFocus
vbwProfiler.vbwProcOut 104
vbwProfiler.vbwExecuteLine 1684
End Property

Public Property Get key() As String
vbwProfiler.vbwProcIn 105
vbwProfiler.vbwExecuteLine 1685
    key = Format$(mSecond, "00000") & ":" _
    & Format$(mSignal, "000")
vbwProfiler.vbwProcOut 105
vbwProfiler.vbwExecuteLine 1686
End Property



