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
vbwProfiler.vbwProcIn 124
vbwProfiler.vbwExecuteLine 1887
    mSecond = NewSecond
vbwProfiler.vbwProcOut 124
vbwProfiler.vbwExecuteLine 1888
End Property

Public Property Get Second() As Long
vbwProfiler.vbwProcIn 125
vbwProfiler.vbwExecuteLine 1889
    Second = mSecond
vbwProfiler.vbwProcOut 125
vbwProfiler.vbwExecuteLine 1890
End Property

Public Property Get Index() As Long
vbwProfiler.vbwProcIn 126
vbwProfiler.vbwExecuteLine 1891
    Index = mIndex
vbwProfiler.vbwProcOut 126
vbwProfiler.vbwExecuteLine 1892
End Property

Public Property Let Index(Item As Long)
vbwProfiler.vbwProcIn 127
vbwProfiler.vbwExecuteLine 1893
    mIndex = Item
vbwProfiler.vbwProcOut 127
vbwProfiler.vbwExecuteLine 1894
End Property

Public Property Let Signal(NewSignal As Long)
vbwProfiler.vbwProcIn 128
vbwProfiler.vbwExecuteLine 1895
    mSignal = NewSignal
vbwProfiler.vbwProcOut 128
vbwProfiler.vbwExecuteLine 1896
End Property

Public Property Get Signal() As Long
vbwProfiler.vbwProcIn 129
vbwProfiler.vbwExecuteLine 1897
    Signal = mSignal
vbwProfiler.vbwProcOut 129
vbwProfiler.vbwExecuteLine 1898
End Property

Public Property Get Enabled() As Long
vbwProfiler.vbwProcIn 130
vbwProfiler.vbwExecuteLine 1899
    Enabled = mEnabled
vbwProfiler.vbwProcOut 130
vbwProfiler.vbwExecuteLine 1900
End Property

Public Property Let Enabled(NewEnabled As Long)
vbwProfiler.vbwProcIn 131
vbwProfiler.vbwExecuteLine 1901
    mEnabled = NewEnabled
vbwProfiler.vbwProcOut 131
vbwProfiler.vbwExecuteLine 1902
End Property


Public Property Let BackColor(NewBackColor As Long)
vbwProfiler.vbwProcIn 132
vbwProfiler.vbwExecuteLine 1903
    mBackColor = NewBackColor
vbwProfiler.vbwProcOut 132
vbwProfiler.vbwExecuteLine 1904
End Property

Public Property Get BackColor() As Long
vbwProfiler.vbwProcIn 133
vbwProfiler.vbwExecuteLine 1905
    BackColor = mBackColor
vbwProfiler.vbwProcOut 133
vbwProfiler.vbwExecuteLine 1906
End Property


Public Property Let Focus(NewFocus As Long)
vbwProfiler.vbwProcIn 134
vbwProfiler.vbwExecuteLine 1907
    mFocus = NewFocus
vbwProfiler.vbwProcOut 134
vbwProfiler.vbwExecuteLine 1908
End Property

Public Property Get Focus() As Long
vbwProfiler.vbwProcIn 135
vbwProfiler.vbwExecuteLine 1909
    Focus = mFocus
vbwProfiler.vbwProcOut 135
vbwProfiler.vbwExecuteLine 1910
End Property

Public Property Get key() As String
vbwProfiler.vbwProcIn 136
vbwProfiler.vbwExecuteLine 1911
    key = Format$(mSecond, "00000") & ":" _
    & Format$(mSignal, "000")
vbwProfiler.vbwProcOut 136
vbwProfiler.vbwExecuteLine 1912
End Property



