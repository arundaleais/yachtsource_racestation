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
    mSecond = NewSecond
End Property

Public Property Get Second() As Long
    Second = mSecond
End Property

Public Property Get Index() As Long
    Index = mIndex
End Property

Public Property Let Index(Item As Long)
    mIndex = Item
End Property

Public Property Let Signal(NewSignal As Long)
    mSignal = NewSignal
End Property

Public Property Get Signal() As Long
    Signal = mSignal
End Property

Public Property Get Enabled() As Long
    Enabled = mEnabled
End Property

Public Property Let Enabled(NewEnabled As Long)
    mEnabled = NewEnabled
End Property


Public Property Let BackColor(NewBackColor As Long)
    mBackColor = NewBackColor
End Property

Public Property Get BackColor() As Long
    BackColor = mBackColor
End Property


Public Property Let Focus(NewFocus As Long)
    mFocus = NewFocus
End Property

Public Property Get Focus() As Long
    Focus = mFocus
End Property

Public Property Get key() As String
    key = Format$(mSecond, "00000") & ":" _
    & Format$(mSignal, "000")
End Property


