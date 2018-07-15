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

Private mSecond As Long
Private mClass As Long
Private mSignal As Long
Private mState As Boolean

Public Property Let Second(NewSecond As Long)
    mSecond = NewSecond
End Property

Public Property Get Second() As Long
    Second = mSecond
End Property

Public Property Let Class(NewClass As Long)
    mClass = NewClass
End Property

Public Property Get Class() As Long
    Class = mClass
End Property

Public Property Let Signal(NewSignal As Long)
    mSignal = NewSignal
End Property

Public Property Get Signal() As Long
    Signal = mSignal
End Property

Public Property Let State(NewState As Boolean)
    mState = NewState
End Property

Public Property Get State() As Boolean
    State = mState
End Property

Public Property Get Key() As String
    Key = Format$(mSecond, "00000") & ":" _
    & Format$(mClass, "000") & ":" _
    & Format$(mSignal, "000")
End Property



