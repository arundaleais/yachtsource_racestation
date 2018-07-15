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

Public Property Let Message(NewMessage As String)
'Combine message is separate in .ini file
    If mMessage <> "" Then mMessage = mMessage & ", "
    mMessage = mMessage & NewMessage
End Property

Public Property Get Message() As String
    Message = mMessage
End Property

Public Property Let Signal(NewSignal As Long)
    mSignal = NewSignal
End Property

Public Property Get Signal() As Long
    Signal = mSignal
End Property

Public Property Let Raised(NewRaised As Boolean)
If mRaised <> NewRaised Then
'MsgBox "signal " & mSignal & " Raised changing to " & NewRaised
End If
    mRaised = NewRaised
End Property

Public Property Get Raised() As Boolean
    Raised = mRaised
End Property

Public Property Let Silent(NewSilent As Boolean)
    mSilent = NewSilent
End Property

Public Property Get Silent() As Boolean
    Silent = mSilent
End Property

Public Property Get key() As String
    key = Format$(mSecond, "00000") & ":" _
    & Format$(mSignal, "000") & ":" & mRaised
End Property

