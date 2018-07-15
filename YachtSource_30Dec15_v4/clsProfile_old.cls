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
'Container to hold the Collection
Public Name As String
Public Profile As Collection

'Public Property Get Key() As String
'    Key = Format$(Elapsed, "00000") & ":" _
'    & Format$(Class, "000") & ":" _
'    & Format$(Signal, "000")
'End Property

Private Sub Class_Initialize()
    Set Profile = New Collection
End Sub

'Public Function Add(Profile As clsProfile, Optional Key As String)
'oProfiles.Add "Profile1"
'End Function
