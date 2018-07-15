VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public Class As Long
Public Elapsed As Long        'Seconds before -ve, after first start time
Public Signal As Long
Public OnOff As Boolean
'Public ElapsedSeconds As Collection 'Key is Elapsed

Public Property Get Key() As String
    Key = Format$(Elapsed, "00000") & ":" _
    & Format$(Class, "000") & ":" _
    & Format$(Signal, "000")
End Property

