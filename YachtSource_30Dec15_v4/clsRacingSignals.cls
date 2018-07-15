VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRacingSignals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Racing Signals"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Postponement"
Attribute VB_Ext_KEY = "Member1" ,"Preparatory"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarPostponement As Postponement
Private mvarPreparatory As Preparatory


Public Property Get Preparatory() As Preparatory
    If mvarPreparatory Is Nothing Then
        Set mvarPreparatory = New Preparatory
    End If


    Set Preparatory = mvarPreparatory
End Property


Public Property Set Preparatory(vData As Preparatory)
    Set mvarPreparatory = vData
End Property



Public Property Get Postponement() As Postponement
    If mvarPostponement Is Nothing Then
        Set mvarPostponement = New Postponement
    End If


    Set Postponement = mvarPostponement
End Property


Public Property Set Postponement(vData As Postponement)
    Set mvarPostponement = vData
End Property
Private Sub Class_Terminate()
  Set mvarPreparatory = Nothing
    Set mvarPostponement = Nothing
End Sub



