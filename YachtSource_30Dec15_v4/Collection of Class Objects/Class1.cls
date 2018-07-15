VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Collection2"
Option Explicit
Public Name As String
'local variable(s) to hold property value(s)
Private mvarProfile As Object 'local copy
Private mvarCollection2 As Collection2


Public Property Get Collection2() As Collection2
    If mvarCollection2 Is Nothing Then
        Set mvarCollection2 = New Collection2
    End If


    Set Collection2 = mvarCollection2
End Property


Public Property Set Collection2(vData As Collection2)
    Set mvarCollection2 = vData
End Property

Private Sub Class_Initialize()

End Sub

Private Sub Class_Terminate()
    Set mvarCollection2 = Nothing
End Sub



Public Property Set Profile(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Profile = Form1
    Set mvarProfile = vData
End Property


Public Property Get Profile() As Object
Attribute Profile.VB_UserMemId = 0
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Profile
    Set Profile = mvarProfile
End Property



