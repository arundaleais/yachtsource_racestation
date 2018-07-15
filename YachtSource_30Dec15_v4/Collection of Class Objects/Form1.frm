VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim myClass1 As New Class1              'Profile
Dim myCollection1 As New Collection1    'Profiles
Dim myClass2 As New Class2              'Second
Dim myCollection2 As New Collection2    'Seconds
Dim myClass3 As New Class3              'Action
Dim myCollection3 As New Collection3    'Actions

Private Sub Command1_Click()
    myCollection1.Name = "Profiles"
    myCollection1.Add ("Profile1")
'Must set default method as Item Programmers Guide p441
'Tools > Procedure Attributes
    Set myClass1 = myCollection1("Profile1")
    myClass1.Name = "Profile1 Name"
'MyClass1 is the current profile
    myClass1.Collection2.Name = "Seconds"
    myClass1.Collection2.Add ("Second1")
    
    Set myClass2 = myClass1.Collection2("Second1")
    myClass2.Name = "Second1"
    myClass2.Collection3.Name = "Seconds"
    myClass2.Collection3.Add ("Action1")
    
    Set myClass3 = myClass2.Collection3("Action1")
    myClass3.Name = "Action1"
    
    Set myCollection1 = Nothing
End Sub
