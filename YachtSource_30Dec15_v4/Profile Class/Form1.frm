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
   Begin VB.CommandButton Command3 
      Caption         =   "NextEvent"
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "New Event"
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim myProfile As clsProfile
Dim ElapsedSecond As Long
Private Sub Command1_Click()
    myProfile.Name = "Profile Name"
Debug.Print myProfile.Name
End Sub

Private Sub Command2_Click()
Dim myClass As Long
    For myClass = 1 To 3
        myProfile.NewEvent ElapsedSecond, myClass, 1, False
    Next myClass
    ElapsedSecond = ElapsedSecond + 1
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Form_Load()
Debug.Print "---Load---"
    Set myProfile = New clsProfile
    ElapsedSecond = -5
End Sub

Private Sub Form_Unload(Cancel As Integer)
Debug.Print "---UnLoad---"
    Set myProfile = Nothing

End Sub
