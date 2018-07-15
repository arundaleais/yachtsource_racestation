VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFlxGd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmMain 
   Caption         =   "frmMain"
   ClientHeight    =   7050
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   ScaleHeight     =   7050
   ScaleWidth      =   7920
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5520
      Top             =   5400
   End
   Begin VB.CommandButton cmdFinish 
      Caption         =   "Finish"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   25
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Frame Frame12 
      Caption         =   "Finish Times"
      Height          =   4935
      Left            =   5400
      TabIndex        =   23
      Top             =   120
      Width           =   2535
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshFinish 
         Height          =   4575
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   8070
         _Version        =   393216
         ScrollBars      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Horn"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   22
      Top             =   6000
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   2775
      Left            =   3240
      ScaleHeight     =   2715
      ScaleWidth      =   1275
      TabIndex        =   21
      Top             =   3120
      Width           =   1335
      Begin VB.Image Image3 
         Height          =   450
         Left            =   480
         Picture         =   "Form1.frx":0000
         Top             =   1320
         Width           =   450
      End
      Begin VB.Image Image2 
         Height          =   450
         Left            =   480
         Picture         =   "Form1.frx":008A
         Top             =   720
         Width           =   450
      End
      Begin VB.Image Image1 
         Height          =   450
         Left            =   480
         Picture         =   "Form1.frx":0112
         Top             =   120
         Width           =   450
      End
      Begin VB.Image Image4 
         Height          =   570
         Left            =   0
         Picture         =   "Form1.frx":019A
         Top             =   1920
         Width           =   1155
      End
   End
   Begin VB.Frame Frame11 
      Caption         =   "Starting"
      Height          =   855
      Left            =   2760
      TabIndex        =   19
      Top             =   2160
      Visible         =   0   'False
      Width           =   2175
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   120
         TabIndex        =   20
         Text            =   "Combo1"
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   2400
      TabIndex        =   14
      Top             =   120
      Width           =   2775
      Begin VB.Frame Frame3 
         Caption         =   "Elapsed Time"
         Height          =   735
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   2535
         Begin VB.Label lblElapsedTime 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "-000:09:36"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   2265
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Current Time"
         Height          =   735
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1815
         Begin VB.Label lblCurrTime 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "14:22:45"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   1500
         End
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Start"
      Height          =   1695
      Left            =   0
      TabIndex        =   10
      Top             =   4320
      Width           =   2295
      Begin VB.CommandButton Command3 
         Caption         =   "General Recall"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   12
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Recall"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   11
         Top             =   1080
         Width           =   1575
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Postponement"
      Height          =   1695
      Left            =   0
      TabIndex        =   6
      Top             =   2520
      Width           =   2295
      Begin VB.CommandButton cmdPostpone 
         Caption         =   "Postpone"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   9
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Frame Frame10 
         Caption         =   "Minutes"
         Height          =   735
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   975
         Begin VB.TextBox txtPostpone 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   8
            Text            =   "00"
            Top             =   240
            Width           =   615
         End
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Preparatory"
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   2295
      Begin VB.Frame Frame6 
         Caption         =   "First Start Time"
         Height          =   855
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   2055
         Begin VB.TextBox txtFirstStartTime 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Text            =   "1400"
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Start"
         Height          =   855
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2055
         Begin VB.OptionButton Option2 
            Caption         =   "Multiple"
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   480
            Width           =   1215
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Single"
            Height          =   255
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   1695
         End
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   6795
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type defElapsed
    Hour As Long
    Min As Long
    Sec As Long
End Type

Private FirstStartTime As Date
Private LastTimeOutput As Date
Private ElapsedTime As defElapsed


Private Sub cmdFinish_Click()
    With mshFinish
'not the fist (blank) row
        If .TextMatrix(.Rows - 1, 0) <> "" Then
            .Rows = .Rows + 1
        End If
        .TextMatrix(.Rows - 1, 0) = .Rows - 1
        .TextMatrix(.Rows - 1, 1) = lblElapsedTime.Caption
    End With

End Sub

Private Sub cmdPostpone_Click()
    txtFirstStartTime = Format$((DateAdd("n", CDbl(txtPostpone.Text), FirstStartTime)), "hhnn")
    FirstStartTime = Date & " " _
    & Format$(txtFirstStartTime, "00:00") & ":00"
End Sub

Private Sub Form_Load()
Dim i As Long
    With mshFinish
        .Width = 2295
        .FormatString = "<No|<Time"
        .ColWidth(0) = 500  'Position
        .ColWidth(1) = 2295 - 500  'Time
'        For i = 1 To 20
'            .Rows = i + 1
'            .TextMatrix(i, 0) = i
'        Next i
'        .TextMatrix(1, 1) = "13:22:45"
    End With
    Picture1.AutoRedraw = True
    FirstStartTime = Date & " " _
    & Format$(txtFirstStartTime, "00:00") & ":00"
Debug.Print Format$(FirstStartTime, "dd-mmm-yyyy")
Debug.Print Format$(FirstStartTime, "hh:mm:ss")
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer


    'close all sub forms
    For i = Forms.Count - 1 To 1 Step -1
        Unload Forms(i)
    Next
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
End Sub


'The timer mus run faster than 1 Count interval (normally 1 sec)
'Otherwise a NewCycle could be skipped if the timer is running
'slower than the actual clock time. This could happen
'if the PC is heavily loaded
Private Sub Timer1_Timer()
Dim CurrTime As Date    'My be speeded up from Now() time for testing
Dim SecDiff As Long
Dim TimeToOutput As Date
'Adjust curr time to speed up
    CurrTime = Now()
    
    If LastTimeOutput = "00:00:00" Then Call ResetOutput(CurrTime)
    
    SecDiff = DateDiff("s", LastTimeOutput, CurrTime)
'No Output due yet
    If SecDiff = 0 Then
Debug.Print "Skip"
        Exit Sub
    End If

    On Error GoTo Postpone_Error
    If txtPostpone.Text = "" Then
        txtPostpone.Text = " 0"
        txtPostpone.SelStart = 1
    End If
    If IsNumeric(txtPostpone.Text) = False Then GoTo Postpone_Error

'Check First start time has not changed
    On Error GoTo FirstStartTime_error
    If txtFirstStartTime.Text = "" Then
        txtFirstStartTime.Text = " 0"
        txtFirstStartTime.SelStart = 1
    End If
    FirstStartTime = Date & " " _
    & Format$(txtFirstStartTime, "00:00") & ":00"
    On Error GoTo 0
    Do
        StatusBar1.Panels(1).Text = Time
'        StatusBar1.Panels(2).Text = CurSecs - CycleStartSecs
        TimeToOutput = DateAdd("s", 1, LastTimeOutput)
        If TimerOutput(TimeToOutput) = True Then LastTimeOutput = TimeToOutput
        SecDiff = DateDiff("s", LastTimeOutput, CurrTime)
        lblElapsedTime = aSecToElapsed(DateDiff("s", FirstStartTime, CurrTime))
        lblCurrTime = Format$(CurrTime, "hh:mm:ss")
If SecDiff > 0 Then Debug.Print "Catch-up"
    Loop Until SecDiff = 0  'Always execute once
    Exit Sub
FirstStartTime_error:
    Call frmDpyBox.DpyBox("The First Start Time is invalid" & vbCrLf, 2, "Invalid Time")
    Resume Next
Postpone_Error:
    Call frmDpyBox.DpyBox("The Postpone Time is invalid" & vbCrLf, 2, "Invalid Time")
    Resume Next
End Sub


Private Sub ResetOutput(StartTime As Date)
        LastTimeOutput = DateAdd("s", -1, StartTime)
End Sub

Private Function aSecToElapsed(Secs As Long) As String
Dim MyElapsed As defElapsed
Dim Sign As Long
Dim aSign As String

'Secs = 3600& * 100&
    Sign = Sgn(Secs)    '-1 = -ve, 0 = 0 , +1 = +ve
    If Sign = -1 Then
        Secs = Secs * Sign 'force +ve
        aSign = "-"
    Else
        aSign = " "
    End If
    MyElapsed.Hour = Int(Secs / 3600&)
    Secs = Secs - MyElapsed.Hour * 3600&
    MyElapsed.Min = Int(Secs / 60&)
    Secs = Secs - MyElapsed.Min * 60&
    MyElapsed.Sec = Secs
    aSecToElapsed = aSign & Format$(MyElapsed.Hour, "###")
    If Abs(MyElapsed.Hour) >= 1 Then aSecToElapsed = aSecToElapsed & ":"
    aSecToElapsed = aSecToElapsed & Format$(MyElapsed.Min, "00") _
    & ":" & Format$(MyElapsed.Sec, "00")
End Function



