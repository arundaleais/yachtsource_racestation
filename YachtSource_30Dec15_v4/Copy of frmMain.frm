VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFlxGd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.ocx"
Begin VB.Form frmMain 
   Caption         =   "frmMain"
   ClientHeight    =   6825
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   ScaleHeight     =   6825
   ScaleWidth      =   8175
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraMain 
      Height          =   4455
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   6975
      Begin VB.Frame Frame12 
         Caption         =   "Finish Times"
         Height          =   4215
         Left            =   4920
         TabIndex        =   21
         Top             =   120
         Width           =   2055
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshFinish 
            Height          =   3855
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   6800
            _Version        =   393216
            ScrollTrack     =   -1  'True
            ScrollBars      =   2
            Appearance      =   0
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
      Begin VB.Frame Frame1 
         Caption         =   "Finish"
         Height          =   855
         Left            =   2280
         TabIndex        =   19
         Top             =   3480
         Width           =   2655
         Begin VB.CommandButton Commands 
            Caption         =   "Commands"
            Height          =   375
            Index           =   4
            Left            =   720
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame fraStart 
         Caption         =   "Start"
         Height          =   1935
         Left            =   2280
         TabIndex        =   17
         Top             =   1680
         Width           =   2655
         Begin VB.CommandButton Commands 
            Caption         =   "Commands"
            Height          =   375
            Index           =   3
            Left            =   600
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   1080
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Current Time"
         Height          =   975
         Left            =   2280
         TabIndex        =   15
         Top             =   720
         Width           =   2655
         Begin VB.Label lblCurrTime 
            AutoSize        =   -1  'True
            BackColor       =   &H80000014&
            Caption         =   "14:22:45"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   2460
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Elapsed Time"
         Height          =   615
         Left            =   2280
         TabIndex        =   13
         Top             =   120
         Width           =   2655
         Begin VB.Label lblElapsedTime 
            AutoSize        =   -1  'True
            BackColor       =   &H80000014&
            Caption         =   "-000:09:36"
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
            Left            =   360
            TabIndex        =   14
            Top             =   240
            Width           =   1845
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Horn"
         Height          =   855
         Left            =   0
         TabIndex        =   11
         Top             =   3480
         Width           =   2295
         Begin VB.CommandButton Commands 
            Caption         =   "Commands"
            Height          =   375
            Index           =   1
            Left            =   600
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Postponement"
         Height          =   1575
         Left            =   0
         TabIndex        =   7
         Top             =   2040
         Width           =   2295
         Begin VB.Frame Frame10 
            Caption         =   "Minutes"
            Height          =   735
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   1215
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
               TabIndex        =   10
               Text            =   "15"
               Top             =   240
               Width           =   855
            End
         End
         Begin VB.CommandButton Commands 
            Caption         =   "Commands"
            Height          =   375
            Index           =   2
            Left            =   600
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   1080
            Width           =   1095
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Preparatory"
         Height          =   1935
         Left            =   0
         TabIndex        =   4
         Top             =   120
         Width           =   2295
         Begin VB.Frame Frame5 
            Caption         =   "Start Sequence"
            Height          =   735
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   2055
            Begin VB.ComboBox cboProfile 
               Height          =   315
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   24
               Top             =   240
               Width           =   1815
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "First Start Time"
            Height          =   735
            Left            =   120
            TabIndex        =   5
            Top             =   1080
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
               TabIndex        =   6
               Text            =   "1400"
               Top             =   240
               Width           =   1095
            End
         End
      End
   End
   Begin VB.CommandButton Commands 
      Caption         =   "Commands"
      Height          =   375
      Index           =   0
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   600
      ScaleHeight     =   1575
      ScaleWidth      =   6135
      TabIndex        =   1
      Tag             =   "2"
      Top             =   0
      Width           =   6135
      Begin VB.Image Flags 
         Height          =   375
         Index           =   40
         Left            =   5400
         Top             =   1080
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   39
         Left            =   4800
         Top             =   1080
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   38
         Left            =   4200
         Top             =   1080
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   37
         Left            =   3600
         Top             =   1080
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   36
         Left            =   3000
         Top             =   1080
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   35
         Left            =   2400
         Top             =   1080
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   34
         Left            =   1800
         Top             =   1080
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   33
         Left            =   1200
         Top             =   1080
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   32
         Left            =   600
         Top             =   1080
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   31
         Left            =   0
         Top             =   1080
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   30
         Left            =   5400
         Top             =   720
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   29
         Left            =   4800
         Top             =   720
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   28
         Left            =   4200
         Top             =   720
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   27
         Left            =   3600
         Top             =   720
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   26
         Left            =   3000
         Top             =   720
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   25
         Left            =   2400
         Top             =   720
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   24
         Left            =   1800
         Top             =   720
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   23
         Left            =   1200
         Top             =   720
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   22
         Left            =   600
         Top             =   720
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   21
         Left            =   0
         Top             =   720
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   20
         Left            =   5400
         Top             =   360
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   19
         Left            =   4800
         Top             =   360
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   18
         Left            =   4200
         Top             =   360
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   17
         Left            =   3600
         Top             =   360
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   16
         Left            =   3000
         Top             =   360
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   15
         Left            =   2400
         Top             =   360
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   14
         Left            =   1800
         Top             =   360
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   13
         Left            =   1200
         Top             =   360
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   12
         Left            =   600
         Top             =   360
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   11
         Left            =   0
         Top             =   360
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   10
         Left            =   5400
         Top             =   0
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   9
         Left            =   4800
         Top             =   0
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   8
         Left            =   4200
         Top             =   0
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   7
         Left            =   3600
         Top             =   0
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   6
         Left            =   3000
         Top             =   0
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   5
         Left            =   2400
         Top             =   0
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   4
         Left            =   1800
         Top             =   0
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   3
         Left            =   1200
         Top             =   0
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   2
         Left            =   600
         Top             =   0
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   1
         Left            =   0
         Top             =   0
         Width           =   615
      End
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   6120
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer ReloadTimer 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6960
      Top             =   0
   End
   Begin VB.Timer SignalTimer 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   500
      Left            =   6600
      Top             =   360
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   6240
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer RaceTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6600
      Top             =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   6570
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
   End
   Begin VB.Image Flags 
      Height          =   375
      Index           =   0
      Left            =   6600
      Top             =   600
      Width           =   615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type defhms
    Hour As Long
    Min As Long
    Sec As Long
End Type
Private FirstStartTime As Date
Private LastTimeOutput As Date  'Used for catch-up
Private NextCommandTop As Single

Private Sub cboProfile_Click()
'Because you get an Unable to unload within this context (Error 365)
'if you try and remove the SignalTimerControls from within
'a Combo Click event we set a flag and do it using timer
'see http://msdn.microsoft.com/en-us/library/aa243662%28v=vs.60%29.aspx
    ReloadTimer.Enabled = True
End Sub

Private Sub LoadSequence()
Dim i As Long

'Load startsequencies - only on initial startup
    With cboProfile
'Only load if no files already loaded
        If .ListCount = 0 Then
            IniFileName = Dir$(Environ("AllUsersProfile") & "\Application Data\Arundale\RacingSignals\Sequences\*.ini")
            Do While IniFileName > ""
'Dont allow *.ini_old
                If Right$(IniFileName, 4) = ".ini" Then
'Remove .ini so it's not displayed
                    i = InStrRev(IniFileName, ".ini")
                    If i > 0 Then
                        IniFileName = Left$(IniFileName, i - 1)
                        .AddItem IniFileName
                    End If
                End If
                IniFileName = Dir$
            Loop
'If none exit sub & program (in .main)
            If .ListCount = 0 Then
                MsgBox "No Start Sequences available" & vbCrLf & "Exiting Program", vbCritical, "No Start Sequence"
                Exit Sub
            End If
        End If
'If No Sequence selected
        If .ListIndex = -1 Then
            MsgBox "Please select a Start Sequence", vbExclamation, "No Start Sequence"
        Else
'When a profile is first selected dont use the reload timer
            Call LoadProfile
'Dont reset Postpone unless loading the profile
'because once racing has been postpned the Signal stay up until
'all starts have been concluded
'    cmdPostpone.BackColor = vbGreen
        End If
    End With
End Sub
Private Sub cmdFinish_Click()
    Call MakeSignals(SignalIdx("Flag", "Finish"), True)
    Call MakeSignals(SignalIdx("Flag", "Horn1Short"), True)
    With mshFinish
'not the first (blank) row
        If .TextMatrix(.Rows - 1, 0) <> "" Then
            .Rows = .Rows + 1
        End If
        .TextMatrix(.Rows - 1, 0) = .Rows - 1
        .TextMatrix(.Rows - 1, 1) = lblCurrTime.Caption
'Scroll to bottom
        .TopRow = .Rows - 1
End With
Debug.Print "Finish"

End Sub

Private Sub cmdHorn_Click()
    
    Call MakeSignals(SignalIdx("Flag", "Horn"), True)
Debug.Print "Horn"
End Sub

Private Sub cmdPostpone_Click()
    
    cmdPostpone.BackColor = vbRed
    
    If ValidatePostponeTime = True Then
'Causes StartTime to be validated
'Which then causes Events to be reloaded
        txtFirstStartTime = Format$((DateAdd("n", CDbl(NulToZero(txtPostpone.Text)), FirstStartTime)), "hhnn")
    End If
End Sub

Private Sub cmdRecall_Click()
    Select Case cmdRecall.BackColor
'Default needs removing a few secs  after start
    Case Is = vbGreen
        cmdRecall.BackColor = vbRed
    Case Is = vbRed
        If SignalTimer(SignalIdx("Recall")).Enabled = True Then
            cmdRecall.BackColor = cbDefault
'Must do all 3 lines to ensure RecallSignal is turned off
            Call MakeSignals(SignalIdx("Recall"), False)
            SignalTimer(SignalIdx("Recall")).Enabled = False
'Do this last, so if the timer is called again
'another off will be generated, and the timer will
'not re-start
            SignalAttributes(SignalIdx("Recall")).CyclesCompleted = 0
            cmdFinish.BackColor = vbGreen
            cmdFinish.SetFocus
        End If
    Case Else
'Dont action
    End Select
Debug.Print "Recall"

End Sub

Private Sub Commands_Click(Index As Integer)
Dim Position As Long

    If IsFlagUp(Index) = False Then
        Call PutFlagUp(Index)
    Else
        Call PutFlagDown(Index)
    End If
    Call CommandColor
End Sub

Private Sub Form_Load()
Dim i As Long
Dim url As String
Dim Major As Long
Dim Minor As Long
Dim Revision As Long
Dim NewVersion As Boolean

    Caption = App.EXEName & " [" & App.Major & "." & App.Minor & "." _
    & App.Revision & "] "

'Check if a later version exists
    url = "http://www.NmeaRouter.com/docs/ais/" & App.EXEName _
    & "_setup_"
    Major = App.Major
    Do
        If HTTPFileExists(url & Major & ".0.0.exe") = False Then Exit Do
        Major = Major + 1
    Loop
    If Major > 0 Then Major = Major - 1   'Highest major that exists
    
    url = url & Major & "."
    If Major = App.Major Then
        Minor = App.Minor
    Else
        Minor = 0
    End If
    Do
        If HTTPFileExists(url & Minor & ".0.exe") = False Then Exit Do
        Minor = Minor + 1
    Loop
    If Minor > 0 Then Minor = Minor - 1

    url = url & Minor & "."
    If Not (Major = App.Major And Minor = App.Minor) Then
        NewVersion = True
    End If
'Only let a user get next revision if he is using a revision
'of his current version. Otherwise he goes up to the next minor version
    If NewVersion = False And App.Revision > 0 Then
        Revision = App.Revision
        Do
            If HTTPFileExists(url & Revision & ".exe") = False Then Exit Do
            Revision = Revision + 1
        Loop
        If Revision > 0 Then Revision = Revision - 1
        If Revision < App.Revision Then
            NewVersion = True
        End If
    End If
    url = url & Revision & ".exe"
    
'If we are working on a higher version in VBE, don't try for newversion
    If App.Major * 2 ^ 8 + App.Minor * 2 ^ 4 + App.Revision > _
    Major * 2 ^ 8 + Minor * 2 ^ 4 + Revision Then
        NewVersion = False
    End If
    If NewVersion = True Then
        Call frmDpyBox.DpyBox("A new update is available", 10, "New Version")
'Check we have internet access
        If HTTPFileExists(url) Then
            Call HttpSpawn(url)
        End If
    End If
'Position cursor at RHS of time displayed
    txtFirstStartTime.SelStart = Len(txtFirstStartTime)
    txtPostpone.SelStart = Len(txtPostpone)
    With mshFinish
        .Width = 1795
        .FormatString = "<No|<Time"
        .ColWidth(0) = 500  'Position
        .ColWidth(1) = 1295  'Time
'        For i = 1 To 20
'            .Rows = i + 1
'            .TextMatrix(i, 0) = i
'        Next i
'        .TextMatrix(1, 1) = "13:22:45"
    End With
' Debug

Visible = True
'Make the base index invisible as it is not used
    Commands(0).Enabled = False
    Commands(0).Visible = False
'Set up initial start time, LoadEvents not called
    FirstStartTime = Date & " " _
    & Format$(NulToZero(txtFirstStartTime), "00:00") & ":00"
Debug.Print Format$(FirstStartTime, "dd-mmm-yyyy")
Debug.Print Format$(FirstStartTime, "hh:mm:ss")

    Call LoadSequence

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





'The timer must run faster than 1 Count interval (normally 1 sec)
'Otherwise a NewCycle could be skipped if the timer is running
'slower than the actual clock time. This could happen
'if the PC is heavily loaded
Private Sub RaceTimer_Timer()
Dim CurrTime As Date    'May be speeded up from Now() time for testing
Dim SecsSinceOutput As Long
Dim TimeToOutput As Date
'Dim MyFirstStartTime As Date

'Adjust curr time to speed up
    CurrTime = Now()
    
    If LastTimeOutput = "00:00:00" Then Call ResetOutput(CurrTime)
    
    SecsSinceOutput = DateDiff("s", LastTimeOutput, CurrTime)
'No Output due yet
    If SecsSinceOutput = 0 Then
Debug.Print "Skip"
        Exit Sub
    End If
    
    Do
'        StatusBar1.Panels(1).Text = Time
'        StatusBar1.Panels(2).Text = CurSecs - CycleStartSecs
        TimeToOutput = DateAdd("s", 1, LastTimeOutput)
        If TimerOutput(TimeToOutput) = True Then LastTimeOutput = TimeToOutput
        SecsSinceOutput = DateDiff("s", LastTimeOutput, CurrTime)
        ElapsedTime = DateDiff("s", FirstStartTime, CurrTime)
Debug.Print ElapsedTime - SecsSinceOutput
        Call DoTimerEvents(ElapsedTime - SecsSinceOutput)
        lblElapsedTime = aSecToElapsed(ElapsedTime)
'        lblElapsedTime = aSecToElapsed(DateDiff("s", FirstStartTime, CurrTime))
        lblCurrTime = Format$(CurrTime, "hh:mm:ss")
If SecsSinceOutput > 0 Then Debug.Print "Catch-up " & SecsSinceOutput
    Loop Until SecsSinceOutput = 0  'Always execute once
End Sub


Private Sub ResetOutput(StartTime As Date)
        LastTimeOutput = DateAdd("s", -1, StartTime)
End Sub

Private Function aSecToElapsed(Secs As Long) As String
Dim hms As defhms
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
    hms.Hour = Int(Secs / 3600&)
    Secs = Secs - hms.Hour * 3600&
    hms.Min = Int(Secs / 60&)
    Secs = Secs - hms.Min * 60&
    hms.Sec = Secs
    aSecToElapsed = aSign & Format$(hms.Hour, "###")
    If Abs(hms.Hour) >= 1 Then aSecToElapsed = aSecToElapsed & ":"
    aSecToElapsed = aSecToElapsed & Format$(hms.Min, "00") _
    & ":" & Format$(hms.Sec, "00")
End Function

Private Sub ReloadTimer_Timer()
    ReloadTimer.Enabled = False
    Call LoadProfile
End Sub

Private Sub SignalTimer_Timer(Index As Integer)
'A cycle is completed every time a flag
'is turned off
    If IsFlagUp(Index) = True Then
        SignalAttributes(Index).CyclesCompleted = SignalAttributes(Index).CyclesCompleted + 1
    End If
    Select Case SignalAttributes(Index).CyclesCompleted
'The timer has started but we do not want the Signal Off
'In fact we should not have started it in the first place
    Case Is >= SignalAttributes(Index).CyclesRequired
'Turn off Signal, before disabling the timer
'Otherwise MakeSignals will start it again
        Call MakeSignals(CLng(Index), False)
        SignalTimer(Index).Enabled = False
'Do this last, so if the timer is called again
'another off will be generated, and the timer will
'not re-start
        SignalAttributes(Index).CyclesCompleted = 0
        If FlagIndex = RecallIdx Then
            cmdRecall.BackColor = cbDefault
            cmdFinish.BackColor = vbGreen
            cmdFinish.SetFocus
        End If
        
    Case Is < SignalAttributes(Index).CyclesRequired
'Reverse the Signal and do another cycle
        Call MakeSignals(CLng(Index), NFlags(FlagIndex).Picture.Handle = 0)
'Keep the timer running
    End Select
    
End Sub

Private Sub txtFirstStartTime_Change()
    If ValidateStartTime = True Then
        Call ResetCmd
    End If
End Sub

Private Sub txtPostpone_Change()
    Call ValidatePostponeTime
End Sub

Private Function ValidateStartTime() As Boolean

    On Error GoTo ValidateStartTime_error
    If txtFirstStartTime = "" Then
        txtFirstStartTime.BackColor = vbRed
    Else
        txtFirstStartTime.BackColor = vbWhite
    End If
    If Len(txtFirstStartTime) = 4 _
    And CLng(NulToZero(txtFirstStartTime)) >= 0 _
    And CLng(NulToZero(txtFirstStartTime)) <= 2400 _
    And IsNumeric(NulToZero(txtFirstStartTime)) = True Then
        FirstStartTime = Date & " " _
        & Format$(NulToZero(txtFirstStartTime), "00:00") & ":00"
        On Error GoTo 0
        txtFirstStartTime.ForeColor = vbBlack
'Must not only reset the flags because once the start sequence
'has commenced the whole profile should be reloaded
'        Call ResetFlags
        ValidateStartTime = True
        Exit Function
    End If
ValidateStartTime_error:
    txtFirstStartTime.ForeColor = vbRed
End Function

Private Function ValidatePostponeTime() As Boolean

    On Error GoTo ValidatePostponeTime_error
    If txtPostpone = "" Then
        txtPostpone.BackColor = vbRed
    Else
        txtPostpone.BackColor = vbWhite
    End If
    If IsNumeric(NulToZero(txtPostpone)) = True Then
        txtPostpone.ForeColor = vbBlack
        ValidatePostponeTime = True
        Exit Function
    End If
ValidatePostponeTime_error:
    txtPostpone.ForeColor = vbRed
End Function

Public Function ResetFlags()
Dim MyImage As Image

Debug.Print "ResetFlags"
    For Each MyImage In frmMain.Flags
        MyImage.Picture = Nothing
    Next
End Function

Public Function ResetCommands()
Dim MyCommand As CommandButton
    For Each MyCommand In Commands
        If MyCommand.Index <> 0 Then
            MyCommand.Enabled = True
            MyCommand.Visible = True
        End If
    Next MyCommand
End Function
Public Function ResetSignalTimers()
Dim MySignalTimer As Timer

    For Each MySignalTimer In frmMain.SignalTimer
        If MySignalTimer.Index > 0 Then  'Dont delete SignalTimer(0)
            Unload MySignalTimer
        End If
    Next
End Function

Public Function ResetFinish()
Dim Row As Long
Dim Col As Long

    With mshFinish
'Clear rows (except 1)
        For Row = 2 To .Rows - 1
            .RemoveItem 1
        Next Row
'Clear Row 1
        For Col = 0 To .Cols - 1
            .TextMatrix(1, Col) = ""
        Next Col
    End With
End Function

Public Function ResetCmd()
'Must have Property Style set to 1=Graphical
'    cmdPostpone.Enabled = True
'Recall Signal says up
'    cmdRecall.BackColor = cbDefault
    cmdFinish.BackColor = cbDefault
'    cmdHorn.BackColor = cbDefault
'    cmdPostpone.SetFocus
End Function

'Requires MSINET.OCX
'See http://officeone.mvps.org/vba/http_file_exists.html
Public Function HTTPFileExists(ByVal url As String) As Boolean
    Dim S As String
    Dim Exists As Boolean
    On Error GoTo Inet1_Error
    With Inet1
        .RequestTimeout = 20
        .Protocol = icHTTP
        .url = url
        .Execute
'see http://support.microsoft.com/kb/182152 =True doesnt work
        Do While .StillExecuting <> False
            DoEvents
        Loop
        S = UCase(.GetHeader())
        Exists = (InStr(1, S, "200 OK") > 0)
    End With
    HTTPFileExists = Exists
    Exit Function
Inet1_Error:
    Select Case Err.Number
    Case Is = 35764 '
    End Select
    
End Function

Public Function HttpSpawn(url As String)
Dim r As Long
Dim Command As String

If Environ("windir") <> "" Then
    r = ShellExecute(0, "open", url, 0, 0, 1)
Else
'try for linux compatibility
    Command = "winebrowser " & url & " ""%1"""

    Shell (Command)
End If
End Function

Public Function LoadCommand(Idx As Long)
'You dont need these unless testing this module in VBE
'If you have a break set frmMain is minimised and
'the Scale values will be 0
    Load Commands(Idx)
    With Commands(Idx)
        .Caption = .Caption & "(" & Idx & ")"
'This will be overwritten with the Name from SignalAttributes
        .Visible = True
        .Enabled = True
'Align first command with top of main frame
WindowState = vbNormal  'Scale will be 0 in VBE (window is minimized)
        .Top = ScaleTop + fraMain.Top + NextCommandTop
        If .Top + .Height > fraMain.Top + fraMain.Height Then
            NextCommandTop = 0
            Width = Width + .Width
WindowState = vbNormal  'Scale will be 0 in VBE (window is minimized)
            .Top = ScaleTop + fraMain.Top + NextCommandTop
        End If
WindowState = vbNormal  'Scale will be 0 in VBE (window is minimized)
        .Left = ScaleWidth - .Width
        NextCommandTop = NextCommandTop + .Height
    End With
End Function

Private Function CommandColor()
Dim MyCommand As CommandButton

    For Each MyCommand In Commands
        If MyCommand.Index > 0 Then 'skip command(0)
            If IsFlagUp(MyCommand.Index) = False Then
                    MyCommand.BackColor = cbDefault
            Else
                MyCommand.BackColor = vbGreen
            End If
        End If
    Next MyCommand
    
End Function

'The Flags(idx) may be allocated to a flag but not visible
Public Function IsFlagUp(Idx) As Boolean
Dim MyFlag As Image
    
    If Idx > 0 Then 'Return false if Idx=0 - could be no parent
        For Each MyFlag In Flags
'Check signalAttributes have been created
            If Idx <= UBound(SignalAttributes) Then
'Check Image has been loaded
                If Not SignalAttributes(Idx).Image Is Nothing Then
                    If MyFlag.Picture = SignalAttributes(Idx).Image Then    'Skip Flags(0)
                        IsFlagUp = True
                        Exit For
                    End If
                End If
            End If
        Next MyFlag
    End If
End Function

'Returns 0 if cannot find this flag
Public Function FlagPosition(Idx) As Long
Dim MyFlag As Image
    
    If Idx > 0 Then 'Return false if Idx=0 - could be no parent
        For Each MyFlag In Flags
'Check signalAttributes have been created
            If Idx <= UBound(SignalAttributes) Then
'Check Image has been loaded
                If Not SignalAttributes(Idx).Image Is Nothing Then
                    If MyFlag.Picture = SignalAttributes(Idx).Image Then    'Skip Flags(0)
                        FlagPosition = MyFlag.Index
                        Exit For
                    End If
                End If
            End If
        Next MyFlag
    End If
End Function

Public Function PutFlagUp(Idx)
Dim MyFlag As Image
Dim ParentPosition As Long

'If we do not have a position yet see if this flag has a parent
'ie a 2 flag hoist and the paren flag is up
        If SignalAttributes(Idx).Position > 0 Then
            Set MyFlag = Flags(SignalAttributes(Idx).Position)
        Else
'If we have a parent consider placing under parent
            ParentPosition = FlagPosition(SignalAttributes(Idx).Parent)
 'If no Parent or Parent is not positioned then position=0
            If ParentPosition > 0 Then
'If there is a parent at the top
                If ParentPosition <= 10 Then
                    Set MyFlag = Flags(ParentPosition + 10)
                Else
'Place above parent
                    Set MyFlag = Flags(ParentPosition - 10)
                End If
            Else
'If no parent Positioned get first empty position
                For Each MyFlag In Flags
                    If MyFlag.Index > 0 Then    'Skip Flags(0)
                        If MyFlag.Picture.Handle = 0 Then
                            Exit For
                        End If
                    End If
                Next MyFlag
            End If
        End If
        If Not MyFlag Is Nothing Then
            Set MyFlag.Picture = SignalAttributes(Idx).Image
        Else
MsgBox "No free Flag Positions", vbExclamation, "PutFlagUp"
        End If
End Function

Public Function PutFlagDown(Idx)
    Set Flags(FlagPosition(Idx)).Picture = Nothing
End Function

