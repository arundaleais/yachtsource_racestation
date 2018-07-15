VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFlxGd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.ocx"
Begin VB.Form frmMain 
   Caption         =   "frmMain"
   ClientHeight    =   6345
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   ScaleHeight     =   6345
   ScaleWidth      =   8205
   StartUpPosition =   3  'Windows Default
   WindowState     =   1  'Minimized
   Begin VB.CommandButton Commands 
      Caption         =   "Commands"
      Height          =   375
      Index           =   10
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   360
      Width           =   1095
   End
   Begin VB.Timer ClearFlagsTimer 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   7320
      Top             =   480
   End
   Begin VB.Timer RecallTimer 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   7320
      Top             =   960
   End
   Begin VB.Timer HoistTimer 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   6720
      Top             =   960
   End
   Begin VB.CommandButton cmdEvents 
      Caption         =   "Show Events"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7080
      TabIndex        =   24
      Top             =   0
      Width           =   1095
   End
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
         TabIndex        =   20
         Top             =   120
         Width           =   2055
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshFinish 
            Height          =   3855
            Left            =   120
            TabIndex        =   21
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
            Index           =   5
            Left            =   720
            Style           =   1  'Graphical
            TabIndex        =   26
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
            Height          =   495
            Index           =   4
            Left            =   600
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   360
            Width           =   1095
         End
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
            Caption         =   "00:00:00"
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
            Caption         =   "00:00"
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
            Width           =   960
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
            TabIndex        =   22
            Top             =   240
            Width           =   2055
            Begin VB.ComboBox cboProfile 
               Height          =   315
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   23
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
               Text            =   "0000"
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
      Left            =   120
      ScaleHeight     =   1575
      ScaleWidth      =   6015
      TabIndex        =   1
      Tag             =   "2"
      Top             =   0
      Width           =   6015
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
      Top             =   6090
      Width           =   8205
      _ExtentX        =   14473
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
Private Type defCols
    Group As String    'Dynamic, Index is ColCount
    Items As Long        'Dynamic
End Type
Private Cols() As defCols
Private NextFreeCol As Long
Private FirstStartTime As Date
Private LastTimeOutput As Date  'Used for catch-up
Private NextCommandTop As Single
Private CmdQ(8) As Integer     'Idx of next signal (if timer)
Private LastHoist As String  'Group of last flag hoisted (Timer Suppresses sound signal)
Private LastStart As Long    'Idx of last Class flag lowered (supresses Queueing Recall)

Private Sub cboProfile_Click()
'Because you get an Unable to unload within this context (Error 365)
'if you try and remove the SignalTimerControls from within
'a Combo Click event we set a flag and do it using timer
'see http://msdn.microsoft.com/en-us/library/aa243662%28v=vs.60%29.aspx
vbwProfiler.vbwProcIn 1
vbwProfiler.vbwExecuteLine 1
    ReloadTimer.Enabled = True
vbwProfiler.vbwProcOut 1
vbwProfiler.vbwExecuteLine 2
End Sub

'Load startsequencies - only on initial startup
Private Sub LoadSequence()
vbwProfiler.vbwProcIn 2
Dim i As Long

vbwProfiler.vbwExecuteLine 3
    With cboProfile
'Only load if no files already loaded
vbwProfiler.vbwExecuteLine 4
        If .ListCount = 0 Then
vbwProfiler.vbwExecuteLine 5
            IniFileName = Dir$(Environ("AllUsersProfile") & "\Application Data\Arundale\RacingSignals\Sequences\*.ini")
'vbwLine 6:            Do While IniFileName > ""
            Do While vbwProfiler.vbwExecuteLine(6) Or IniFileName > ""
'Dont allow *.ini_old
vbwProfiler.vbwExecuteLine 7
                If Right$(IniFileName, 4) = ".ini" Then
'Remove .ini so it's not displayed
vbwProfiler.vbwExecuteLine 8
                    i = InStrRev(IniFileName, ".ini")
vbwProfiler.vbwExecuteLine 9
                    If i > 0 Then
vbwProfiler.vbwExecuteLine 10
                        IniFileName = Left$(IniFileName, i - 1)
vbwProfiler.vbwExecuteLine 11
                        .AddItem IniFileName
                    End If
vbwProfiler.vbwExecuteLine 12 'B
                End If
vbwProfiler.vbwExecuteLine 13 'B
vbwProfiler.vbwExecuteLine 14
                IniFileName = Dir$
vbwProfiler.vbwExecuteLine 15
            Loop
'If none exit sub & program (in .main)
vbwProfiler.vbwExecuteLine 16
            If .ListCount = 0 Then
vbwProfiler.vbwExecuteLine 17
                MsgBox "No Start Sequences available" & vbCrLf & "Exiting Program", vbCritical, "No Start Sequence"
vbwProfiler.vbwProcOut 2
vbwProfiler.vbwExecuteLine 18
                Exit Sub
            End If
vbwProfiler.vbwExecuteLine 19 'B
        End If
vbwProfiler.vbwExecuteLine 20 'B
'If No Sequence selected
vbwProfiler.vbwExecuteLine 21
        If .ListIndex = -1 Then
vbwProfiler.vbwExecuteLine 22
            MsgBox "Please select a Start Sequence", vbExclamation, "No Start Sequence"
        Else
vbwProfiler.vbwExecuteLine 23 'B
'When a profile is first selected dont use the reload timer
vbwProfiler.vbwExecuteLine 24
            Call LoadProfile
'Dont reset Postpone unless loading the profile
'because once racing has been postpned the Signal stay up until
'all starts have been concluded
'    cmdPostpone.BackColor = vbGreen
        End If
vbwProfiler.vbwExecuteLine 25 'B
vbwProfiler.vbwExecuteLine 26
    End With
vbwProfiler.vbwProcOut 2
vbwProfiler.vbwExecuteLine 27
End Sub
Private Sub cmdFinish_Click_old()
vbwProfiler.vbwProcIn 3
vbwProfiler.vbwExecuteLine 28
    Call MakeSignals(SignalIdx("Flag", "Finish"), True)
vbwProfiler.vbwExecuteLine 29
    Call MakeSignals(SignalIdx("Flag", "Horn1Short"), True)
vbwProfiler.vbwExecuteLine 30
    With mshFinish
'not the first (blank) row
vbwProfiler.vbwExecuteLine 31
        If .TextMatrix(.Rows - 1, 0) <> "" Then
vbwProfiler.vbwExecuteLine 32
            .Rows = .Rows + 1
        End If
vbwProfiler.vbwExecuteLine 33 'B
vbwProfiler.vbwExecuteLine 34
        .TextMatrix(.Rows - 1, 0) = .Rows - 1
vbwProfiler.vbwExecuteLine 35
        .TextMatrix(.Rows - 1, 1) = lblCurrTime.Caption
'Scroll to bottom
vbwProfiler.vbwExecuteLine 36
        .TopRow = .Rows - 1
vbwProfiler.vbwExecuteLine 37
End With
vbwProfiler.vbwExecuteLine 38
Debug.Print "Finish"

vbwProfiler.vbwProcOut 3
vbwProfiler.vbwExecuteLine 39
End Sub

#If False Then
Private Sub cmdHorn_Click_old()
    
    Call MakeSignals(SignalIdx("Flag", "Horn"), True)
Debug.Print "Horn"
End Sub


Private Sub cmdPostpone_Click_old()
    
'    cmdPostpone.BackColor = vbRed
    
    If ValidatePostponeTime = True Then
'Causes StartTime to be validated
'Which then causes Events to be reloaded
        txtFirstStartTime = Format$((DateAdd("n", CDbl(NulToZero(txtPostpone.Text)), FirstStartTime)), "hhnn")
    End If
End Sub

Private Sub cmdRecall_Click_old()
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
            SignalAttributes(SignalIdx("Recall")).OnCycles = 0
            cmdFinish.BackColor = vbGreen
            cmdFinish.SetFocus
        End If
    Case Else
'Dont action
    End Select
Debug.Print "Recall"

End Sub
#End If

'Only used to clear all the flags off the display 3 secs after loading the profile
Private Sub ClearFlagsTimer_Timer()
vbwProfiler.vbwProcIn 4
Dim Idx As Long
Dim i As Long
vbwProfiler.vbwExecuteLine 40
    ClearFlagsTimer.Enabled = False
vbwProfiler.vbwExecuteLine 41
    For Idx = 1 To UBound(SignalAttributes)
vbwProfiler.vbwExecuteLine 42
        With SignalAttributes(Idx)
vbwProfiler.vbwExecuteLine 43
            If SignalAttributes(Idx).Flag.Pos <> 0 Then
vbwProfiler.vbwExecuteLine 44
                Call LowerFlag(Idx)
            End If
vbwProfiler.vbwExecuteLine 45 'B
vbwProfiler.vbwExecuteLine 46
        End With
vbwProfiler.vbwExecuteLine 47
    Next Idx

vbwProfiler.vbwExecuteLine 48
    HoistTimer.Enabled = False
vbwProfiler.vbwExecuteLine 49
    RecallTimer.Enabled = False
vbwProfiler.vbwExecuteLine 50
Debug.Print "RecallTimer disabled"
vbwProfiler.vbwExecuteLine 51
    LastHoist = ""
vbwProfiler.vbwExecuteLine 52
    i = SignalFromName("Recall Class")
vbwProfiler.vbwExecuteLine 53
    If i > 0 Then
vbwProfiler.vbwExecuteLine 54
         Set SignalAttributes(i).Image = Nothing
    End If
vbwProfiler.vbwExecuteLine 55 'B
vbwProfiler.vbwExecuteLine 56
    LastStart = 0
vbwProfiler.vbwExecuteLine 57
    Call ResetCols
vbwProfiler.vbwProcOut 4
vbwProfiler.vbwExecuteLine 58
End Sub

Private Sub cmdEvents_Click()
vbwProfiler.vbwProcIn 5
vbwProfiler.vbwExecuteLine 59
    If frmEvents.Visible Then
vbwProfiler.vbwExecuteLine 60
        Unload frmEvents
    Else
vbwProfiler.vbwExecuteLine 61 'B
vbwProfiler.vbwExecuteLine 62
        frmEvents.Visible = True
    End If
vbwProfiler.vbwExecuteLine 63 'B
vbwProfiler.vbwProcOut 5
vbwProfiler.vbwExecuteLine 64
End Sub

Private Sub Commands_Click(Index As Integer)
vbwProfiler.vbwProcIn 6
Dim Position As Long
Dim NextCommand As Long

vbwProfiler.vbwExecuteLine 65
Debug.Print "--- " & Commands(Index).Caption & " ---"

vbwProfiler.vbwExecuteLine 66
    With SignalAttributes(Index)
'If this command is queued then just remove it (same as clicking when up)
'This must be done in the Click event because the user is making the request
'You cannot do it in RaiseRequest or LowerRequest because all queued events
'would get removed.
vbwProfiler.vbwExecuteLine 67
        If Commands(Index).BackColor = vbCyan Then
vbwProfiler.vbwExecuteLine 68
            NextCommand = DequeCmd(.Group)
vbwProfiler.vbwProcOut 6
vbwProfiler.vbwExecuteLine 69
            Exit Sub
        End If
vbwProfiler.vbwExecuteLine 70 'B
'If we have another commandButton queued in this group, remove this before
'actioning a raise request so we dont have 2 flags in same group queued
'This is important with Recall & General Recall
vbwProfiler.vbwExecuteLine 71
        If .Flag.Pos = 0 Then
vbwProfiler.vbwExecuteLine 72
            NextCommand = DequeCmd(.Group)
vbwProfiler.vbwExecuteLine 73
            Call RaiseRequest(CLng(Index))
        Else
vbwProfiler.vbwExecuteLine 74 'B
vbwProfiler.vbwExecuteLine 75
            Call LowerRequest(CLng(Index))
        End If
vbwProfiler.vbwExecuteLine 76 'B
vbwProfiler.vbwExecuteLine 77
    End With
vbwProfiler.vbwProcOut 6
vbwProfiler.vbwExecuteLine 78
End Sub


Private Sub Flags_Click(Index As Integer)
'MsgBox Flags(Index).Picture.Handle
vbwProfiler.vbwProcIn 7
vbwProfiler.vbwProcOut 7
vbwProfiler.vbwExecuteLine 79
End Sub

Private Sub Form_Load()
vbwProfiler.vbwProcIn 8
Dim i As Long
Dim url As String
Dim Major As Long
Dim Minor As Long
Dim Revision As Long
Dim NewVersion As Boolean

vbwProfiler.vbwExecuteLine 80
    Caption = App.EXEName & " [" & App.Major & "." & App.Minor & "." _
    & App.Revision & "] "

'Check if a later version exists
vbwProfiler.vbwExecuteLine 81
    url = "http://www.NmeaRouter.com/docs/ais/" & App.EXEName _
    & "_setup_"
vbwProfiler.vbwExecuteLine 82
    Major = App.Major
vbwProfiler.vbwExecuteLine 83
    Do
vbwProfiler.vbwExecuteLine 84
        If HTTPFileExists(url & Major & ".0.0.exe") = False Then
vbwProfiler.vbwExecuteLine 85
             Exit Do
        End If
vbwProfiler.vbwExecuteLine 86 'B
vbwProfiler.vbwExecuteLine 87
        Major = Major + 1
vbwProfiler.vbwExecuteLine 88
    Loop
vbwProfiler.vbwExecuteLine 89
    If Major > 0 Then 'Highest major that exists
vbwProfiler.vbwExecuteLine 90
         Major = Major - 1
    End If
vbwProfiler.vbwExecuteLine 91 'B

vbwProfiler.vbwExecuteLine 92
    url = url & Major & "."
vbwProfiler.vbwExecuteLine 93
    If Major = App.Major Then
vbwProfiler.vbwExecuteLine 94
        Minor = App.Minor
    Else
vbwProfiler.vbwExecuteLine 95 'B
vbwProfiler.vbwExecuteLine 96
        Minor = 0
    End If
vbwProfiler.vbwExecuteLine 97 'B
vbwProfiler.vbwExecuteLine 98
    Do
vbwProfiler.vbwExecuteLine 99
        If HTTPFileExists(url & Minor & ".0.exe") = False Then
vbwProfiler.vbwExecuteLine 100
             Exit Do
        End If
vbwProfiler.vbwExecuteLine 101 'B
vbwProfiler.vbwExecuteLine 102
        Minor = Minor + 1
vbwProfiler.vbwExecuteLine 103
    Loop
vbwProfiler.vbwExecuteLine 104
    If Minor > 0 Then
vbwProfiler.vbwExecuteLine 105
         Minor = Minor - 1
    End If
vbwProfiler.vbwExecuteLine 106 'B

vbwProfiler.vbwExecuteLine 107
    url = url & Minor & "."
vbwProfiler.vbwExecuteLine 108
    If Not (Major = App.Major And Minor = App.Minor) Then
vbwProfiler.vbwExecuteLine 109
        NewVersion = True
    End If
vbwProfiler.vbwExecuteLine 110 'B
'Only let a user get next revision if he is using a revision
'of his current version. Otherwise he goes up to the next minor version
vbwProfiler.vbwExecuteLine 111
    If NewVersion = False And App.Revision > 0 Then
vbwProfiler.vbwExecuteLine 112
        Revision = App.Revision
vbwProfiler.vbwExecuteLine 113
        Do
vbwProfiler.vbwExecuteLine 114
            If HTTPFileExists(url & Revision & ".exe") = False Then
vbwProfiler.vbwExecuteLine 115
                 Exit Do
            End If
vbwProfiler.vbwExecuteLine 116 'B
vbwProfiler.vbwExecuteLine 117
            Revision = Revision + 1
vbwProfiler.vbwExecuteLine 118
        Loop
vbwProfiler.vbwExecuteLine 119
        If Revision > 0 Then
vbwProfiler.vbwExecuteLine 120
             Revision = Revision - 1
        End If
vbwProfiler.vbwExecuteLine 121 'B
vbwProfiler.vbwExecuteLine 122
        If Revision < App.Revision Then
vbwProfiler.vbwExecuteLine 123
            NewVersion = True
        End If
vbwProfiler.vbwExecuteLine 124 'B
    End If
vbwProfiler.vbwExecuteLine 125 'B
vbwProfiler.vbwExecuteLine 126
    url = url & Revision & ".exe"

'If we are working on a higher version in VBE, don't try for newversion
vbwProfiler.vbwExecuteLine 127
    If App.Major * 2 ^ 8 + App.Minor * 2 ^ 4 + App.Revision > _
    Major * 2 ^ 8 + Minor * 2 ^ 4 + Revision Then
vbwProfiler.vbwExecuteLine 128
        NewVersion = False
    End If
vbwProfiler.vbwExecuteLine 129 'B
vbwProfiler.vbwExecuteLine 130
    If NewVersion = True Then
vbwProfiler.vbwExecuteLine 131
        Call frmDpyBox.DpyBox("A new update is available", 10, "New Version")
'Check we have internet access
vbwProfiler.vbwExecuteLine 132
        If HTTPFileExists(url) Then
vbwProfiler.vbwExecuteLine 133
            Call HttpSpawn(url)
        End If
vbwProfiler.vbwExecuteLine 134 'B
    End If
vbwProfiler.vbwExecuteLine 135 'B
'Position cursor at RHS of time displayed
vbwProfiler.vbwExecuteLine 136
    txtFirstStartTime.SelStart = Len(txtFirstStartTime)
vbwProfiler.vbwExecuteLine 137
    txtPostpone.SelStart = Len(txtPostpone)
vbwProfiler.vbwExecuteLine 138
    With mshFinish
vbwProfiler.vbwExecuteLine 139
        .Width = 1795
vbwProfiler.vbwExecuteLine 140
        .FormatString = "<No|<Time"
vbwProfiler.vbwExecuteLine 141
        .ColWidth(0) = 500  'Position
vbwProfiler.vbwExecuteLine 142
        .ColWidth(1) = 1295  'Time
'        For i = 1 To 20
'            .Rows = i + 1
'            .TextMatrix(i, 0) = i
'        Next i
'        .TextMatrix(1, 1) = "13:22:45"
vbwProfiler.vbwExecuteLine 143
    End With

'Flags(0) exists - but not used
vbwProfiler.vbwExecuteLine 144
    RowCount = FlagRow(Flags.Count - 1)
vbwProfiler.vbwExecuteLine 145
    ColCount = FlagCol(Flags.Count - 1)
vbwProfiler.vbwExecuteLine 146
    ColCountFree = ColCount 'Reduces by number of Fixed cols
vbwProfiler.vbwExecuteLine 147
    ReDim Cols(1 To ColCount)
vbwProfiler.vbwExecuteLine 148
Visible = True
'Make the base index invisible as it is not used
vbwProfiler.vbwExecuteLine 149
    Commands(0).Enabled = False
vbwProfiler.vbwExecuteLine 150
    Commands(0).Visible = False
'Set up initial start time, LoadEvents not called
vbwProfiler.vbwExecuteLine 151
    FirstStartTime = Date & " " _
    & Format$(NulToZero(txtFirstStartTime), "00:00") & ":00"
vbwProfiler.vbwExecuteLine 152
Debug.Print Format$(FirstStartTime, "dd-mmm-yyyy")
vbwProfiler.vbwExecuteLine 153
Debug.Print Format$(FirstStartTime, "hh:mm:ss")

vbwProfiler.vbwExecuteLine 154
    Call LoadSequence

vbwProfiler.vbwProcOut 8
vbwProfiler.vbwExecuteLine 155
End Sub


Private Sub Form_Unload(Cancel As Integer)
vbwProfiler.vbwProcIn 9
    Dim i As Integer

    'close all sub forms
vbwProfiler.vbwExecuteLine 156
    For i = Forms.Count - 1 To 1 Step -1
vbwProfiler.vbwExecuteLine 157
        Unload Forms(i)
vbwProfiler.vbwExecuteLine 158
    Next
vbwProfiler.vbwExecuteLine 159
    If Me.WindowState <> vbMinimized Then
vbwProfiler.vbwExecuteLine 160
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
vbwProfiler.vbwExecuteLine 161
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
vbwProfiler.vbwExecuteLine 162
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
vbwProfiler.vbwExecuteLine 163
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
vbwProfiler.vbwExecuteLine 164 'B
vbwProfiler.vbwProcOut 9
vbwProfiler.vbwExecuteLine 165
End Sub

Private Sub RecallTimer_Timer()
vbwProfiler.vbwProcIn 10
Dim i As Long

vbwProfiler.vbwExecuteLine 166
    RecallTimer.Enabled = False
'Set laststart to 0 so any subsequent Recalls will be queued, if a class flag is UP
vbwProfiler.vbwExecuteLine 167
    i = SignalFromName("Recall Class")
vbwProfiler.vbwExecuteLine 168
    If i > 0 Then
vbwProfiler.vbwExecuteLine 169
         Set SignalAttributes(i).Image = Nothing
    End If
vbwProfiler.vbwExecuteLine 170 'B
vbwProfiler.vbwExecuteLine 171
    LastStart = 0
vbwProfiler.vbwExecuteLine 172
Debug.Print "RecallTimer disabled"
vbwProfiler.vbwProcOut 10
vbwProfiler.vbwExecuteLine 173
End Sub

Private Sub HoistTimer_Timer()
vbwProfiler.vbwProcIn 11
vbwProfiler.vbwExecuteLine 174
    HoistTimer.Enabled = False
'Set last hoist to Blank so any subsequent hoist will action the sound signal
vbwProfiler.vbwExecuteLine 175
    LastHoist = ""
vbwProfiler.vbwExecuteLine 176
Debug.Print "HoistTimer disabled"
vbwProfiler.vbwProcOut 11
vbwProfiler.vbwExecuteLine 177
End Sub

'The timer must run faster than 1 Count interval (normally 1 sec)
'Otherwise a NewCycle could be skipped if the timer is running
'slower than the actual clock time. This could happen
'if the PC is heavily loaded
Private Sub RaceTimer_Timer()
vbwProfiler.vbwProcIn 12
Dim CurrTime As Date    'May be speeded up from Now() time for testing
Dim SecsSinceOutput As Long
Dim TimeToOutput As Date
'Dim MyFirstStartTime As Date

'Adjust curr time to speed up
vbwProfiler.vbwExecuteLine 178
    CurrTime = Now()

vbwProfiler.vbwExecuteLine 179
    If LastTimeOutput = "00:00:00" Then
vbwProfiler.vbwExecuteLine 180
         Call ResetOutput(CurrTime)
    End If
vbwProfiler.vbwExecuteLine 181 'B

vbwProfiler.vbwExecuteLine 182
    SecsSinceOutput = DateDiff("s", LastTimeOutput, CurrTime)
'No Output due yet
vbwProfiler.vbwExecuteLine 183
    If SecsSinceOutput = 0 Then
vbwProfiler.vbwExecuteLine 184
Debug.Print "Skip"
vbwProfiler.vbwProcOut 12
vbwProfiler.vbwExecuteLine 185
        Exit Sub
    End If
vbwProfiler.vbwExecuteLine 186 'B

vbwProfiler.vbwExecuteLine 187
    Do
'        StatusBar1.Panels(1).Text = Time
'        StatusBar1.Panels(2).Text = CurSecs - CycleStartSecs
vbwProfiler.vbwExecuteLine 188
        TimeToOutput = DateAdd("s", 1, LastTimeOutput)
vbwProfiler.vbwExecuteLine 189
        If TimerOutput(TimeToOutput) = True Then
vbwProfiler.vbwExecuteLine 190
             LastTimeOutput = TimeToOutput
        End If
vbwProfiler.vbwExecuteLine 191 'B
vbwProfiler.vbwExecuteLine 192
        SecsSinceOutput = DateDiff("s", LastTimeOutput, CurrTime)
vbwProfiler.vbwExecuteLine 193
        ElapsedTime = DateDiff("s", FirstStartTime, CurrTime)
'Debug.Print ElapsedTime - SecsSinceOutput
vbwProfiler.vbwExecuteLine 194
        Call DoTimerEvents(ElapsedTime - SecsSinceOutput)
vbwProfiler.vbwExecuteLine 195
        lblElapsedTime = aSecToElapsed(ElapsedTime)
'        lblElapsedTime = aSecToElapsed(DateDiff("s", FirstStartTime, CurrTime))
vbwProfiler.vbwExecuteLine 196
        lblCurrTime = Format$(CurrTime, "hh:mm:ss")
vbwProfiler.vbwExecuteLine 197
If SecsSinceOutput > 0 Then
vbwProfiler.vbwExecuteLine 198
     Debug.Print "Catch-up " & SecsSinceOutput
End If
vbwProfiler.vbwExecuteLine 199 'B
vbwProfiler.vbwExecuteLine 200
    Loop Until SecsSinceOutput = 0  'Always execute once
vbwProfiler.vbwProcOut 12
vbwProfiler.vbwExecuteLine 201
End Sub


Private Sub ResetOutput(StartTime As Date)
vbwProfiler.vbwProcIn 13
vbwProfiler.vbwExecuteLine 202
        LastTimeOutput = DateAdd("s", -1, StartTime)
vbwProfiler.vbwProcOut 13
vbwProfiler.vbwExecuteLine 203
End Sub

Private Function aSecToElapsed(Secs As Long) As String
vbwProfiler.vbwProcIn 14
Dim hms As defhms
Dim Sign As Long
Dim aSign As String

'Secs = 3600& * 100&
vbwProfiler.vbwExecuteLine 204
    Sign = Sgn(Secs)    '-1 = -ve, 0 = 0 , +1 = +ve
vbwProfiler.vbwExecuteLine 205
    If Sign = -1 Then
vbwProfiler.vbwExecuteLine 206
        Secs = Secs * Sign 'force +ve
vbwProfiler.vbwExecuteLine 207
        aSign = "-"
    Else
vbwProfiler.vbwExecuteLine 208 'B
vbwProfiler.vbwExecuteLine 209
        aSign = " "
    End If
vbwProfiler.vbwExecuteLine 210 'B
vbwProfiler.vbwExecuteLine 211
    hms.Hour = Int(Secs / 3600&)
vbwProfiler.vbwExecuteLine 212
    Secs = Secs - hms.Hour * 3600&
vbwProfiler.vbwExecuteLine 213
    hms.Min = Int(Secs / 60&)
vbwProfiler.vbwExecuteLine 214
    Secs = Secs - hms.Min * 60&
vbwProfiler.vbwExecuteLine 215
    hms.Sec = Secs
vbwProfiler.vbwExecuteLine 216
    aSecToElapsed = aSign & Format$(hms.Hour, "###")
vbwProfiler.vbwExecuteLine 217
    If Abs(hms.Hour) >= 1 Then
vbwProfiler.vbwExecuteLine 218
         aSecToElapsed = aSecToElapsed & ":"
    End If
vbwProfiler.vbwExecuteLine 219 'B
vbwProfiler.vbwExecuteLine 220
    aSecToElapsed = aSecToElapsed & Format$(hms.Min, "00") _
    & ":" & Format$(hms.Sec, "00")
vbwProfiler.vbwProcOut 14
vbwProfiler.vbwExecuteLine 221
End Function

Private Sub ReloadTimer_Timer()
vbwProfiler.vbwProcIn 15
vbwProfiler.vbwExecuteLine 222
    ReloadTimer.Enabled = False
vbwProfiler.vbwExecuteLine 223
    Call LoadProfile
vbwProfiler.vbwProcOut 15
vbwProfiler.vbwExecuteLine 224
End Sub

Private Sub SignalTimer_Timer(Index As Integer)
vbwProfiler.vbwProcIn 16
Dim FlagIdx  As Long
Dim kb As String
Dim CyclesCompleted As Long
Dim LinkedFlagPos As Long

vbwProfiler.vbwExecuteLine 225
    With SignalAttributes(Index)
vbwProfiler.vbwExecuteLine 226
kb = SignalTimer(Index).Enabled
'Debug.Print Flags(FlagIdx).Visible
'A cycle is completed every time a flag is turned off AFTER it has been on

vbwProfiler.vbwExecuteLine 227
        If .Flag.Pos Then
vbwProfiler.vbwExecuteLine 228
            If Flags(.Flag.Pos).Visible = True Then
vbwProfiler.vbwExecuteLine 229
                .OnCycles = .OnCycles + 1
vbwProfiler.vbwExecuteLine 230
                SignalTimer(Index).Interval = .TTD
            Else
vbwProfiler.vbwExecuteLine 231 'B
vbwProfiler.vbwExecuteLine 232
                SignalTimer(Index).Interval = .TTL
vbwProfiler.vbwExecuteLine 233
                CyclesCompleted = .OnCycles

            End If
vbwProfiler.vbwExecuteLine 234 'B
        Else
vbwProfiler.vbwExecuteLine 235 'B
vbwProfiler.vbwExecuteLine 236
            .OnCycles = .OnCycles + 1
'Terminate Timer & Lower flag
vbwProfiler.vbwExecuteLine 237
            CyclesCompleted = .OnCycles
'MsgBox "Signal(" & Index & ")." & .Name & " has no associated Flag", vbCritical, "SignalTimer_Timer"
        End If
vbwProfiler.vbwExecuteLine 238 'B
'Debug.Print CyclesCompleted & "(" & Index & ")"

'Continuous
vbwProfiler.vbwExecuteLine 239
        If .CyclesRequired = 0 Then
vbwProfiler.vbwExecuteLine 240
            If Loading = False Then
vbwProfiler.vbwExecuteLine 241
                 CyclesCompleted = -1
            End If
vbwProfiler.vbwExecuteLine 242 'B
        End If
vbwProfiler.vbwExecuteLine 243 'B

vbwProfiler.vbwExecuteLine 244
        If Loading And CyclesCompleted > 5 Then
vbwProfiler.vbwExecuteLine 245
             CyclesCompleted = .CyclesRequired
        End If
vbwProfiler.vbwExecuteLine 246 'B
vbwProfiler.vbwExecuteLine 247
        Select Case CyclesCompleted
'The timer has started but we do not want the Signal Off
'In fact we should not have started it in the first place
'vbwLine 248:        Case Is >= .CyclesRequired
        Case Is >= IIf(vbwProfiler.vbwExecuteLine(248), VBWPROFILER_EMPTY, _
        .CyclesRequired)
'This only occurs when the flag is about to be made invisible
'Turn off Signal, before disabling the timer
'Otherwise MakeSignals will start it again
'Click the command button (set to True) to put the flag down
'Only disable if not Continuous
vbwProfiler.vbwExecuteLine 249
            SignalTimer(Index).Enabled = False
'Must be after timer is disabled
vbwProfiler.vbwExecuteLine 250
            .OnCycles = 0
vbwProfiler.vbwExecuteLine 251
            Call LowerRequest(Index)
'        Commands(Index).Value = True
'Click the command button
'kb = SignalTimer(Index).Enabled    'Must be turned off
'Do this last, so if the timer is called again
'another off will be generated, and the timer will
'not re-start
'Remove this from the queue and re-enable with next signal (if any)
'        Call DequeTimer(Index)
'vbwLine 252:        Case Is > .CyclesRequired
        Case Is > IIf(vbwProfiler.vbwExecuteLine(252), VBWPROFILER_EMPTY, _
        .CyclesRequired)
'Continuous
'vbwLine 253:        Case Is < .CyclesRequired
        Case Is < IIf(vbwProfiler.vbwExecuteLine(253), VBWPROFILER_EMPTY, _
        .CyclesRequired)
'Reverse the Visibility of this flag and do another cycle
'No linked Flags are activated
vbwProfiler.vbwExecuteLine 254
            Call FlagVisibility(Index, Not Flags(.Flag.Pos).Visible)

'Change the Visibility of any Linked flag UP Position only
'Because if it is the Horn that is linked we do not want to keep cycling it
''            If .Linkup(lidx).Flag > 0 Then
'Linked Flag must be raised as well (Pos > 0)
''                If SignalAttributes(.Linkup(lidx).Flag).Flag.Pos > 0 Then
''                    Call FlagVisibility(.Linkup(lidx).Flag, Flags(.Flag.Pos).Visible)
''                End If
''            End If
'Keep the timer running
        End Select
vbwProfiler.vbwExecuteLine 255 'B
vbwProfiler.vbwExecuteLine 256
    End With
vbwProfiler.vbwProcOut 16
vbwProfiler.vbwExecuteLine 257
End Sub

Private Sub txtFirstStartTime_Change()
vbwProfiler.vbwProcIn 17
vbwProfiler.vbwExecuteLine 258
    If ValidateStartTime = True Then
vbwProfiler.vbwExecuteLine 259
        Call ResetCmd
    End If
vbwProfiler.vbwExecuteLine 260 'B
vbwProfiler.vbwProcOut 17
vbwProfiler.vbwExecuteLine 261
End Sub

Private Sub txtPostpone_Change()
vbwProfiler.vbwProcIn 18
vbwProfiler.vbwExecuteLine 262
    Call ValidatePostponeTime
vbwProfiler.vbwProcOut 18
vbwProfiler.vbwExecuteLine 263
End Sub

Private Function ValidateStartTime() As Boolean
vbwProfiler.vbwProcIn 19

vbwProfiler.vbwExecuteLine 264
    On Error GoTo ValidateStartTime_error
vbwProfiler.vbwExecuteLine 265
    If txtFirstStartTime = "" Then
vbwProfiler.vbwExecuteLine 266
        txtFirstStartTime.BackColor = vbRed
    Else
vbwProfiler.vbwExecuteLine 267 'B
vbwProfiler.vbwExecuteLine 268
        txtFirstStartTime.BackColor = vbWhite
    End If
vbwProfiler.vbwExecuteLine 269 'B
vbwProfiler.vbwExecuteLine 270
    If Len(txtFirstStartTime) = 4 _
    And CLng(NulToZero(txtFirstStartTime)) >= 0 _
    And CLng(NulToZero(txtFirstStartTime)) <= 2400 _
    And IsNumeric(NulToZero(txtFirstStartTime)) = True Then
vbwProfiler.vbwExecuteLine 271
        FirstStartTime = Date & " " _
        & Format$(NulToZero(txtFirstStartTime), "00:00") & ":00"
vbwProfiler.vbwExecuteLine 272
        On Error GoTo 0
vbwProfiler.vbwExecuteLine 273
        txtFirstStartTime.ForeColor = vbBlack
'Must not only reset the flags because once the start sequence
'has commenced the whole profile should be reloaded
'        Call ResetFlags
vbwProfiler.vbwExecuteLine 274
        ValidateStartTime = True
vbwProfiler.vbwProcOut 19
vbwProfiler.vbwExecuteLine 275
        Exit Function
    End If
vbwProfiler.vbwExecuteLine 276 'B
ValidateStartTime_error:
vbwProfiler.vbwExecuteLine 277
    txtFirstStartTime.ForeColor = vbRed
vbwProfiler.vbwProcOut 19
vbwProfiler.vbwExecuteLine 278
End Function

Private Function ValidatePostponeTime() As Boolean
vbwProfiler.vbwProcIn 20

vbwProfiler.vbwExecuteLine 279
    On Error GoTo ValidatePostponeTime_error
vbwProfiler.vbwExecuteLine 280
    If txtPostpone = "" Then
vbwProfiler.vbwExecuteLine 281
        txtPostpone.BackColor = vbRed
    Else
vbwProfiler.vbwExecuteLine 282 'B
vbwProfiler.vbwExecuteLine 283
        txtPostpone.BackColor = vbWhite
    End If
vbwProfiler.vbwExecuteLine 284 'B
vbwProfiler.vbwExecuteLine 285
    If IsNumeric(NulToZero(txtPostpone)) = True Then
vbwProfiler.vbwExecuteLine 286
        txtPostpone.ForeColor = vbBlack
vbwProfiler.vbwExecuteLine 287
        ValidatePostponeTime = True
vbwProfiler.vbwProcOut 20
vbwProfiler.vbwExecuteLine 288
        Exit Function
    End If
vbwProfiler.vbwExecuteLine 289 'B
ValidatePostponeTime_error:
vbwProfiler.vbwExecuteLine 290
    txtPostpone.ForeColor = vbRed
vbwProfiler.vbwProcOut 20
vbwProfiler.vbwExecuteLine 291
End Function

Public Function ResetFlags()
vbwProfiler.vbwProcIn 21
Dim MyImage As Image

vbwProfiler.vbwExecuteLine 292
    For Each MyImage In frmMain.Flags
vbwProfiler.vbwExecuteLine 293
        MyImage.Picture = Nothing
vbwProfiler.vbwExecuteLine 294
    Next
vbwProfiler.vbwProcOut 21
vbwProfiler.vbwExecuteLine 295
End Function

#If False Then
Public Function ResetEvents()
    Set CurrEvent = Nothing
End Function
#End If

Public Function ResetCommands()
vbwProfiler.vbwProcIn 22
Dim MyCommand As CommandButton
vbwProfiler.vbwExecuteLine 296
    For Each MyCommand In Commands
vbwProfiler.vbwExecuteLine 297
        If MyCommand.Index <> 0 Then
vbwProfiler.vbwExecuteLine 298
            MyCommand.Enabled = True
vbwProfiler.vbwExecuteLine 299
            MyCommand.Visible = True
        End If
vbwProfiler.vbwExecuteLine 300 'B
vbwProfiler.vbwExecuteLine 301
    Next MyCommand
vbwProfiler.vbwProcOut 22
vbwProfiler.vbwExecuteLine 302
End Function
Public Function ResetSignalTimers()
vbwProfiler.vbwProcIn 23
Dim MySignalTimer As Timer
Dim i As Long
vbwProfiler.vbwExecuteLine 303
    For Each MySignalTimer In frmMain.SignalTimer
vbwProfiler.vbwExecuteLine 304
        If MySignalTimer.Index > 0 Then  'Dont delete SignalTimer(0)
vbwProfiler.vbwExecuteLine 305
            Unload MySignalTimer
        End If
vbwProfiler.vbwExecuteLine 306 'B
vbwProfiler.vbwExecuteLine 307
    Next
vbwProfiler.vbwExecuteLine 308
    HoistTimer.Enabled = False
vbwProfiler.vbwExecuteLine 309
    LastHoist = ""
vbwProfiler.vbwExecuteLine 310
    RecallTimer.Enabled = False
vbwProfiler.vbwExecuteLine 311
    i = SignalFromName("Recall Class")
vbwProfiler.vbwExecuteLine 312
    If i > 0 Then
vbwProfiler.vbwExecuteLine 313
         Set SignalAttributes(i).Image = Nothing
    End If
vbwProfiler.vbwExecuteLine 314 'B
vbwProfiler.vbwExecuteLine 315
    LastStart = 0
vbwProfiler.vbwExecuteLine 316
Debug.Print "HoistTimer disabled"
vbwProfiler.vbwProcOut 23
vbwProfiler.vbwExecuteLine 317
End Function

Public Function ResetFinish()
vbwProfiler.vbwProcIn 24
Dim Row As Long
Dim Col As Long

vbwProfiler.vbwExecuteLine 318
    With mshFinish
'Clear rows (except 1)
vbwProfiler.vbwExecuteLine 319
        For Row = 2 To .Rows - 1
vbwProfiler.vbwExecuteLine 320
            .RemoveItem 1
vbwProfiler.vbwExecuteLine 321
        Next Row
'Clear Row 1
vbwProfiler.vbwExecuteLine 322
        For Col = 0 To .Cols - 1
vbwProfiler.vbwExecuteLine 323
            .TextMatrix(1, Col) = ""
vbwProfiler.vbwExecuteLine 324
        Next Col
vbwProfiler.vbwExecuteLine 325
    End With
vbwProfiler.vbwProcOut 24
vbwProfiler.vbwExecuteLine 326
End Function

'Called when the validated Start Time is changed
Public Function ResetCmd()
vbwProfiler.vbwProcIn 25
vbwProfiler.vbwExecuteLine 327
    Call CommandButtonVisibility(-1000)   'Before start time
'Must have Property Style set to 1=Graphical
'    cmdPostpone.Enabled = True
'Recall Signal says up
'    cmdRecall.BackColor = cbDefault
'    cmdFinish.BackColor = cbDefault
'    cmdHorn.BackColor = cbDefault
'    cmdPostpone.SetFocus
vbwProfiler.vbwProcOut 25
vbwProfiler.vbwExecuteLine 328
End Function

'Requires MSINET.OCX
'See http://officeone.mvps.org/vba/http_file_exists.html
Public Function HTTPFileExists(ByVal url As String) As Boolean
vbwProfiler.vbwProcIn 26
    Dim S As String
    Dim Exists As Boolean
vbwProfiler.vbwExecuteLine 329
    On Error GoTo Inet1_Error
vbwProfiler.vbwExecuteLine 330
    With Inet1
vbwProfiler.vbwExecuteLine 331
        .RequestTimeout = 20
vbwProfiler.vbwExecuteLine 332
        .Protocol = icHTTP
vbwProfiler.vbwExecuteLine 333
        .url = url
vbwProfiler.vbwExecuteLine 334
        .Execute
'see http://support.microsoft.com/kb/182152 =True doesnt work
'vbwLine 335:        Do While .StillExecuting <> False
        Do While vbwProfiler.vbwExecuteLine(335) Or .StillExecuting <> False
vbwProfiler.vbwExecuteLine 336
            DoEvents
vbwProfiler.vbwExecuteLine 337
        Loop
vbwProfiler.vbwExecuteLine 338
        S = UCase(.GetHeader())
vbwProfiler.vbwExecuteLine 339
        Exists = (InStr(1, S, "200 OK") > 0)
vbwProfiler.vbwExecuteLine 340
    End With
vbwProfiler.vbwExecuteLine 341
    HTTPFileExists = Exists
vbwProfiler.vbwProcOut 26
vbwProfiler.vbwExecuteLine 342
    Exit Function
Inet1_Error:
vbwProfiler.vbwExecuteLine 343
    Select Case Err.Number
'vbwLine 344:    Case Is = 35764 '
    Case Is = IIf(vbwProfiler.vbwExecuteLine(344), VBWPROFILER_EMPTY, _
        35764 )'
    End Select
vbwProfiler.vbwExecuteLine 345 'B

vbwProfiler.vbwProcOut 26
vbwProfiler.vbwExecuteLine 346
End Function

Public Function HttpSpawn(url As String)
vbwProfiler.vbwProcIn 27
Dim r As Long
Dim Command As String

vbwProfiler.vbwExecuteLine 347
If Environ("windir") <> "" Then
vbwProfiler.vbwExecuteLine 348
    r = ShellExecute(0, "open", url, 0, 0, 1)
Else
vbwProfiler.vbwExecuteLine 349 'B
'try for linux compatibility
vbwProfiler.vbwExecuteLine 350
    Command = "winebrowser " & url & " ""%1"""

vbwProfiler.vbwExecuteLine 351
    Shell (Command)
End If
vbwProfiler.vbwExecuteLine 352 'B
vbwProfiler.vbwProcOut 27
vbwProfiler.vbwExecuteLine 353
End Function

Public Function PositionCommand(Idx As Long)
'You dont need these unless testing this module in VBE
'If you have a break set frmMain is minimised and
'the Scale values will be 0
'Dont leave a blank gap
vbwProfiler.vbwProcIn 28
vbwProfiler.vbwExecuteLine 354
    With Commands(Idx)
vbwProfiler.vbwExecuteLine 355
        .Caption = .Caption & "(" & Idx & ")"
vbwProfiler.vbwExecuteLine 356
        If .Visible = True Then
'This will be overwritten with the Name from SignalAttributes
'Align first command with top of main frame
vbwProfiler.vbwExecuteLine 357
            WindowState = vbNormal  'Scale will be 0 in VBE (window is minimized)
vbwProfiler.vbwExecuteLine 358
            .Top = ScaleTop + fraMain.Top + NextCommandTop
vbwProfiler.vbwExecuteLine 359
            If .Top + .Height > fraMain.Top + fraMain.Height Then
vbwProfiler.vbwExecuteLine 360
                NextCommandTop = 0
vbwProfiler.vbwExecuteLine 361
                Width = Width + .Width
vbwProfiler.vbwExecuteLine 362
                WindowState = vbNormal  'Scale will be 0 in VBE (window is minimized)
vbwProfiler.vbwExecuteLine 363
                .Top = ScaleTop + fraMain.Top + NextCommandTop
            End If
vbwProfiler.vbwExecuteLine 364 'B
vbwProfiler.vbwExecuteLine 365
            WindowState = vbNormal  'Scale will be 0 in VBE (window is minimized)
vbwProfiler.vbwExecuteLine 366
            .Left = ScaleWidth - .Width
vbwProfiler.vbwExecuteLine 367
            NextCommandTop = NextCommandTop + .Height
        End If
vbwProfiler.vbwExecuteLine 368 'B
vbwProfiler.vbwExecuteLine 369
    End With
vbwProfiler.vbwProcOut 28
vbwProfiler.vbwExecuteLine 370
End Function

Private Function CommandColor()
vbwProfiler.vbwProcIn 29
Dim MyCommand As CommandButton
Dim i As Integer


vbwProfiler.vbwExecuteLine 371
    For Each MyCommand In Commands
vbwProfiler.vbwExecuteLine 372
        If MyCommand.Index > 0 Then 'skip command(0)
'Command may have been created before SignalAttributes
vbwProfiler.vbwExecuteLine 373
            If MyCommand.Index <= UBound(SignalAttributes) Then
vbwProfiler.vbwExecuteLine 374
                If SignalAttributes(MyCommand.Index).Flag.Pos = 0 Then
vbwProfiler.vbwExecuteLine 375
                    MyCommand.BackColor = cbDefault
                Else
vbwProfiler.vbwExecuteLine 376 'B
vbwProfiler.vbwExecuteLine 377
                    MyCommand.BackColor = vbGreen
                End If
vbwProfiler.vbwExecuteLine 378 'B
vbwProfiler.vbwExecuteLine 379
                For i = 0 To UBound(CmdQ)
vbwProfiler.vbwExecuteLine 380
                    If CmdQ(i) = MyCommand.Index Then
vbwProfiler.vbwExecuteLine 381
                        MyCommand.BackColor = vbCyan
                    End If
vbwProfiler.vbwExecuteLine 382 'B
vbwProfiler.vbwExecuteLine 383
                Next i
            End If
vbwProfiler.vbwExecuteLine 384 'B
        End If
vbwProfiler.vbwExecuteLine 385 'B
vbwProfiler.vbwExecuteLine 386
    Next MyCommand
vbwProfiler.vbwProcOut 29
vbwProfiler.vbwExecuteLine 387
End Function

'Called by the a Command to Raise a flag
'Must called by the Link (Sound may be clicked, with Sound still running)
'Queues if fixed position and Fixed position in use)
'Queues the the command if HoistTimer is running for this Group
'Queues Recall if ClassFlag is UP
'Actions Linked Flag by calling LinkRequest (If not Queued)
'Starts HoistTimer for this Group, if not Queueable (Flags.Queue=False)

Public Function RaiseRequest(ByVal Idx As Long)
vbwProfiler.vbwProcIn 30
Dim SoundEnabled As Boolean
Dim Pos As Long
Dim QueueSignal As Long
Dim NextCmd As Long
Dim MyLink As defLink
Dim ClassIdx As Long
Dim RecallIdx As Long
'Dim PreparatoryIdx As Long

vbwProfiler.vbwExecuteLine 388
    If Idx > UBound(SignalAttributes) Then
vbwProfiler.vbwProcOut 30
vbwProfiler.vbwExecuteLine 389
        Exit Function
    End If
vbwProfiler.vbwExecuteLine 390 'B

'Check if Request requires Queueing or actioning
'If Fixed position and Position is in use
vbwProfiler.vbwExecuteLine 391
    With SignalAttributes(Idx)

vbwProfiler.vbwExecuteLine 392
        Select Case .Name
'vbwLine 393:        Case Is = "Finish"
        Case Is = IIf(vbwProfiler.vbwExecuteLine(393), VBWPROFILER_EMPTY, _
        "Finish")
'A Finish Requires Completely different handling, Only Raise the Link UP event
'Do not RaiseRequest to put Finish Flag up (would toggle finish command actions)
vbwProfiler.vbwExecuteLine 394
            If .Name = "Finish" Then
'A Finish always clocks the time and must give correct no Linked signals
vbwProfiler.vbwExecuteLine 395
               If Loading = False Then
vbwProfiler.vbwExecuteLine 396
                    Call FinishTime
'The linked signal requires Queueing, if currently flashing
vbwProfiler.vbwExecuteLine 397
                    Call LinkRequest(Idx)
                End If
vbwProfiler.vbwExecuteLine 398 'B
vbwProfiler.vbwProcOut 30
vbwProfiler.vbwExecuteLine 399
            Exit Function
            End If
vbwProfiler.vbwExecuteLine 400 'B
'vbwLine 401:        Case Is = "Recall", "General Recall"
        Case Is = IIf(vbwProfiler.vbwExecuteLine(401), VBWPROFILER_EMPTY, _
        "Recall"), "General Recall"
'Not actually used if Class flag is above recall
vbwProfiler.vbwExecuteLine 402
            Call LowerGroup(Idx)
        End Select
vbwProfiler.vbwExecuteLine 403 'B
'Debug.Print "RaiseReq " & .Name
vbwProfiler.vbwExecuteLine 404
        Pos = RC(.Flag.FixedRow, .Flag.FixedCol)
'If the Flag has a fixed position, check if any flag is already in this position
vbwProfiler.vbwExecuteLine 405
        If Pos > 0 And .Flag.Queue = True Then
vbwProfiler.vbwExecuteLine 406
            If Flags(Pos).Picture.Handle <> 0 Then
vbwProfiler.vbwExecuteLine 407
                QueueSignal = Idx
'Debug.Print "Q Flag is UP"
            End If
vbwProfiler.vbwExecuteLine 408 'B
        End If
vbwProfiler.vbwExecuteLine 409 'B

'Queues the the command if HoistTimer is running for this Group
'So linked sound signal not made as another flag will be raised on same col
vbwProfiler.vbwExecuteLine 410
        If .Group = LastHoist And .Flag.Queue Then
vbwProfiler.vbwExecuteLine 411
            QueueSignal = Idx
'Debug.Print "Q Timer On"
        End If
vbwProfiler.vbwExecuteLine 412 'B

'Queues Recall if Any ClassFlag is UP (Must Wait until Class Flag is dropped)
'If Recall is pressed within 10 seconds of dropping the Class flag
'It must not be queued as it is a recall for the Class that has just started.
vbwProfiler.vbwExecuteLine 413
        If .Group = "Recall" Then
vbwProfiler.vbwExecuteLine 414
            ClassIdx = GroupIdx("Class")
'Class Flag (may be another Class and NOT the one just started) is up and not
'within 10 secs of last start
vbwProfiler.vbwExecuteLine 415
            If ClassIdx > 0 And LastStart = 0 Then
vbwProfiler.vbwExecuteLine 416
                QueueSignal = Idx
'Debug.Print "Q Class Recall"
            End If
vbwProfiler.vbwExecuteLine 417 'B
        End If
vbwProfiler.vbwExecuteLine 418 'B

vbwProfiler.vbwExecuteLine 419
        If Loading = True Then
vbwProfiler.vbwExecuteLine 420
            QueueSignal = 0
        End If
vbwProfiler.vbwExecuteLine 421 'B

vbwProfiler.vbwExecuteLine 422
        If QueueSignal > 0 Then
'If ClassIdx > 0 Then
'                MyLink.Temp = True
'                MyLink.Flag = QueueSignal
'                MyLink.Raise = True
'                MyLink.Type = "DownLink"
'                Call CreateLink(ClassIdx, MyLink)
'                Commands(QueueSignal).BackColor = vbCyan
'            Else
vbwProfiler.vbwExecuteLine 423
                Call QueueRequest(QueueSignal)
'            End If
        Else
vbwProfiler.vbwExecuteLine 424 'B

'Put the Flag up
vbwProfiler.vbwExecuteLine 425
            Call RaiseFlag(Idx)

'Actions Linked Flag by calling LinkRequest (If not Queued)
vbwProfiler.vbwExecuteLine 426
            Call LinkRequest(Idx)

'Start HoistTimer for this Group, if not Queueable (Flags.Queue=False)
'So we dont Create a Second Sound signal
vbwProfiler.vbwExecuteLine 427
            If .Flag.Queue = False Then
vbwProfiler.vbwExecuteLine 428
                HoistTimer.Enabled = False
vbwProfiler.vbwExecuteLine 429
                HoistTimer.Enabled = True
vbwProfiler.vbwExecuteLine 430
                LastHoist = .Group
vbwProfiler.vbwExecuteLine 431
Debug.Print "HoistTimer Enabled"
            End If
vbwProfiler.vbwExecuteLine 432 'B
        End If  'Not Queued
vbwProfiler.vbwExecuteLine 433 'B

vbwProfiler.vbwExecuteLine 434
    End With

vbwProfiler.vbwProcOut 30
vbwProfiler.vbwExecuteLine 435
End Function

'Called by RaiseRequest and to action the UP link
Public Function RaiseFlag(ByVal Idx As Long)
vbwProfiler.vbwProcIn 31
Dim Col As Long
Dim Row As Long

'Load Profile-Linked Signals with a higher idx will not have been created
'Debug.Print "Raise " & SignalAttributes(Idx).Name
'Action Command now
'Display Image first (if there is one for this Signal)
vbwProfiler.vbwExecuteLine 436
    With SignalAttributes(Idx)

vbwProfiler.vbwExecuteLine 437
        If Not .Image Is Nothing Then
vbwProfiler.vbwExecuteLine 438
            Call NextFreeGroupFlagPos(Idx)
vbwProfiler.vbwExecuteLine 439
            If .Flag.Col > 0 Then
vbwProfiler.vbwExecuteLine 440
                .Flag.Pos = RC(.Flag.Row, .Flag.Col)
            Else
vbwProfiler.vbwExecuteLine 441 'B
vbwProfiler.vbwExecuteLine 442
MsgBox "No free Flag positions", vbCritical, "RaiseFlag"
            End If
vbwProfiler.vbwExecuteLine 443 'B
        End If
vbwProfiler.vbwExecuteLine 444 'B

'If we have a flag position then create it (not set if no Image)
vbwProfiler.vbwExecuteLine 445
        If .Flag.Pos > 0 Then
vbwProfiler.vbwExecuteLine 446
            Flags(.Flag.Pos).Picture = .Image
'You have to set it to False becuase FlagVisibility only reacts to a change
vbwProfiler.vbwExecuteLine 447
            Flags(.Flag.Pos).Visible = False
'Must use flagvisibility to create controller event
vbwProfiler.vbwExecuteLine 448
            Call FlagVisibility(Idx, True)
vbwProfiler.vbwExecuteLine 449
            .Flag.Changed = True
        End If
vbwProfiler.vbwExecuteLine 450 'B

vbwProfiler.vbwExecuteLine 451
        Commands(Idx).BackColor = vbGreen
'May still be a timer even if no image to display
vbwProfiler.vbwExecuteLine 452
        If .TTL > 0 Then
vbwProfiler.vbwExecuteLine 453
            SignalTimer(Idx).Interval = .TTL
vbwProfiler.vbwExecuteLine 454
            SignalTimer(Idx).Enabled = True
        End If
vbwProfiler.vbwExecuteLine 455 'B
vbwProfiler.vbwExecuteLine 456
    End With
vbwProfiler.vbwExecuteLine 457
    Call ResetCols  'Resets Cols().Group & .Items from SignalAttributes

vbwProfiler.vbwProcOut 31
vbwProfiler.vbwExecuteLine 458
End Function

'Called by command to Lower as Flag or when SignalTimer terminates
'Lowers This Flag
'Actions Linked Flag by calling LinkRequest
'Calls LowerFlag to Lower any subservient flags WITHOUT actioning any link
'(Only action the Link of the TOP flag)
'Dequeues Recall when Class Lowered by calling RaiseRequest
'Dequeues any Commands in the same Group Calling RaiseRequest
'Never called by the Link
'Never Queues the Command
Public Function LowerRequest(ByVal Idx As Long)
vbwProfiler.vbwProcIn 32
Dim NextCmd As Integer
Dim i As Long

'Load Profile-Linked Signals with a higher idx will not have been created
vbwProfiler.vbwExecuteLine 459
    If Idx > UBound(SignalAttributes) Then
vbwProfiler.vbwProcOut 32
vbwProfiler.vbwExecuteLine 460
        Exit Function
    End If
vbwProfiler.vbwExecuteLine 461 'B

vbwProfiler.vbwExecuteLine 462
    With SignalAttributes(Idx)
'Debug.Print "LowerReq " & SignalAttributes(Idx).Name

vbwProfiler.vbwExecuteLine 463
For i = 0 To UBound(CmdQ)
vbwProfiler.vbwExecuteLine 464
    If CmdQ(CInt(i)) <> 0 Then
vbwProfiler.vbwExecuteLine 465
Debug.Print "Queued(" & i & ")=" & CmdQ(CInt(i))
    End If
vbwProfiler.vbwExecuteLine 466 'B
vbwProfiler.vbwExecuteLine 467
Next i

 'Lower Flag then any below the Flag
vbwProfiler.vbwExecuteLine 468
        Call LowerFlag(Idx)
vbwProfiler.vbwExecuteLine 469
        Call LinkRequest(Idx)

'Overlapped Position Recall above ClassFlag Requesting Recall
'Dequeues Recall when Any Class Lowered by calling RaiseRequest
vbwProfiler.vbwExecuteLine 470
        If .Group = "Class" Then
vbwProfiler.vbwExecuteLine 471
            NextCmd = DequeCmd("Recall")
vbwProfiler.vbwExecuteLine 472
            If NextCmd <> 0 Then
vbwProfiler.vbwExecuteLine 473
                Call RaiseRequest(NextCmd)
            End If
vbwProfiler.vbwExecuteLine 474 'B
'Call CompressCols
        End If
vbwProfiler.vbwExecuteLine 475 'B

'Dequeues any Commands in the same Group Calling RaiseRequest
vbwProfiler.vbwExecuteLine 476
        NextCmd = DequeCmd(.Group)
vbwProfiler.vbwExecuteLine 477
        If NextCmd <> 0 Then
vbwProfiler.vbwExecuteLine 478
            Call RaiseRequest(NextCmd)
        End If
vbwProfiler.vbwExecuteLine 479 'B
vbwProfiler.vbwExecuteLine 480
    End With

vbwProfiler.vbwExecuteLine 481
    Call ResetCols
'    Call CommandColor

vbwProfiler.vbwProcOut 32
vbwProfiler.vbwExecuteLine 482
End Function

'Called by LowerRequest
'Lowers the Flag and any subservient flags
'Does not action any links
Private Function LowerFlag(ByVal Idx As Long)
vbwProfiler.vbwProcIn 33
Dim StartCol As Long
Dim StartRow As Long
Dim Group As String
Dim i As Long
Dim Remove As Boolean

'Debug.Print "LowerFlag " & SignalAttributes(Idx).Name

vbwProfiler.vbwExecuteLine 483
    With SignalAttributes(Idx)
vbwProfiler.vbwExecuteLine 484
        StartCol = .Flag.Col
vbwProfiler.vbwExecuteLine 485
        StartRow = .Flag.Row
vbwProfiler.vbwExecuteLine 486
        Group = .Group
vbwProfiler.vbwExecuteLine 487
    End With

'Calls LowerFlag to Lower any subservient flags WITHOUT actioning any link
'(Only the Link of the TOP flag is actioned by LowerRequest)
vbwProfiler.vbwExecuteLine 488
    For i = 1 To UBound(SignalAttributes)
vbwProfiler.vbwExecuteLine 489
        With SignalAttributes(i)
vbwProfiler.vbwExecuteLine 490
            If .Group = Group Or (Group = "Class" And .Group = "Preparatory") Then
'If in different col or lower row in same col remove
vbwProfiler.vbwExecuteLine 491
                If .Flag.Col = StartCol And .Flag.Row >= StartRow Then

'Stop first. otherwise Timer will fail when it calls FlagVisibility
vbwProfiler.vbwExecuteLine 492
                    If SignalTimer(i).Enabled = True Then
vbwProfiler.vbwExecuteLine 493
                        SignalTimer(i).Enabled = False
                    End If
vbwProfiler.vbwExecuteLine 494 'B

'Clear the flag (if it exists)
vbwProfiler.vbwExecuteLine 495
                    If Flags(.Flag.Pos).Picture.Handle <> 0 Then
'If .Flag.Pos=0, FlagVisibility reports an error so must do first
vbwProfiler.vbwExecuteLine 496
                        Call FlagVisibility(i, False)
vbwProfiler.vbwExecuteLine 497
                        Flags(.Flag.Pos).Picture = Nothing
                    End If
vbwProfiler.vbwExecuteLine 498 'B
vbwProfiler.vbwExecuteLine 499
                    .Flag.Pos = 0
vbwProfiler.vbwExecuteLine 500
                    .Flag.Col = 0
vbwProfiler.vbwExecuteLine 501
                    .Flag.Row = 0
vbwProfiler.vbwExecuteLine 502
                    Commands(i).BackColor = cbDefault
                End If
vbwProfiler.vbwExecuteLine 503 'B
            End If
vbwProfiler.vbwExecuteLine 504 'B
vbwProfiler.vbwExecuteLine 505
        End With
vbwProfiler.vbwExecuteLine 506
    Next i

'Stop Hoist Timer if last Flag up in this Group
vbwProfiler.vbwExecuteLine 507
    If Group = LastHoist Then
vbwProfiler.vbwExecuteLine 508
        HoistTimer.Enabled = False
vbwProfiler.vbwExecuteLine 509
        LastHoist = ""
vbwProfiler.vbwExecuteLine 510
Debug.Print "HoistTimer disabled"
    End If
vbwProfiler.vbwExecuteLine 511 'B
'Keep the last start Flag for 10 secs
vbwProfiler.vbwExecuteLine 512
    If Group = "Class" Then
vbwProfiler.vbwExecuteLine 513
        RecallTimer.Enabled = True
vbwProfiler.vbwExecuteLine 514
        LastStart = Idx
vbwProfiler.vbwExecuteLine 515
        i = SignalFromName("Recall Class")
vbwProfiler.vbwExecuteLine 516
        If i > 0 Then
vbwProfiler.vbwExecuteLine 517
             Set SignalAttributes(i).Image = SignalAttributes(Idx).Image
        End If
vbwProfiler.vbwExecuteLine 518 'B
vbwProfiler.vbwExecuteLine 519
Debug.Print "RecallTimer enabled"
    End If
vbwProfiler.vbwExecuteLine 520 'B
'    Call CompressCols

vbwProfiler.vbwProcOut 33
vbwProfiler.vbwExecuteLine 521
End Function

'Called by LowerRequest
'Lowers the Flag and any subservient flags
'Does not action any links
Private Function LowerFlag_old(ByVal Idx As Long)
vbwProfiler.vbwProcIn 34
Dim StartCol As Long
Dim StartRow As Long
Dim Group As String
Dim i As Long

'Debug.Print "LowerFlag " & SignalAttributes(Idx).Name

vbwProfiler.vbwExecuteLine 522
    With SignalAttributes(Idx)
vbwProfiler.vbwExecuteLine 523
        StartCol = .Flag.Col
vbwProfiler.vbwExecuteLine 524
        StartRow = .Flag.Row
vbwProfiler.vbwExecuteLine 525
        Group = .Group
vbwProfiler.vbwExecuteLine 526
    End With

'Calls LowerFlag to Lower any subservient flags WITHOUT actioning any link
'(Only the Link of the TOP flag is actioned by LowerRequest)
vbwProfiler.vbwExecuteLine 527
    For i = 1 To UBound(SignalAttributes)
vbwProfiler.vbwExecuteLine 528
        With SignalAttributes(i)
vbwProfiler.vbwExecuteLine 529
            If .Group = Group Then

vbwProfiler.vbwExecuteLine 530
               If .Flag.Pos > 0 Then
'Stop first. otherwise Timer will fail when it calls FlagVisibility
vbwProfiler.vbwExecuteLine 531
                   If SignalTimer(i).Enabled = True Then
vbwProfiler.vbwExecuteLine 532
                        SignalTimer(i).Enabled = False
                    End If
vbwProfiler.vbwExecuteLine 533 'B

'If in different col or lower row in same col remove
'                    If .Flag.Col <> StartCol Or .Flag.Row >= StartRow Then
'Change to only flags in same Col so Class Flags in different Cols are not dropped
vbwProfiler.vbwExecuteLine 534
                    If .Flag.Col = StartCol And .Flag.Row >= StartRow Then
'Clear the flag (if it exists)
vbwProfiler.vbwExecuteLine 535
                        If Flags(.Flag.Pos).Picture.Handle <> 0 Then
'If .Flag.Pos=0, FlagVisibility reports an error so must do first
vbwProfiler.vbwExecuteLine 536
                            Call FlagVisibility(i, False)
vbwProfiler.vbwExecuteLine 537
                            Flags(.Flag.Pos).Picture = Nothing
                        End If
vbwProfiler.vbwExecuteLine 538 'B
vbwProfiler.vbwExecuteLine 539
                        .Flag.Pos = 0
vbwProfiler.vbwExecuteLine 540
                        .Flag.Col = 0
vbwProfiler.vbwExecuteLine 541
                        .Flag.Row = 0
vbwProfiler.vbwExecuteLine 542
                        Commands(i).BackColor = cbDefault
'Stop Hoist Timer if last Flag up in this Group
vbwProfiler.vbwExecuteLine 543
                        If .Group = LastHoist Then
vbwProfiler.vbwExecuteLine 544
                            HoistTimer.Enabled = False
vbwProfiler.vbwExecuteLine 545
                            LastHoist = ""
vbwProfiler.vbwExecuteLine 546
Debug.Print "HoistTimer disabled"
                        End If
vbwProfiler.vbwExecuteLine 547 'B
'Keep the last start Flag for 10 secs
vbwProfiler.vbwExecuteLine 548
                        If .Group = "Class" Then
vbwProfiler.vbwExecuteLine 549
                            RecallTimer.Enabled = True
vbwProfiler.vbwExecuteLine 550
                            LastStart = Idx
vbwProfiler.vbwExecuteLine 551
Debug.Print "RecallTimer enabled"
                        End If
vbwProfiler.vbwExecuteLine 552 'B
                    End If
vbwProfiler.vbwExecuteLine 553 'B
                End If
vbwProfiler.vbwExecuteLine 554 'B
            End If
vbwProfiler.vbwExecuteLine 555 'B
vbwProfiler.vbwExecuteLine 556
        End With
vbwProfiler.vbwExecuteLine 557
    Next i

'    Call CompressCols

vbwProfiler.vbwProcOut 34
vbwProfiler.vbwExecuteLine 558
End Function

'Not actually used if Class flag is above recall
'Lower any flags that are up in this Group (without calling Linked flags)
Private Function LowerGroup(ByVal Idx As Long)
vbwProfiler.vbwProcIn 35
Dim i As Long
Dim Group As String

vbwProfiler.vbwExecuteLine 559
    Group = SignalAttributes(Idx).Group
vbwProfiler.vbwExecuteLine 560
    For i = 1 To UBound(SignalAttributes)
vbwProfiler.vbwExecuteLine 561
        With SignalAttributes(i)
vbwProfiler.vbwExecuteLine 562
            If .Group = Group And .Flag.Pos > 0 Then
vbwProfiler.vbwExecuteLine 563
                Call LowerRequest(i)
            End If
vbwProfiler.vbwExecuteLine 564 'B
vbwProfiler.vbwExecuteLine 565
        End With
vbwProfiler.vbwExecuteLine 566
    Next i
vbwProfiler.vbwProcOut 35
vbwProfiler.vbwExecuteLine 567
End Function

'Calling Flag must be positioned (Up or Down) before LinkRequest is Called
'If HoistTimer for this Group is running (LastHoist = IdxGroup) dont action Link
'If Queueable (Flags.Queue=True) there should not be a link

Private Function LinkRequest(ByVal Idx As Long)
vbwProfiler.vbwProcIn 36
Dim Lidx As Long
Dim MyLink As defLink

vbwProfiler.vbwExecuteLine 568
    With SignalAttributes(Idx)
vbwProfiler.vbwExecuteLine 569
        If IsLinksInitialised(.Links) Then
vbwProfiler.vbwExecuteLine 570
            For Lidx = 0 To UBound(.Links)
vbwProfiler.vbwExecuteLine 571
                MyLink = .Links(Lidx)
vbwProfiler.vbwExecuteLine 572
                If MyLink.Flag > 0 Then
'If MyLink.Flag = 4 Then Stop
vbwProfiler.vbwExecuteLine 573
                    If .Flag.Pos > 0 And MyLink.Type = "UpLink" Then
vbwProfiler.vbwExecuteLine 574
                        Call LinkExecute(Idx, MyLink)
                    End If
vbwProfiler.vbwExecuteLine 575 'B
vbwProfiler.vbwExecuteLine 576
                    If .Flag.Pos = 0 And MyLink.Type = "DownLink" Then
vbwProfiler.vbwExecuteLine 577
                        Call LinkExecute(Idx, MyLink)
                    End If
vbwProfiler.vbwExecuteLine 578 'B
                End If
vbwProfiler.vbwExecuteLine 579 'B
'Stop 'Link execute can delete a links index which causes a subscript error
'Change for to a loop with mo0re checking
vbwProfiler.vbwExecuteLine 580
            Next Lidx
        End If
vbwProfiler.vbwExecuteLine 581 'B
vbwProfiler.vbwExecuteLine 582
    End With
vbwProfiler.vbwProcOut 36
vbwProfiler.vbwExecuteLine 583
End Function

'IDx is the Signal containing the Link to Link
Private Function LinkExecute(Idx As Long, Link As defLink)
vbwProfiler.vbwProcIn 37
Dim LinkRejected As String
Dim Silent As Boolean

vbwProfiler.vbwExecuteLine 584
        With Link  'Raising Signal

'On ProfileLoad the linked flag may not have been created yet
'.Name is cleared when the Hoist Timer has finished its cycle (5 secs)
vbwProfiler.vbwExecuteLine 585
            If .Flag <> 0 And .Flag <= UBound(SignalAttributes) Then
vbwProfiler.vbwExecuteLine 586
Debug.Print "Link " & SignalAttributes(Idx).Name & " > " & SignalAttributes(.Flag).Name
vbwProfiler.vbwExecuteLine 587
                If SignalAttributes(Idx).Group = LastHoist Then
vbwProfiler.vbwExecuteLine 588
                    LinkRejected = "Suppressed, LastHoist(" & LastHoist & ")"
                End If
vbwProfiler.vbwExecuteLine 589 'B
vbwProfiler.vbwExecuteLine 590
                If SignalAttributes(Idx).Silent = True And SignalAttributes(.Flag).Group = "Sound" Then
vbwProfiler.vbwExecuteLine 591
                    LinkRejected = "Silenced"
                End If
vbwProfiler.vbwExecuteLine 592 'B
vbwProfiler.vbwExecuteLine 593
                If LinkRejected = "" Then
vbwProfiler.vbwExecuteLine 594
                    If .Raise = True Then   'Raise Linked flag
vbwProfiler.vbwExecuteLine 595
                        Call RaiseRequest(.Flag)
                    Else
vbwProfiler.vbwExecuteLine 596 'B
vbwProfiler.vbwExecuteLine 597
                        Call LowerRequest(.Flag)   'Lower Linked flag
                    End If
vbwProfiler.vbwExecuteLine 598 'B
                Else
vbwProfiler.vbwExecuteLine 599 'B
vbwProfiler.vbwExecuteLine 600
Debug.Print LinkRejected
                End If
vbwProfiler.vbwExecuteLine 601 'B
            Else
vbwProfiler.vbwExecuteLine 602 'B
'There are no Linked Flags to this Flag
'Debug.Print "Link " & SignalAttributes(Idx).Name & " > none"
            End If
vbwProfiler.vbwExecuteLine 603 'B
'If a temporary link delete it
vbwProfiler.vbwExecuteLine 604
        If Link.Temp = True Then
vbwProfiler.vbwExecuteLine 605
            Call LinkTempRemove(Idx, Link)
        End If
vbwProfiler.vbwExecuteLine 606 'B
vbwProfiler.vbwExecuteLine 607
        End With
vbwProfiler.vbwProcOut 37
vbwProfiler.vbwExecuteLine 608
End Function

Private Function RC(ByVal Row As Long, ByVal Col As Long) As Long
'Both must be valid as a pair
vbwProfiler.vbwProcIn 38
vbwProfiler.vbwExecuteLine 609
    If Row > 0 And Col > 0 Then
vbwProfiler.vbwExecuteLine 610
        RC = (Row - 1) * 10 + Col
    End If
vbwProfiler.vbwExecuteLine 611 'B
vbwProfiler.vbwProcOut 38
vbwProfiler.vbwExecuteLine 612
End Function

Private Function FlagRow(ByVal Pos As Long) As Long
vbwProfiler.vbwProcIn 39
vbwProfiler.vbwExecuteLine 613
    If Pos > 0 Then
vbwProfiler.vbwExecuteLine 614
        FlagRow = (Pos - 1) \ 10 + 1
    End If
vbwProfiler.vbwExecuteLine 615 'B
vbwProfiler.vbwProcOut 39
vbwProfiler.vbwExecuteLine 616
End Function
    
Private Function FlagCol(ByVal Pos As Long) As Long
vbwProfiler.vbwProcIn 40
vbwProfiler.vbwExecuteLine 617
    If Pos > 0 Then
vbwProfiler.vbwExecuteLine 618
        FlagCol = Pos - (FlagRow(Pos) - 1) * 10
    End If
vbwProfiler.vbwExecuteLine 619 'B
vbwProfiler.vbwProcOut 40
vbwProfiler.vbwExecuteLine 620
End Function

'Called when Raising Flag, SignalAttributes Col & Row = 0 if no Position available
Private Function NextFreeGroupFlagPos(ByVal Idx As Long)
vbwProfiler.vbwProcIn 41
Dim Col As Long
Dim Row As Long
Dim Pos As Long
Dim ClassIdx As Long

'If we do not have a set position see if this flag has a parent
'ie a 2 flag hoist and the parent flag is up

'    Call ResetCols
'If Idx = 9 Then Stop
vbwProfiler.vbwExecuteLine 621
   With SignalAttributes(Idx).Flag
'Get the Column first
vbwProfiler.vbwExecuteLine 622
        If .FixedCol > 0 Then
vbwProfiler.vbwExecuteLine 623
            .Col = .FixedCol
        End If
vbwProfiler.vbwExecuteLine 624 'B

'See if this flag wants placing in same col as the first Class Flag
'DONT REMOVE may want to use it later
vbwProfiler.vbwExecuteLine 625
            If .Col = 0 Then
vbwProfiler.vbwExecuteLine 626
            Select Case SignalAttributes(Idx).Group
'vbwLine 627:            Case Is = "Preparatory", "Shortened"
            Case Is = IIf(vbwProfiler.vbwExecuteLine(627), VBWPROFILER_EMPTY, _
        "Preparatory"), "Shortened"
'Not Recall as next Class flag may be up
vbwProfiler.vbwExecuteLine 628
                   ClassIdx = GroupIdx("Class")
vbwProfiler.vbwExecuteLine 629
                    If ClassIdx > 0 Then
'Put flag in same col
vbwProfiler.vbwExecuteLine 630
Debug.Print "Top Row"
vbwProfiler.vbwExecuteLine 631
                        .Col = SignalAttributes(ClassIdx).Flag.Col
vbwProfiler.vbwExecuteLine 632
                        .Row = Cols(.Col).Items + 1   '1st free row
'                        Call ShiftDown(.Row, .Col)
                    End If
vbwProfiler.vbwExecuteLine 633 'B
            End Select
vbwProfiler.vbwExecuteLine 634 'B
        End If
vbwProfiler.vbwExecuteLine 635 'B

vbwProfiler.vbwExecuteLine 636
        If .Col = 0 Then
'See if we have a flag Raised in this Group with a spare Row available
vbwProfiler.vbwExecuteLine 637
            If Left$(SignalAttributes(Idx).Name, 6) <> "Class " Then
'Class Flags are always in separate cols (Keep in the same group)
vbwProfiler.vbwExecuteLine 638
                For Col = 1 To ColCountFree
vbwProfiler.vbwExecuteLine 639
                    If Cols(Col).Group = SignalAttributes(Idx).Group Then
vbwProfiler.vbwExecuteLine 640
                        If Cols(Col).Items < RowCount Then
vbwProfiler.vbwExecuteLine 641
                            .Col = Col
vbwProfiler.vbwExecuteLine 642
                            Exit For
                        End If
vbwProfiler.vbwExecuteLine 643 'B
                    End If
vbwProfiler.vbwExecuteLine 644 'B
vbwProfiler.vbwExecuteLine 645
                Next Col
            End If
vbwProfiler.vbwExecuteLine 646 'B
        End If
vbwProfiler.vbwExecuteLine 647 'B

'If no Col Group found, get First free col
vbwProfiler.vbwExecuteLine 648
        If .Col = 0 Then
vbwProfiler.vbwExecuteLine 649
            For Col = 1 To ColCountFree
vbwProfiler.vbwExecuteLine 650
                If Cols(Col).Items = 0 Then
'.Group is created by ResetCols
vbwProfiler.vbwExecuteLine 651
                    .Col = Col
vbwProfiler.vbwExecuteLine 652
                    Exit For
                End If
vbwProfiler.vbwExecuteLine 653 'B
vbwProfiler.vbwExecuteLine 654
            Next Col
        End If
vbwProfiler.vbwExecuteLine 655 'B

'If a Class flag see if we can place it in a free column but lower row
'Should only happen on initial load
vbwProfiler.vbwExecuteLine 656
        If .Col = 0 Then
vbwProfiler.vbwExecuteLine 657
            For Row = 2 To RowCount
vbwProfiler.vbwExecuteLine 658
                For Col = 1 To ColCountFree
vbwProfiler.vbwExecuteLine 659
                    If Cols(Col).Group = SignalAttributes(Idx).Group Then
vbwProfiler.vbwExecuteLine 660
                        If Cols(Col).Items < RowCount Then
vbwProfiler.vbwExecuteLine 661
                            .Col = Col
vbwProfiler.vbwExecuteLine 662
                            .Row = Row
vbwProfiler.vbwExecuteLine 663
                            Exit For
                        End If
vbwProfiler.vbwExecuteLine 664 'B
                    End If
vbwProfiler.vbwExecuteLine 665 'B
vbwProfiler.vbwExecuteLine 666
                If .Col > 0 Then
vbwProfiler.vbwExecuteLine 667
                     Exit For
                End If
vbwProfiler.vbwExecuteLine 668 'B
vbwProfiler.vbwExecuteLine 669
                Next Col
vbwProfiler.vbwExecuteLine 670
            If .Col > 0 Then
vbwProfiler.vbwExecuteLine 671
                 Exit For
            End If
vbwProfiler.vbwExecuteLine 672 'B
vbwProfiler.vbwExecuteLine 673
            Next Row
        End If
vbwProfiler.vbwExecuteLine 674 'B

'On initial load place in any free slot
vbwProfiler.vbwExecuteLine 675
        If .Col = 0 Then
vbwProfiler.vbwExecuteLine 676
            For Row = 1 To RowCount
vbwProfiler.vbwExecuteLine 677
                For Col = 1 To ColCount
vbwProfiler.vbwExecuteLine 678
                    If Cols(Col).Items < RowCount Then
vbwProfiler.vbwExecuteLine 679
                        .Col = Col
vbwProfiler.vbwExecuteLine 680
                        .Row = Cols(Col).Items + 1
vbwProfiler.vbwExecuteLine 681
                        Exit For
                    End If
vbwProfiler.vbwExecuteLine 682 'B
vbwProfiler.vbwExecuteLine 683
                If .Col > 0 Then
vbwProfiler.vbwExecuteLine 684
                     Exit For
                End If
vbwProfiler.vbwExecuteLine 685 'B
vbwProfiler.vbwExecuteLine 686
                Next Col
vbwProfiler.vbwExecuteLine 687
            If .Col > 0 Then
vbwProfiler.vbwExecuteLine 688
                 Exit For
            End If
vbwProfiler.vbwExecuteLine 689 'B
vbwProfiler.vbwExecuteLine 690
            Next Row
        End If
vbwProfiler.vbwExecuteLine 691 'B

vbwProfiler.vbwExecuteLine 692
        If .Col = 0 Then
vbwProfiler.vbwExecuteLine 693
MsgBox "No free Cols", vbCritical, "NextFreeGroupFlagPos"
vbwProfiler.vbwProcOut 41
vbwProfiler.vbwExecuteLine 694
            Exit Function
        End If
vbwProfiler.vbwExecuteLine 695 'B

vbwProfiler.vbwExecuteLine 696
        If .Row = 0 Then
vbwProfiler.vbwExecuteLine 697
            If .FixedRow > 0 Then
vbwProfiler.vbwExecuteLine 698
                .Row = .FixedRow
            Else
vbwProfiler.vbwExecuteLine 699 'B
vbwProfiler.vbwExecuteLine 700
                .Row = Cols(.Col).Items + 1
            End If
vbwProfiler.vbwExecuteLine 701 'B
        End If
vbwProfiler.vbwExecuteLine 702 'B
vbwProfiler.vbwExecuteLine 703
    End With
'Debug.Print "NextPos=" & NextFreeGroupFlagPos & " (" & Row & "," & Col & ")"
vbwProfiler.vbwProcOut 41
vbwProfiler.vbwExecuteLine 704
End Function

Private Function DequeCmd(Optional Group As String) As Integer
vbwProfiler.vbwProcIn 42
Dim i As Long
vbwProfiler.vbwExecuteLine 705
    For i = 0 To UBound(CmdQ)
vbwProfiler.vbwExecuteLine 706
        If CmdQ(i) <> 0 Then
vbwProfiler.vbwExecuteLine 707
            If Group = "" Or SignalAttributes(CmdQ(i)).Group = Group Then
vbwProfiler.vbwExecuteLine 708
                If DequeCmd = 0 Then
vbwProfiler.vbwExecuteLine 709
                    DequeCmd = CmdQ(i)
vbwProfiler.vbwExecuteLine 710
Debug.Print "Deque " & SignalAttributes(CmdQ(i)).Name & " (" & Group & ")"
vbwProfiler.vbwExecuteLine 711
                    Commands(CmdQ(i)).BackColor = cbDefault
vbwProfiler.vbwExecuteLine 712
                    CmdQ(i) = 0
                End If
vbwProfiler.vbwExecuteLine 713 'B
            End If
vbwProfiler.vbwExecuteLine 714 'B
        End If
vbwProfiler.vbwExecuteLine 715 'B
'Shift remaining commands up the queue
vbwProfiler.vbwExecuteLine 716
        If DequeCmd <> 0 Then
vbwProfiler.vbwExecuteLine 717
            If i = UBound(CmdQ) Then
vbwProfiler.vbwExecuteLine 718
                CmdQ(i) = 0
            Else
vbwProfiler.vbwExecuteLine 719 'B
vbwProfiler.vbwExecuteLine 720
                CmdQ(i) = CmdQ(i + 1)
            End If
vbwProfiler.vbwExecuteLine 721 'B
        End If
vbwProfiler.vbwExecuteLine 722 'B
vbwProfiler.vbwExecuteLine 723
    Next i

'Stop
vbwProfiler.vbwProcOut 42
vbwProfiler.vbwExecuteLine 724
End Function

Private Function QueueRequest(Idx As Long)
vbwProfiler.vbwProcIn 43
Dim i As Long

vbwProfiler.vbwExecuteLine 725
    For i = 0 To UBound(CmdQ)
vbwProfiler.vbwExecuteLine 726
        If CmdQ(i) = 0 Then
vbwProfiler.vbwExecuteLine 727
            CmdQ(i) = Idx
vbwProfiler.vbwExecuteLine 728
            Commands(Idx).BackColor = vbCyan
vbwProfiler.vbwExecuteLine 729
Debug.Print "Queue " & SignalAttributes(CmdQ(i)).Name
vbwProfiler.vbwProcOut 43
vbwProfiler.vbwExecuteLine 730
            Exit Function
        Else
vbwProfiler.vbwExecuteLine 731 'B
'Only q the same command once (must not queue Recall more than once)
vbwProfiler.vbwExecuteLine 732
            If CmdQ(i) = Idx Then
vbwProfiler.vbwProcOut 43
vbwProfiler.vbwExecuteLine 733
                 Exit Function
            End If
vbwProfiler.vbwExecuteLine 734 'B
        End If
vbwProfiler.vbwExecuteLine 735 'B
vbwProfiler.vbwExecuteLine 736
    Next i
'MsgBox "Command Queue is full (" & UBound(CmdQ) & ") maximum"
vbwProfiler.vbwProcOut 43
vbwProfiler.vbwExecuteLine 737
End Function

Private Function FinishTime()
vbwProfiler.vbwProcIn 44
vbwProfiler.vbwExecuteLine 738
    With mshFinish
'not the first (blank) row
vbwProfiler.vbwExecuteLine 739
        If .TextMatrix(.Rows - 1, 0) <> "" Then
vbwProfiler.vbwExecuteLine 740
            .Rows = .Rows + 1
        End If
vbwProfiler.vbwExecuteLine 741 'B
vbwProfiler.vbwExecuteLine 742
        .TextMatrix(.Rows - 1, 0) = .Rows - 1
vbwProfiler.vbwExecuteLine 743
        .TextMatrix(.Rows - 1, 1) = lblCurrTime.Caption
'Scroll to bottom
vbwProfiler.vbwExecuteLine 744
        .TopRow = .Rows - 1
vbwProfiler.vbwExecuteLine 745
End With

vbwProfiler.vbwProcOut 44
vbwProfiler.vbwExecuteLine 746
End Function

Private Function FlagVisibility(ByVal Idx As Long, Visible As Boolean)
vbwProfiler.vbwProcIn 45
Dim Pos As Long
Dim Cidx As Long
vbwProfiler.vbwExecuteLine 747
    Pos = SignalAttributes(Idx).Flag.Pos
'See if visiblility has changed (To generate Controller event)
vbwProfiler.vbwExecuteLine 748
    If Pos > 0 Then
vbwProfiler.vbwExecuteLine 749
        If Flags(Pos).Visible <> Visible Then
vbwProfiler.vbwExecuteLine 750
            Flags(Pos).Visible = Visible
vbwProfiler.vbwExecuteLine 751
            Cidx = SignalAttributes(Idx).Controller
vbwProfiler.vbwExecuteLine 752
            If Cidx <> -1 Then
vbwProfiler.vbwExecuteLine 753
                With Controllers(Cidx)
vbwProfiler.vbwExecuteLine 754
                    If Visible Then
vbwProfiler.vbwExecuteLine 755
Debug.Print .Connection & "(" & Cidx & ")" & .On
vbwProfiler.vbwExecuteLine 756
                        If .Sound <> "" Then
vbwProfiler.vbwExecuteLine 757
                             Call PlayWav(.Sound)
                        End If
vbwProfiler.vbwExecuteLine 758 'B
'Call Beep(300, CInt(SignalAttributes(Idx).TTL))
                    Else
vbwProfiler.vbwExecuteLine 759 'B
vbwProfiler.vbwExecuteLine 760
Debug.Print .Connection & "(" & Cidx & ")" & .Off
vbwProfiler.vbwExecuteLine 761
                        If .Sound <> "" Then
vbwProfiler.vbwExecuteLine 762
                             Call StopWav
                        End If
vbwProfiler.vbwExecuteLine 763 'B
                    End If
vbwProfiler.vbwExecuteLine 764 'B
vbwProfiler.vbwExecuteLine 765
                End With
            End If
vbwProfiler.vbwExecuteLine 766 'B
        End If
vbwProfiler.vbwExecuteLine 767 'B
    Else
vbwProfiler.vbwExecuteLine 768 'B
vbwProfiler.vbwExecuteLine 769
        MsgBox "Flag " & SignalAttributes(Idx).Name & " not Raised", vbCritical, "FlagVisibility"
    End If
vbwProfiler.vbwExecuteLine 770 'B
vbwProfiler.vbwProcOut 45
vbwProfiler.vbwExecuteLine 771
End Function

Private Function ResetCols()
vbwProfiler.vbwProcIn 46
Dim Idx As Long
Dim Col As Long

vbwProfiler.vbwExecuteLine 772
    ReDim Cols(ColCount)
vbwProfiler.vbwExecuteLine 773
    For Idx = 1 To UBound(SignalAttributes)
vbwProfiler.vbwExecuteLine 774
        With SignalAttributes(Idx)
vbwProfiler.vbwExecuteLine 775
            If .Flag.Col > 0 And .Flag.Row = 1 Then
vbwProfiler.vbwExecuteLine 776
                Cols(.Flag.Col).Group = .Group
            End If
vbwProfiler.vbwExecuteLine 777 'B
vbwProfiler.vbwExecuteLine 778
            If .Flag.FixedCol > 0 Then
vbwProfiler.vbwExecuteLine 779
                Cols(.Flag.FixedCol).Group = .Group
            End If
vbwProfiler.vbwExecuteLine 780 'B
vbwProfiler.vbwExecuteLine 781
            If SignalAttributes(Idx).Flag.Col > 0 Then
vbwProfiler.vbwExecuteLine 782
                Cols(.Flag.Col).Items = Cols(.Flag.Col).Items + 1
            End If
vbwProfiler.vbwExecuteLine 783 'B
vbwProfiler.vbwExecuteLine 784
        End With
vbwProfiler.vbwExecuteLine 785
    Next Idx
vbwProfiler.vbwProcOut 46
vbwProfiler.vbwExecuteLine 786
End Function

'Used to Check if a Class Flag is up when Recall is asked for
'If 2 Class flags are up it will select the lowest class (Idx is in class order)
Private Function GroupIdx(ByVal Group As String) As Long
vbwProfiler.vbwProcIn 47
Dim Idx As Long
vbwProfiler.vbwExecuteLine 787
    For Idx = 1 To UBound(SignalAttributes)
vbwProfiler.vbwExecuteLine 788
        With SignalAttributes(Idx)
vbwProfiler.vbwExecuteLine 789
            If .Group = Group And .Flag.Pos > 0 Then
vbwProfiler.vbwExecuteLine 790
                 GroupIdx = Idx
vbwProfiler.vbwExecuteLine 791
                Exit For
            End If
vbwProfiler.vbwExecuteLine 792 'B
vbwProfiler.vbwExecuteLine 793
        End With
vbwProfiler.vbwExecuteLine 794
    Next Idx
vbwProfiler.vbwProcOut 47
vbwProfiler.vbwExecuteLine 795
End Function


Private Function CompressCols()
vbwProfiler.vbwProcIn 48
Dim LowestFixedCol As Long
Dim Idx As Long
Dim Col As Long
Dim Row As Long
Dim Pos As Long
Dim PosFrom As Long
Dim PosTo As Long
'Ensure Cols() is correct
vbwProfiler.vbwExecuteLine 796
    Call ResetCols
vbwProfiler.vbwExecuteLine 797
    For Idx = 1 To UBound(SignalAttributes)
vbwProfiler.vbwExecuteLine 798
        With SignalAttributes(Idx).Flag
vbwProfiler.vbwExecuteLine 799
            If .FixedCol > 0 Then
vbwProfiler.vbwExecuteLine 800
                If LowestFixedCol = 0 Then
vbwProfiler.vbwExecuteLine 801
                     LowestFixedCol = .FixedCol
                End If
vbwProfiler.vbwExecuteLine 802 'B
vbwProfiler.vbwExecuteLine 803
                If LowestFixedCol > .FixedCol Then
vbwProfiler.vbwExecuteLine 804
                    LowestFixedCol = .FixedCol
                End If
vbwProfiler.vbwExecuteLine 805 'B
            End If
vbwProfiler.vbwExecuteLine 806 'B
vbwProfiler.vbwExecuteLine 807
        End With
vbwProfiler.vbwExecuteLine 808
    Next Idx
'Exit Function
'Stop

vbwProfiler.vbwExecuteLine 809
    For Col = 1 To LowestFixedCol - 2
vbwProfiler.vbwExecuteLine 810
        If Cols(Col).Items = 0 And Cols(Col + 1).Items > 0 Then
vbwProfiler.vbwExecuteLine 811
            For Idx = 1 To UBound(SignalAttributes)
vbwProfiler.vbwExecuteLine 812
                With SignalAttributes(Idx).Flag
vbwProfiler.vbwExecuteLine 813
                    If .Col = Col + 1 Then
'Move Flags(pos).Picture
vbwProfiler.vbwExecuteLine 814
                        For Row = 1 To RowCount
'GetPos of Empty Col
vbwProfiler.vbwExecuteLine 815
                            PosTo = RC(Row, Col)
vbwProfiler.vbwExecuteLine 816
                            PosFrom = RC(Row, Col + 1)
'                            Flags(CInt(Pos)).Picture = Flags(CInt(.Pos)).Picture
vbwProfiler.vbwExecuteLine 817
                            Flags(CInt(PosTo)) = Flags(CInt(PosFrom))
vbwProfiler.vbwExecuteLine 818
                            Flags(CInt(PosTo)).Visible = Flags(CInt(PosFrom)).Visible

vbwProfiler.vbwExecuteLine 819
                            Flags(CInt(PosFrom)).Picture = Nothing
vbwProfiler.vbwExecuteLine 820
                            Flags(CInt(PosFrom)).Visible = False
vbwProfiler.vbwExecuteLine 821
                            .Row = Row
vbwProfiler.vbwExecuteLine 822
                            .Col = Col
vbwProfiler.vbwExecuteLine 823
                            .Pos = PosTo
'Stop

'Reset Pos,Col,Row
vbwProfiler.vbwExecuteLine 824
                        Next Row
                    End If
vbwProfiler.vbwExecuteLine 825 'B
vbwProfiler.vbwExecuteLine 826
                End With
vbwProfiler.vbwExecuteLine 827
            Next Idx
        End If
vbwProfiler.vbwExecuteLine 828 'B
vbwProfiler.vbwExecuteLine 829
    Next Col
vbwProfiler.vbwExecuteLine 830
    Call ResetCols
vbwProfiler.vbwProcOut 48
vbwProfiler.vbwExecuteLine 831
End Function

'Shift This Column Down 1 from this Row
Private Function ShiftDown(ByVal Row As Long, ByVal Col As Long)
vbwProfiler.vbwProcIn 49
Dim Idx As Long
Dim PosFrom As Long
Dim PosTo As Long
vbwProfiler.vbwExecuteLine 832
    If Cols(Col).Items = RowCount Then
vbwProfiler.vbwExecuteLine 833
MsgBox "No Free Rows"
    End If
vbwProfiler.vbwExecuteLine 834 'B
vbwProfiler.vbwExecuteLine 835
    For Row = Cols(Col).Items To Row Step -1
vbwProfiler.vbwExecuteLine 836
        PosFrom = RC(Row, Col)
vbwProfiler.vbwExecuteLine 837
        For Idx = 1 To UBound(SignalAttributes)
vbwProfiler.vbwExecuteLine 838
            With SignalAttributes(Idx).Flag
vbwProfiler.vbwExecuteLine 839
                If .Pos = PosFrom Then
vbwProfiler.vbwExecuteLine 840
                    PosTo = RC(Row + 1, Col)
vbwProfiler.vbwExecuteLine 841
                    Flags(CInt(PosTo)) = Flags(CInt(PosFrom))
vbwProfiler.vbwExecuteLine 842
                    Flags(CInt(PosTo)).Visible = Flags(CInt(PosFrom)).Visible
vbwProfiler.vbwExecuteLine 843
                    Flags(CInt(PosFrom)).Picture = Nothing
vbwProfiler.vbwExecuteLine 844
                    Flags(CInt(PosFrom)).Visible = False
'Change Flag Position on SignalAttributes of flag were moving
vbwProfiler.vbwExecuteLine 845
                    .Row = Row + 1
vbwProfiler.vbwExecuteLine 846
                    .Col = Col
vbwProfiler.vbwExecuteLine 847
                    .Pos = PosTo
                End If
vbwProfiler.vbwExecuteLine 848 'B
vbwProfiler.vbwExecuteLine 849
            End With
vbwProfiler.vbwExecuteLine 850
        Next Idx
vbwProfiler.vbwExecuteLine 851
    Next Row

vbwProfiler.vbwProcOut 49
vbwProfiler.vbwExecuteLine 852
End Function
Private Function LinkTempRemove(ByVal Idx As Long, Link As defLink)
vbwProfiler.vbwProcIn 50
Dim Lidx As Long
Dim i As Long
vbwProfiler.vbwExecuteLine 853
    With SignalAttributes(Idx)
vbwProfiler.vbwExecuteLine 854
        If IsArrayInitialised(.Links) Then
vbwProfiler.vbwExecuteLine 855
            For Lidx = 0 To UBound(.Links)
vbwProfiler.vbwExecuteLine 856
                If .Links(Lidx).Temp = True _
                And .Links(Lidx).Flag = Link.Flag _
                And .Links(Lidx).Raise = Link.Raise _
                And .Links(Lidx).Type = Link.Type Then
'Shift down, if not at the bottom of the array
vbwProfiler.vbwExecuteLine 857
                    For i = Lidx To UBound(.Links) - 1
vbwProfiler.vbwExecuteLine 858
                        .Links(i) = .Links(i + 1)
vbwProfiler.vbwExecuteLine 859
                    Next i
'If a link has been removed the Redim the array, withoust the last element
vbwProfiler.vbwExecuteLine 860
                    ReDim Preserve .Links(UBound(.Links) - 1)
vbwProfiler.vbwProcOut 50
vbwProfiler.vbwExecuteLine 861
                Exit Function
                End If
vbwProfiler.vbwExecuteLine 862 'B
vbwProfiler.vbwExecuteLine 863
            Next Lidx
        End If
vbwProfiler.vbwExecuteLine 864 'B
vbwProfiler.vbwExecuteLine 865
    End With
vbwProfiler.vbwProcOut 50
vbwProfiler.vbwExecuteLine 866
End Function

'Set Command Button Visibility
'Called by DoTimerEvents
Public Function CommandButtonVisibility(ElapsedTime As Long)
'If Timer Events have started and not finished you cannot Postpone
vbwProfiler.vbwProcIn 51
vbwProfiler.vbwExecuteLine 867
        Select Case ElapsedTime
'vbwLine 868:        Case Is >= Myprofile.LastEvent.Second
        Case Is >= IIf(vbwProfiler.vbwExecuteLine(868), VBWPROFILER_EMPTY, _
        Myprofile.LastEvent.Second)
'Start Sequence finished
'Only do if recall timer not running, if so do when recalltimer has finished
vbwProfiler.vbwExecuteLine 869
            With Commands(CommandFromCaption("Finish"))
vbwProfiler.vbwExecuteLine 870
                .BackColor = vbGreen
vbwProfiler.vbwExecuteLine 871
                .Enabled = True
vbwProfiler.vbwExecuteLine 872
                .SetFocus
vbwProfiler.vbwExecuteLine 873
            End With
'vbwLine 874:        Case Is >= Myprofile.FirstEvent.Second
        Case Is >= IIf(vbwProfiler.vbwExecuteLine(874), VBWPROFILER_EMPTY, _
        Myprofile.FirstEvent.Second)
'Start Sequence Running
vbwProfiler.vbwExecuteLine 875
            With Commands(CommandFromCaption("Postpone"))
vbwProfiler.vbwExecuteLine 876
                .BackColor = cbDefault
vbwProfiler.vbwExecuteLine 877
                .Enabled = False
vbwProfiler.vbwExecuteLine 878
            End With
vbwProfiler.vbwExecuteLine 879
            With Commands(CommandFromCaption("Recall"))
vbwProfiler.vbwExecuteLine 880
                .BackColor = vbGreen
vbwProfiler.vbwExecuteLine 881
                .Enabled = True
vbwProfiler.vbwExecuteLine 882
                .SetFocus
vbwProfiler.vbwExecuteLine 883
            End With
        Case Else
vbwProfiler.vbwExecuteLine 884 'B
'Start Sequence not started
vbwProfiler.vbwExecuteLine 885
            With Commands(CommandFromCaption("Postpone"))
vbwProfiler.vbwExecuteLine 886
                .BackColor = vbGreen
vbwProfiler.vbwExecuteLine 887
                .Enabled = True
vbwProfiler.vbwExecuteLine 888
                .SetFocus
vbwProfiler.vbwExecuteLine 889
            End With
vbwProfiler.vbwExecuteLine 890
            With Commands(CommandFromCaption("Recall"))
vbwProfiler.vbwExecuteLine 891
                .BackColor = cbDefault
vbwProfiler.vbwExecuteLine 892
                .Enabled = False
vbwProfiler.vbwExecuteLine 893
            End With
vbwProfiler.vbwExecuteLine 894
            With Commands(CommandFromCaption("General Recall"))
vbwProfiler.vbwExecuteLine 895
                .BackColor = cbDefault
vbwProfiler.vbwExecuteLine 896
                .Enabled = False
vbwProfiler.vbwExecuteLine 897
            End With
vbwProfiler.vbwExecuteLine 898
            With Commands(CommandFromCaption("Finish"))
vbwProfiler.vbwExecuteLine 899
                .BackColor = cbDefault
vbwProfiler.vbwExecuteLine 900
                .Enabled = False
vbwProfiler.vbwExecuteLine 901
            End With
        End Select
vbwProfiler.vbwExecuteLine 902 'B

vbwProfiler.vbwProcOut 51
vbwProfiler.vbwExecuteLine 903
End Function

'Return the Command Button IDX, as we should find it within the first 6 buttons (Fixed)
Private Function CommandFromCaption(ByVal CbName As String) As Integer
vbwProfiler.vbwProcIn 52
Dim Index As Integer
vbwProfiler.vbwExecuteLine 904
    For Index = 1 To Commands.Count
vbwProfiler.vbwExecuteLine 905
        If Commands(Index).Caption = CbName Then
vbwProfiler.vbwExecuteLine 906
            CommandFromCaption = Index
vbwProfiler.vbwProcOut 52
vbwProfiler.vbwExecuteLine 907
            Exit Function
        End If
vbwProfiler.vbwExecuteLine 908 'B
vbwProfiler.vbwExecuteLine 909
    Next Index
vbwProfiler.vbwExecuteLine 910
Stop
vbwProfiler.vbwProcOut 52
vbwProfiler.vbwExecuteLine 911
End Function


Public Function DoTimerEvents(ElapsedTime As Long)
vbwProfiler.vbwProcIn 53
Dim Eidx As Long
Dim Sidx As Long
Dim Bidx As Long
vbwProfiler.vbwExecuteLine 912
    For Eidx = 0 To UBound(Evts)
vbwProfiler.vbwExecuteLine 913
        If Evts(Eidx).ElapsedTime = ElapsedTime Then
vbwProfiler.vbwExecuteLine 914
            If IsSignalsInitialised(Evts(Eidx).Signals) Then
vbwProfiler.vbwExecuteLine 915
                For Sidx = 0 To UBound(Evts(Eidx).Signals)
vbwProfiler.vbwExecuteLine 916
                    With Evts(Eidx).Signals(Sidx)
'Silent on SignalAttribues is only used by LinkRquest and is set temporarily
'for this call only
vbwProfiler.vbwExecuteLine 917
                        If .Silent = "True" Then
vbwProfiler.vbwExecuteLine 918
                             SignalAttributes(.Signal).Silent = True
                        End If
vbwProfiler.vbwExecuteLine 919 'B
vbwProfiler.vbwExecuteLine 920
                        If .Raise = True Then
vbwProfiler.vbwExecuteLine 921
                            If SignalAttributes(.Signal).Flag.Pos = 0 Then
vbwProfiler.vbwExecuteLine 922
                                Call frmMain.RaiseRequest(.Signal)
                            End If
vbwProfiler.vbwExecuteLine 923 'B
                        Else
vbwProfiler.vbwExecuteLine 924 'B
'Lower Recall is aked for even when not up
vbwProfiler.vbwExecuteLine 925
                            If SignalAttributes(.Signal).Flag.Pos > 0 Then
vbwProfiler.vbwExecuteLine 926
                                Call frmMain.LowerRequest(.Signal)
                            End If
vbwProfiler.vbwExecuteLine 927 'B
                        End If
vbwProfiler.vbwExecuteLine 928 'B
vbwProfiler.vbwExecuteLine 929
                        SignalAttributes(.Signal).Silent = False
vbwProfiler.vbwExecuteLine 930
                    End With
vbwProfiler.vbwExecuteLine 931
                Next Sidx
            End If
vbwProfiler.vbwExecuteLine 932 'B
vbwProfiler.vbwExecuteLine 933
            If IsButtonsInitialised(Evts(Eidx).Buttons) Then
vbwProfiler.vbwExecuteLine 934
                For Bidx = 0 To UBound(Evts(Eidx).Buttons)
vbwProfiler.vbwExecuteLine 935
                    With Evts(Eidx).Buttons(Bidx)
'Stop
'The Command button properties can be set immediately if any
vbwProfiler.vbwExecuteLine 936
                        If .Enabled <> "" Then
vbwProfiler.vbwExecuteLine 937
                            Commands(.Button).Enabled = AtoBool(.Enabled)
                        End If
vbwProfiler.vbwExecuteLine 938 'B
vbwProfiler.vbwExecuteLine 939
                    End With
vbwProfiler.vbwExecuteLine 940
                Next Bidx
            End If
vbwProfiler.vbwExecuteLine 941 'B
vbwProfiler.vbwExecuteLine 942
            If Evts(Eidx).Focus > 0 Then
'Must be enabled to put focus on it
vbwProfiler.vbwExecuteLine 943
                Commands(Evts(Eidx).Focus).Enabled = True
vbwProfiler.vbwExecuteLine 944
                Commands(Evts(Eidx).Focus).SetFocus
vbwProfiler.vbwExecuteLine 945
                Commands(Evts(Eidx).Focus).BackColor = vbGreen
            End If
vbwProfiler.vbwExecuteLine 946 'B
        End If  'This Event
vbwProfiler.vbwExecuteLine 947 'B
vbwProfiler.vbwExecuteLine 948
        If Evts(Eidx).ElapsedTime > ElapsedTime Then
vbwProfiler.vbwProcOut 53
vbwProfiler.vbwExecuteLine 949
            Exit Function
        End If  'This event not found
vbwProfiler.vbwExecuteLine 950 'B
vbwProfiler.vbwExecuteLine 951
    Next Eidx


#If False Then
    
    If Myprofile.IsEventDue(ElapsedTime) = False Then
'If ElapsedTime = 0 Then Stop
        Exit Function
    End If


'    For Each MyEvent In Myprofile
'         If MyEvent.Second = ElapsedTime Then
'            Call DoEvent(MyEvent)
'        End If
'    Next MyEvent

    If Myprofile.LastEvent.Second = ElapsedTime Then
        Call ResetEvents
    End If
'    Call frmMain.CommandButtonVisibility(ElapsedTime)
#End If
vbwProfiler.vbwProcOut 53
vbwProfiler.vbwExecuteLine 952
End Function

#If False Then
Public Function DoEvent(EventNo As Long)

Stop

            
'Check if Signal requires changing
            If SignalAttributes(CurrEvent.Signal).Flag.Pos > 0 Then
                Raised = True
            End If

'Do not generate Sound Signal for this event (ClassRecall)
'            If CurrEvent.Raised <> Raised Then
                If CurrEvent.Raised = True Then
                    If SignalAttributes(CurrEvent.Signal).Flag.Pos = 0 Then
                        Call frmMain.RaiseRequest(CurrEvent.Signal)
                    Else
'Flag already up
'Stop
                    End If
                Else
                    If SignalAttributes(CurrEvent.Signal).Flag.Pos > 0 Then
                        Call frmMain.LowerRequest(CurrEvent.Signal)
                    Else
'Flag already down (can happen because both Recall and General Recall have a down event
'even if not up)
'Stop
                    End If
                End If
'            End If

'Change Command Buttton enabled (if required)
'            If frmMain.Commands(CInt(CurrEvent.Signal)).Enabled <> CurrEvent.CommandEnabled Then
'Stop
'                frmMain.Commands(CInt(CurrEvent.Signal)).Enabled = CurrEvent.CommandEnabled
'            End If
        

End Function
#End If

#If False Then
Public Function DoEvent_old(MyEvent As clsEvent)
Dim Raised As Boolean
        
            Set CurrEvent = MyEvent
Debug.Print CurrEvent.Second & " " & CurrEvent.Signal & " " & CurrEvent.Raised
            
'Check if Signal requires changing
            If SignalAttributes(CurrEvent.Signal).Flag.Pos > 0 Then
                Raised = True
            End If

'Do not generate Sound Signal for this event (ClassRecall)
'            If CurrEvent.Raised <> Raised Then
                If CurrEvent.Raised = True Then
                    If SignalAttributes(CurrEvent.Signal).Flag.Pos = 0 Then
                        Call frmMain.RaiseRequest(CurrEvent.Signal)
                    Else
'Flag already up
'Stop
                    End If
                Else
                    If SignalAttributes(CurrEvent.Signal).Flag.Pos > 0 Then
                        Call frmMain.LowerRequest(CurrEvent.Signal)
                    Else
'Flag already down (can happen because both Recall and General Recall have a down event
'even if not up)
'Stop
                    End If
                End If
'            End If

'Change Command Buttton enabled (if required)
'            If frmMain.Commands(CInt(CurrEvent.Signal)).Enabled <> CurrEvent.CommandEnabled Then
'Stop
'                frmMain.Commands(CInt(CurrEvent.Signal)).Enabled = CurrEvent.CommandEnabled
'            End If
        

End Function
#End If


