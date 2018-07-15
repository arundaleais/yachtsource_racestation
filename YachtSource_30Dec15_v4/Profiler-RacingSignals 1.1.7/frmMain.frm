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
Private DebugFlags As Boolean
Private DebugHoist As Boolean
Private DebugSignalTimer As Boolean
Private DebugQueue As Boolean
Private DebugLink As Boolean
Private DebugConnection As Boolean
Private DebugRecall As Boolean

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

Private Sub LoadSequence()
vbwProfiler.vbwProcIn 2
Dim i As Long

'Load startsequencies - only on initial startup
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

Private Sub cmdEvents_Click()
vbwProfiler.vbwProcIn 4
vbwProfiler.vbwExecuteLine 40
    If frmEvents.Visible Then
vbwProfiler.vbwExecuteLine 41
        frmEvents.Visible = False
    Else
vbwProfiler.vbwExecuteLine 42 'B
vbwProfiler.vbwExecuteLine 43
        frmEvents.Visible = True
    End If
vbwProfiler.vbwExecuteLine 44 'B
vbwProfiler.vbwProcOut 4
vbwProfiler.vbwExecuteLine 45
End Sub

Private Sub Commands_Click(Index As Integer)
vbwProfiler.vbwProcIn 5
Dim Position As Long
Dim NextCommand As Long

vbwProfiler.vbwExecuteLine 46
Debug.Print "--- " & Commands(Index).Caption & " ---"

vbwProfiler.vbwExecuteLine 47
    With SignalAttributes(Index)
'If this command is queued then just remove it (same as clicking when up)
'This must be done in the Click event because the user is making the request
'You cannot do it in RaiseRequest or LowerRequest because all queued events
'would get removed.
vbwProfiler.vbwExecuteLine 48
        If Commands(Index).BackColor = vbCyan Then
vbwProfiler.vbwExecuteLine 49
            NextCommand = DequeCmd(.Group)
vbwProfiler.vbwProcOut 5
vbwProfiler.vbwExecuteLine 50
            Exit Sub
        End If
vbwProfiler.vbwExecuteLine 51 'B
'If we have another commandButton queued in this group, remove this before
'actioning a raise request so we dont have 2 flags in same group queued
'This is important with Recall & General Recall
vbwProfiler.vbwExecuteLine 52
        If .Flag.Pos = 0 Then
vbwProfiler.vbwExecuteLine 53
            NextCommand = DequeCmd(.Group)
vbwProfiler.vbwExecuteLine 54
            Call RaiseRequest(CLng(Index))
        Else
vbwProfiler.vbwExecuteLine 55 'B
vbwProfiler.vbwExecuteLine 56
            Call LowerRequest(CLng(Index))
        End If
vbwProfiler.vbwExecuteLine 57 'B
vbwProfiler.vbwExecuteLine 58
    End With
vbwProfiler.vbwProcOut 5
vbwProfiler.vbwExecuteLine 59
End Sub


Private Sub Flags_Click(Index As Integer)
'MsgBox Flags(Index).Picture.Handle
vbwProfiler.vbwProcIn 6
vbwProfiler.vbwProcOut 6
vbwProfiler.vbwExecuteLine 60
End Sub

Private Sub Form_Load()
vbwProfiler.vbwProcIn 7
Dim i As Long
Dim url As String
Dim Major As Long
Dim Minor As Long
Dim Revision As Long
Dim NewVersion As Boolean

vbwProfiler.vbwExecuteLine 61
DebugQueue = True
vbwProfiler.vbwExecuteLine 62
    Caption = App.EXEName & " [" & App.Major & "." & App.Minor & "." _
    & App.Revision & "] "

'Check if a later version exists
vbwProfiler.vbwExecuteLine 63
    url = "http://www.NmeaRouter.com/docs/ais/" & App.EXEName _
    & "_setup_"
vbwProfiler.vbwExecuteLine 64
    Major = App.Major
vbwProfiler.vbwExecuteLine 65
    Do
vbwProfiler.vbwExecuteLine 66
        If HTTPFileExists(url & Major & ".0.0.exe") = False Then
vbwProfiler.vbwExecuteLine 67
             Exit Do
        End If
vbwProfiler.vbwExecuteLine 68 'B
vbwProfiler.vbwExecuteLine 69
        Major = Major + 1
vbwProfiler.vbwExecuteLine 70
    Loop
vbwProfiler.vbwExecuteLine 71
    If Major > 0 Then 'Highest major that exists
vbwProfiler.vbwExecuteLine 72
         Major = Major - 1
    End If
vbwProfiler.vbwExecuteLine 73 'B

vbwProfiler.vbwExecuteLine 74
    url = url & Major & "."
vbwProfiler.vbwExecuteLine 75
    If Major = App.Major Then
vbwProfiler.vbwExecuteLine 76
        Minor = App.Minor
    Else
vbwProfiler.vbwExecuteLine 77 'B
vbwProfiler.vbwExecuteLine 78
        Minor = 0
    End If
vbwProfiler.vbwExecuteLine 79 'B
vbwProfiler.vbwExecuteLine 80
    Do
vbwProfiler.vbwExecuteLine 81
        If HTTPFileExists(url & Minor & ".0.exe") = False Then
vbwProfiler.vbwExecuteLine 82
             Exit Do
        End If
vbwProfiler.vbwExecuteLine 83 'B
vbwProfiler.vbwExecuteLine 84
        Minor = Minor + 1
vbwProfiler.vbwExecuteLine 85
    Loop
vbwProfiler.vbwExecuteLine 86
    If Minor > 0 Then
vbwProfiler.vbwExecuteLine 87
         Minor = Minor - 1
    End If
vbwProfiler.vbwExecuteLine 88 'B

vbwProfiler.vbwExecuteLine 89
    url = url & Minor & "."
vbwProfiler.vbwExecuteLine 90
    If Not (Major = App.Major And Minor = App.Minor) Then
vbwProfiler.vbwExecuteLine 91
        NewVersion = True
    End If
vbwProfiler.vbwExecuteLine 92 'B
'Only let a user get next revision if he is using a revision
'of his current version. Otherwise he goes up to the next minor version
vbwProfiler.vbwExecuteLine 93
    If NewVersion = False And App.Revision > 0 Then
vbwProfiler.vbwExecuteLine 94
        Revision = App.Revision
vbwProfiler.vbwExecuteLine 95
        Do
vbwProfiler.vbwExecuteLine 96
            If HTTPFileExists(url & Revision & ".exe") = False Then
vbwProfiler.vbwExecuteLine 97
                 Exit Do
            End If
vbwProfiler.vbwExecuteLine 98 'B
vbwProfiler.vbwExecuteLine 99
            Revision = Revision + 1
vbwProfiler.vbwExecuteLine 100
        Loop
vbwProfiler.vbwExecuteLine 101
        If Revision > 0 Then
vbwProfiler.vbwExecuteLine 102
             Revision = Revision - 1
        End If
vbwProfiler.vbwExecuteLine 103 'B
vbwProfiler.vbwExecuteLine 104
        If Revision < App.Revision Then
vbwProfiler.vbwExecuteLine 105
            NewVersion = True
        End If
vbwProfiler.vbwExecuteLine 106 'B
    End If
vbwProfiler.vbwExecuteLine 107 'B
vbwProfiler.vbwExecuteLine 108
    url = url & Revision & ".exe"

'If we are working on a higher version in VBE, don't try for newversion
vbwProfiler.vbwExecuteLine 109
    If App.Major * 2 ^ 8 + App.Minor * 2 ^ 4 + App.Revision > _
    Major * 2 ^ 8 + Minor * 2 ^ 4 + Revision Then
vbwProfiler.vbwExecuteLine 110
        NewVersion = False
    End If
vbwProfiler.vbwExecuteLine 111 'B
vbwProfiler.vbwExecuteLine 112
    If NewVersion = True Then
vbwProfiler.vbwExecuteLine 113
        Call frmDpyBox.DpyBox("A new update is available", 10, "New Version")
'Check we have internet access
vbwProfiler.vbwExecuteLine 114
        If HTTPFileExists(url) Then
vbwProfiler.vbwExecuteLine 115
            Call HttpSpawn(url)
        End If
vbwProfiler.vbwExecuteLine 116 'B
    End If
vbwProfiler.vbwExecuteLine 117 'B
'Position cursor at RHS of time displayed
vbwProfiler.vbwExecuteLine 118
    txtFirstStartTime.SelStart = Len(txtFirstStartTime)
vbwProfiler.vbwExecuteLine 119
    txtPostpone.SelStart = Len(txtPostpone)
vbwProfiler.vbwExecuteLine 120
    With mshFinish
vbwProfiler.vbwExecuteLine 121
        .Width = 1795
vbwProfiler.vbwExecuteLine 122
        .FormatString = "<No|<Time"
vbwProfiler.vbwExecuteLine 123
        .ColWidth(0) = 500  'Position
vbwProfiler.vbwExecuteLine 124
        .ColWidth(1) = 1295  'Time
'        For i = 1 To 20
'            .Rows = i + 1
'            .TextMatrix(i, 0) = i
'        Next i
'        .TextMatrix(1, 1) = "13:22:45"
vbwProfiler.vbwExecuteLine 125
    End With

'Flags(0) exists - but not used
vbwProfiler.vbwExecuteLine 126
    RowCount = FlagRow(Flags.Count - 1)
vbwProfiler.vbwExecuteLine 127
    ColCount = FlagCol(Flags.Count - 1)
vbwProfiler.vbwExecuteLine 128
    ColCountFree = ColCount 'Reduces by number of Fixed cols
vbwProfiler.vbwExecuteLine 129
    ReDim Cols(1 To ColCount)
vbwProfiler.vbwExecuteLine 130
Visible = True
'Make the base index invisible as it is not used
vbwProfiler.vbwExecuteLine 131
    Commands(0).Enabled = False
vbwProfiler.vbwExecuteLine 132
    Commands(0).Visible = False
'Set up initial start time, LoadEvents not called
vbwProfiler.vbwExecuteLine 133
    FirstStartTime = Date & " " _
    & Format$(NulToZero(txtFirstStartTime), "00:00") & ":00"
vbwProfiler.vbwExecuteLine 134
Debug.Print Format$(FirstStartTime, "dd-mmm-yyyy")
vbwProfiler.vbwExecuteLine 135
Debug.Print Format$(FirstStartTime, "hh:mm:ss")

vbwProfiler.vbwExecuteLine 136
    Call LoadSequence

vbwProfiler.vbwProcOut 7
vbwProfiler.vbwExecuteLine 137
End Sub


Private Sub Form_Unload(Cancel As Integer)
vbwProfiler.vbwProcIn 8
    Dim i As Integer

    'close all sub forms
vbwProfiler.vbwExecuteLine 138
    For i = Forms.Count - 1 To 1 Step -1
vbwProfiler.vbwExecuteLine 139
        Unload Forms(i)
vbwProfiler.vbwExecuteLine 140
    Next
vbwProfiler.vbwExecuteLine 141
    If Me.WindowState <> vbMinimized Then
vbwProfiler.vbwExecuteLine 142
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
vbwProfiler.vbwExecuteLine 143
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
vbwProfiler.vbwExecuteLine 144
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
vbwProfiler.vbwExecuteLine 145
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
vbwProfiler.vbwExecuteLine 146 'B
vbwProfiler.vbwProcOut 8
vbwProfiler.vbwExecuteLine 147
End Sub

Private Sub RecallTimer_Timer()
vbwProfiler.vbwProcIn 9
vbwProfiler.vbwExecuteLine 148
    RecallTimer.Enabled = False
'Set laststart to 0 so any subsequent Recalls will be queued, if a class flag is UP
vbwProfiler.vbwExecuteLine 149
    LastStart = 0
vbwProfiler.vbwExecuteLine 150
Debug.Print "RecallTimer disabled"
vbwProfiler.vbwProcOut 9
vbwProfiler.vbwExecuteLine 151
End Sub

Private Sub HoistTimer_Timer()
vbwProfiler.vbwProcIn 10
vbwProfiler.vbwExecuteLine 152
    HoistTimer.Enabled = False
'Set last hoist to Blank so any subsequent hoist will action the sound signal
vbwProfiler.vbwExecuteLine 153
    LastHoist = ""
vbwProfiler.vbwExecuteLine 154
Debug.Print "HoistTimer disabled"
vbwProfiler.vbwProcOut 10
vbwProfiler.vbwExecuteLine 155
End Sub

'The timer must run faster than 1 Count interval (normally 1 sec)
'Otherwise a NewCycle could be skipped if the timer is running
'slower than the actual clock time. This could happen
'if the PC is heavily loaded
Private Sub RaceTimer_Timer()
vbwProfiler.vbwProcIn 11
Dim CurrTime As Date    'May be speeded up from Now() time for testing
Dim SecsSinceOutput As Long
Dim TimeToOutput As Date
'Dim MyFirstStartTime As Date

'Adjust curr time to speed up
vbwProfiler.vbwExecuteLine 156
    CurrTime = Now()

vbwProfiler.vbwExecuteLine 157
    If LastTimeOutput = "00:00:00" Then
vbwProfiler.vbwExecuteLine 158
         Call ResetOutput(CurrTime)
    End If
vbwProfiler.vbwExecuteLine 159 'B

vbwProfiler.vbwExecuteLine 160
    SecsSinceOutput = DateDiff("s", LastTimeOutput, CurrTime)
'No Output due yet
vbwProfiler.vbwExecuteLine 161
    If SecsSinceOutput = 0 Then
vbwProfiler.vbwExecuteLine 162
Debug.Print "Skip"
vbwProfiler.vbwProcOut 11
vbwProfiler.vbwExecuteLine 163
        Exit Sub
    End If
vbwProfiler.vbwExecuteLine 164 'B

vbwProfiler.vbwExecuteLine 165
    Do
'        StatusBar1.Panels(1).Text = Time
'        StatusBar1.Panels(2).Text = CurSecs - CycleStartSecs
vbwProfiler.vbwExecuteLine 166
        TimeToOutput = DateAdd("s", 1, LastTimeOutput)
vbwProfiler.vbwExecuteLine 167
        If TimerOutput(TimeToOutput) = True Then
vbwProfiler.vbwExecuteLine 168
             LastTimeOutput = TimeToOutput
        End If
vbwProfiler.vbwExecuteLine 169 'B
vbwProfiler.vbwExecuteLine 170
        SecsSinceOutput = DateDiff("s", LastTimeOutput, CurrTime)
vbwProfiler.vbwExecuteLine 171
        ElapsedTime = DateDiff("s", FirstStartTime, CurrTime)
'Debug.Print ElapsedTime - SecsSinceOutput
vbwProfiler.vbwExecuteLine 172
        Call DoTimerEvents(ElapsedTime - SecsSinceOutput)
vbwProfiler.vbwExecuteLine 173
        lblElapsedTime = aSecToElapsed(ElapsedTime)
'        lblElapsedTime = aSecToElapsed(DateDiff("s", FirstStartTime, CurrTime))
vbwProfiler.vbwExecuteLine 174
        lblCurrTime = Format$(CurrTime, "hh:mm:ss")
vbwProfiler.vbwExecuteLine 175
If SecsSinceOutput > 0 Then
vbwProfiler.vbwExecuteLine 176
     Debug.Print "Catch-up " & SecsSinceOutput
End If
vbwProfiler.vbwExecuteLine 177 'B
vbwProfiler.vbwExecuteLine 178
    Loop Until SecsSinceOutput = 0  'Always execute once
vbwProfiler.vbwProcOut 11
vbwProfiler.vbwExecuteLine 179
End Sub


Private Sub ResetOutput(StartTime As Date)
vbwProfiler.vbwProcIn 12
vbwProfiler.vbwExecuteLine 180
        LastTimeOutput = DateAdd("s", -1, StartTime)
vbwProfiler.vbwProcOut 12
vbwProfiler.vbwExecuteLine 181
End Sub

Private Function aSecToElapsed(Secs As Long) As String
vbwProfiler.vbwProcIn 13
Dim hms As defhms
Dim Sign As Long
Dim aSign As String

'Secs = 3600& * 100&
vbwProfiler.vbwExecuteLine 182
    Sign = Sgn(Secs)    '-1 = -ve, 0 = 0 , +1 = +ve
vbwProfiler.vbwExecuteLine 183
    If Sign = -1 Then
vbwProfiler.vbwExecuteLine 184
        Secs = Secs * Sign 'force +ve
vbwProfiler.vbwExecuteLine 185
        aSign = "-"
    Else
vbwProfiler.vbwExecuteLine 186 'B
vbwProfiler.vbwExecuteLine 187
        aSign = " "
    End If
vbwProfiler.vbwExecuteLine 188 'B
vbwProfiler.vbwExecuteLine 189
    hms.Hour = Int(Secs / 3600&)
vbwProfiler.vbwExecuteLine 190
    Secs = Secs - hms.Hour * 3600&
vbwProfiler.vbwExecuteLine 191
    hms.Min = Int(Secs / 60&)
vbwProfiler.vbwExecuteLine 192
    Secs = Secs - hms.Min * 60&
vbwProfiler.vbwExecuteLine 193
    hms.Sec = Secs
vbwProfiler.vbwExecuteLine 194
    aSecToElapsed = aSign & Format$(hms.Hour, "###")
vbwProfiler.vbwExecuteLine 195
    If Abs(hms.Hour) >= 1 Then
vbwProfiler.vbwExecuteLine 196
         aSecToElapsed = aSecToElapsed & ":"
    End If
vbwProfiler.vbwExecuteLine 197 'B
vbwProfiler.vbwExecuteLine 198
    aSecToElapsed = aSecToElapsed & Format$(hms.Min, "00") _
    & ":" & Format$(hms.Sec, "00")
vbwProfiler.vbwProcOut 13
vbwProfiler.vbwExecuteLine 199
End Function

Private Sub ReloadTimer_Timer()
vbwProfiler.vbwProcIn 14
vbwProfiler.vbwExecuteLine 200
    ReloadTimer.Enabled = False
vbwProfiler.vbwExecuteLine 201
    Call LoadProfile
vbwProfiler.vbwProcOut 14
vbwProfiler.vbwExecuteLine 202
End Sub

Private Sub SignalTimer_Timer(Index As Integer)
vbwProfiler.vbwProcIn 15
Dim FlagIdx  As Long
Dim kb As String
Dim CyclesCompleted As Long
Dim LinkedFlagPos As Long

vbwProfiler.vbwExecuteLine 203
    With SignalAttributes(Index)
vbwProfiler.vbwExecuteLine 204
kb = SignalTimer(Index).Enabled
'Debug.Print Flags(FlagIdx).Visible
'A cycle is completed every time a flag is turned off AFTER it has been on

vbwProfiler.vbwExecuteLine 205
        If .Flag.Pos Then
vbwProfiler.vbwExecuteLine 206
            If Flags(.Flag.Pos).Visible = True Then
vbwProfiler.vbwExecuteLine 207
                .OnCycles = .OnCycles + 1
vbwProfiler.vbwExecuteLine 208
                SignalTimer(Index).Interval = .TTD
            Else
vbwProfiler.vbwExecuteLine 209 'B
vbwProfiler.vbwExecuteLine 210
                SignalTimer(Index).Interval = .TTL
vbwProfiler.vbwExecuteLine 211
                CyclesCompleted = .OnCycles

            End If
vbwProfiler.vbwExecuteLine 212 'B
        Else
vbwProfiler.vbwExecuteLine 213 'B
vbwProfiler.vbwExecuteLine 214
            .OnCycles = .OnCycles + 1
'Terminate Timer & Lower flag
vbwProfiler.vbwExecuteLine 215
            CyclesCompleted = .OnCycles
'MsgBox "Signal(" & Index & ")." & .Name & " has no associated Flag", vbCritical, "SignalTimer_Timer"
        End If
vbwProfiler.vbwExecuteLine 216 'B
'Debug.Print CyclesCompleted & "(" & Index & ")"

'Continuous
vbwProfiler.vbwExecuteLine 217
        If .CyclesRequired = 0 Then
vbwProfiler.vbwExecuteLine 218
            If Loading = False Then
vbwProfiler.vbwExecuteLine 219
                 CyclesCompleted = -1
            End If
vbwProfiler.vbwExecuteLine 220 'B
        End If
vbwProfiler.vbwExecuteLine 221 'B

vbwProfiler.vbwExecuteLine 222
        If Loading And CyclesCompleted > 5 Then
vbwProfiler.vbwExecuteLine 223
             CyclesCompleted = .CyclesRequired
        End If
vbwProfiler.vbwExecuteLine 224 'B
vbwProfiler.vbwExecuteLine 225
        Select Case CyclesCompleted
'The timer has started but we do not want the Signal Off
'In fact we should not have started it in the first place
'vbwLine 226:        Case Is >= .CyclesRequired
        Case Is >= IIf(vbwProfiler.vbwExecuteLine(226), VBWPROFILER_EMPTY, _
        .CyclesRequired)
'This only occurs when the flag is about to be made invisible
'Turn off Signal, before disabling the timer
'Otherwise MakeSignals will start it again
'Click the command button (set to True) to put the flag down
'Only disable if not Continuous
vbwProfiler.vbwExecuteLine 227
            SignalTimer(Index).Enabled = False
'Must be after timer is disabled
vbwProfiler.vbwExecuteLine 228
            .OnCycles = 0
vbwProfiler.vbwExecuteLine 229
            Call LowerRequest(Index)
'        Commands(Index).Value = True
'Click the command button
'kb = SignalTimer(Index).Enabled    'Must be turned off
'Do this last, so if the timer is called again
'another off will be generated, and the timer will
'not re-start
'Remove this from the queue and re-enable with next signal (if any)
'        Call DequeTimer(Index)
'vbwLine 230:        Case Is > .CyclesRequired
        Case Is > IIf(vbwProfiler.vbwExecuteLine(230), VBWPROFILER_EMPTY, _
        .CyclesRequired)
'Continuous
'vbwLine 231:        Case Is < .CyclesRequired
        Case Is < IIf(vbwProfiler.vbwExecuteLine(231), VBWPROFILER_EMPTY, _
        .CyclesRequired)
'Reverse the Visibility of this flag and do another cycle
'No linked Flags are activated
vbwProfiler.vbwExecuteLine 232
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
vbwProfiler.vbwExecuteLine 233 'B
vbwProfiler.vbwExecuteLine 234
    End With
vbwProfiler.vbwProcOut 15
vbwProfiler.vbwExecuteLine 235
End Sub

#If False Then
Private Sub txtFirstStartTime_Change()
    If ValidateStartTime = True Then
        Call ResetCmd
    End If
End Sub
#End If

Private Sub txtPostpone_Change()
vbwProfiler.vbwProcIn 16
vbwProfiler.vbwExecuteLine 236
    Call ValidatePostponeTime
vbwProfiler.vbwProcOut 16
vbwProfiler.vbwExecuteLine 237
End Sub

Private Function ValidateStartTime() As Boolean
vbwProfiler.vbwProcIn 17

vbwProfiler.vbwExecuteLine 238
    On Error GoTo ValidateStartTime_error
vbwProfiler.vbwExecuteLine 239
    If txtFirstStartTime = "" Then
vbwProfiler.vbwExecuteLine 240
        txtFirstStartTime.BackColor = vbRed
    Else
vbwProfiler.vbwExecuteLine 241 'B
vbwProfiler.vbwExecuteLine 242
        txtFirstStartTime.BackColor = vbWhite
    End If
vbwProfiler.vbwExecuteLine 243 'B
vbwProfiler.vbwExecuteLine 244
    If Len(txtFirstStartTime) = 4 _
    And CLng(NulToZero(txtFirstStartTime)) >= 0 _
    And CLng(NulToZero(txtFirstStartTime)) <= 2400 _
    And IsNumeric(NulToZero(txtFirstStartTime)) = True Then
vbwProfiler.vbwExecuteLine 245
        FirstStartTime = Date & " " _
        & Format$(NulToZero(txtFirstStartTime), "00:00") & ":00"
vbwProfiler.vbwExecuteLine 246
        On Error GoTo 0
vbwProfiler.vbwExecuteLine 247
        txtFirstStartTime.ForeColor = vbBlack
'Must not only reset the flags because once the start sequence
'has commenced the whole profile should be reloaded
'        Call ResetFlags
vbwProfiler.vbwExecuteLine 248
        ValidateStartTime = True
vbwProfiler.vbwProcOut 17
vbwProfiler.vbwExecuteLine 249
        Exit Function
    End If
vbwProfiler.vbwExecuteLine 250 'B
ValidateStartTime_error:
vbwProfiler.vbwExecuteLine 251
    txtFirstStartTime.ForeColor = vbRed
vbwProfiler.vbwProcOut 17
vbwProfiler.vbwExecuteLine 252
End Function

Private Function ValidatePostponeTime() As Boolean
vbwProfiler.vbwProcIn 18

vbwProfiler.vbwExecuteLine 253
    On Error GoTo ValidatePostponeTime_error
vbwProfiler.vbwExecuteLine 254
    If txtPostpone = "" Then
vbwProfiler.vbwExecuteLine 255
        txtPostpone.BackColor = vbRed
    Else
vbwProfiler.vbwExecuteLine 256 'B
vbwProfiler.vbwExecuteLine 257
        txtPostpone.BackColor = vbWhite
    End If
vbwProfiler.vbwExecuteLine 258 'B
vbwProfiler.vbwExecuteLine 259
    If IsNumeric(NulToZero(txtPostpone)) = True Then
vbwProfiler.vbwExecuteLine 260
        txtPostpone.ForeColor = vbBlack
vbwProfiler.vbwExecuteLine 261
        ValidatePostponeTime = True
vbwProfiler.vbwProcOut 18
vbwProfiler.vbwExecuteLine 262
        Exit Function
    End If
vbwProfiler.vbwExecuteLine 263 'B
ValidatePostponeTime_error:
vbwProfiler.vbwExecuteLine 264
    txtPostpone.ForeColor = vbRed
vbwProfiler.vbwProcOut 18
vbwProfiler.vbwExecuteLine 265
End Function

Public Function ResetFlags()
vbwProfiler.vbwProcIn 19
Dim MyImage As Image

vbwProfiler.vbwExecuteLine 266
If DebugFlags Then
vbwProfiler.vbwExecuteLine 267
     Debug.Print "ResetFlags"
End If
vbwProfiler.vbwExecuteLine 268 'B
vbwProfiler.vbwExecuteLine 269
    For Each MyImage In frmMain.Flags
vbwProfiler.vbwExecuteLine 270
        MyImage.Picture = Nothing
vbwProfiler.vbwExecuteLine 271
    Next
vbwProfiler.vbwProcOut 19
vbwProfiler.vbwExecuteLine 272
End Function

Public Function ResetCommands()
vbwProfiler.vbwProcIn 20
Dim MyCommand As CommandButton
vbwProfiler.vbwExecuteLine 273
    For Each MyCommand In Commands
vbwProfiler.vbwExecuteLine 274
        If MyCommand.Index <> 0 Then
vbwProfiler.vbwExecuteLine 275
            MyCommand.Enabled = True
vbwProfiler.vbwExecuteLine 276
            MyCommand.Visible = True
        End If
vbwProfiler.vbwExecuteLine 277 'B
vbwProfiler.vbwExecuteLine 278
    Next MyCommand
vbwProfiler.vbwProcOut 20
vbwProfiler.vbwExecuteLine 279
End Function
Public Function ResetSignalTimers()
vbwProfiler.vbwProcIn 21
Dim MySignalTimer As Timer

vbwProfiler.vbwExecuteLine 280
    For Each MySignalTimer In frmMain.SignalTimer
vbwProfiler.vbwExecuteLine 281
        If MySignalTimer.Index > 0 Then  'Dont delete SignalTimer(0)
vbwProfiler.vbwExecuteLine 282
            Unload MySignalTimer
        End If
vbwProfiler.vbwExecuteLine 283 'B
vbwProfiler.vbwExecuteLine 284
    Next
vbwProfiler.vbwExecuteLine 285
    HoistTimer.Enabled = False
vbwProfiler.vbwExecuteLine 286
    LastHoist = ""
vbwProfiler.vbwExecuteLine 287
    RecallTimer.Enabled = False
vbwProfiler.vbwExecuteLine 288
    LastStart = 0
vbwProfiler.vbwExecuteLine 289
Debug.Print "HoistTimer disabled"
vbwProfiler.vbwProcOut 21
vbwProfiler.vbwExecuteLine 290
End Function

Public Function ResetFinish()
vbwProfiler.vbwProcIn 22
Dim Row As Long
Dim Col As Long

vbwProfiler.vbwExecuteLine 291
    With mshFinish
'Clear rows (except 1)
vbwProfiler.vbwExecuteLine 292
        For Row = 2 To .Rows - 1
vbwProfiler.vbwExecuteLine 293
            .RemoveItem 1
vbwProfiler.vbwExecuteLine 294
        Next Row
'Clear Row 1
vbwProfiler.vbwExecuteLine 295
        For Col = 0 To .Cols - 1
vbwProfiler.vbwExecuteLine 296
            .TextMatrix(1, Col) = ""
vbwProfiler.vbwExecuteLine 297
        Next Col
vbwProfiler.vbwExecuteLine 298
    End With
vbwProfiler.vbwProcOut 22
vbwProfiler.vbwExecuteLine 299
End Function

#If False Then
Public Function ResetCmd()
'Must have Property Style set to 1=Graphical
'    cmdPostpone.Enabled = True
'Recall Signal says up
'    cmdRecall.BackColor = cbDefault
    cmdFinish.BackColor = cbDefault
'    cmdHorn.BackColor = cbDefault
'    cmdPostpone.SetFocus
End Function
#End If

'Requires MSINET.OCX
'See http://officeone.mvps.org/vba/http_file_exists.html
Public Function HTTPFileExists(ByVal url As String) As Boolean
vbwProfiler.vbwProcIn 23
    Dim S As String
    Dim Exists As Boolean
vbwProfiler.vbwExecuteLine 300
    On Error GoTo Inet1_Error
vbwProfiler.vbwExecuteLine 301
    With Inet1
vbwProfiler.vbwExecuteLine 302
        .RequestTimeout = 20
vbwProfiler.vbwExecuteLine 303
        .Protocol = icHTTP
vbwProfiler.vbwExecuteLine 304
        .url = url
vbwProfiler.vbwExecuteLine 305
        .Execute
'see http://support.microsoft.com/kb/182152 =True doesnt work
'vbwLine 306:        Do While .StillExecuting <> False
        Do While vbwProfiler.vbwExecuteLine(306) Or .StillExecuting <> False
vbwProfiler.vbwExecuteLine 307
            DoEvents
vbwProfiler.vbwExecuteLine 308
        Loop
vbwProfiler.vbwExecuteLine 309
        S = UCase(.GetHeader())
vbwProfiler.vbwExecuteLine 310
        Exists = (InStr(1, S, "200 OK") > 0)
vbwProfiler.vbwExecuteLine 311
    End With
vbwProfiler.vbwExecuteLine 312
    HTTPFileExists = Exists
vbwProfiler.vbwProcOut 23
vbwProfiler.vbwExecuteLine 313
    Exit Function
Inet1_Error:
vbwProfiler.vbwExecuteLine 314
    Select Case Err.Number
'vbwLine 315:    Case Is = 35764 '
    Case Is = IIf(vbwProfiler.vbwExecuteLine(315), VBWPROFILER_EMPTY, _
        35764 )'
    End Select
vbwProfiler.vbwExecuteLine 316 'B

vbwProfiler.vbwProcOut 23
vbwProfiler.vbwExecuteLine 317
End Function

Public Function HttpSpawn(url As String)
vbwProfiler.vbwProcIn 24
Dim r As Long
Dim Command As String

vbwProfiler.vbwExecuteLine 318
If Environ("windir") <> "" Then
vbwProfiler.vbwExecuteLine 319
    r = ShellExecute(0, "open", url, 0, 0, 1)
Else
vbwProfiler.vbwExecuteLine 320 'B
'try for linux compatibility
vbwProfiler.vbwExecuteLine 321
    Command = "winebrowser " & url & " ""%1"""

vbwProfiler.vbwExecuteLine 322
    Shell (Command)
End If
vbwProfiler.vbwExecuteLine 323 'B
vbwProfiler.vbwProcOut 24
vbwProfiler.vbwExecuteLine 324
End Function

Public Function PositionCommand(Idx As Long)
'You dont need these unless testing this module in VBE
'If you have a break set frmMain is minimised and
'the Scale values will be 0
'Dont leave a blank gap
vbwProfiler.vbwProcIn 25
vbwProfiler.vbwExecuteLine 325
    With Commands(Idx)
vbwProfiler.vbwExecuteLine 326
        .Caption = .Caption & "(" & Idx & ")"
vbwProfiler.vbwExecuteLine 327
        If .Visible = True Then
'This will be overwritten with the Name from SignalAttributes
'Align first command with top of main frame
vbwProfiler.vbwExecuteLine 328
            WindowState = vbNormal  'Scale will be 0 in VBE (window is minimized)
vbwProfiler.vbwExecuteLine 329
            .Top = ScaleTop + fraMain.Top + NextCommandTop
vbwProfiler.vbwExecuteLine 330
            If .Top + .Height > fraMain.Top + fraMain.Height Then
vbwProfiler.vbwExecuteLine 331
                NextCommandTop = 0
vbwProfiler.vbwExecuteLine 332
                Width = Width + .Width
vbwProfiler.vbwExecuteLine 333
                WindowState = vbNormal  'Scale will be 0 in VBE (window is minimized)
vbwProfiler.vbwExecuteLine 334
                .Top = ScaleTop + fraMain.Top + NextCommandTop
            End If
vbwProfiler.vbwExecuteLine 335 'B
vbwProfiler.vbwExecuteLine 336
            WindowState = vbNormal  'Scale will be 0 in VBE (window is minimized)
vbwProfiler.vbwExecuteLine 337
            .Left = ScaleWidth - .Width
vbwProfiler.vbwExecuteLine 338
            NextCommandTop = NextCommandTop + .Height
        End If
vbwProfiler.vbwExecuteLine 339 'B
vbwProfiler.vbwExecuteLine 340
    End With
vbwProfiler.vbwProcOut 25
vbwProfiler.vbwExecuteLine 341
End Function

Private Function CommandColor()
vbwProfiler.vbwProcIn 26
Dim MyCommand As CommandButton
Dim i As Integer


vbwProfiler.vbwExecuteLine 342
    For Each MyCommand In Commands
vbwProfiler.vbwExecuteLine 343
        If MyCommand.Index > 0 Then 'skip command(0)
'Command may have been created before SignalAttributes
vbwProfiler.vbwExecuteLine 344
            If MyCommand.Index <= UBound(SignalAttributes) Then
vbwProfiler.vbwExecuteLine 345
                If SignalAttributes(MyCommand.Index).Flag.Pos = 0 Then
vbwProfiler.vbwExecuteLine 346
                    MyCommand.BackColor = cbDefault
                Else
vbwProfiler.vbwExecuteLine 347 'B
vbwProfiler.vbwExecuteLine 348
                    MyCommand.BackColor = vbGreen
                End If
vbwProfiler.vbwExecuteLine 349 'B
vbwProfiler.vbwExecuteLine 350
                For i = 0 To UBound(CmdQ)
vbwProfiler.vbwExecuteLine 351
                    If CmdQ(i) = MyCommand.Index Then
vbwProfiler.vbwExecuteLine 352
                        MyCommand.BackColor = vbCyan
                    End If
vbwProfiler.vbwExecuteLine 353 'B
vbwProfiler.vbwExecuteLine 354
                Next i
            End If
vbwProfiler.vbwExecuteLine 355 'B
        End If
vbwProfiler.vbwExecuteLine 356 'B
vbwProfiler.vbwExecuteLine 357
    Next MyCommand
vbwProfiler.vbwProcOut 26
vbwProfiler.vbwExecuteLine 358
End Function

'Called by the a Command to Raise a flag
'Must called by the Link (Sound may be clicked, with Sound still running)
'Queues if fixed position and Fixed position in use)
'Queues the the command if HoistTimer is running for this Group
'Queues Recall if ClassFlag is UP
'Actions Linked Flag by calling LinkRequest (If not Queued)
'Starts HoistTimer for this Group, if not Queueable (Flags.Queue=False)

Public Function RaiseRequest(ByVal Idx As Long)
vbwProfiler.vbwProcIn 27
Dim SoundEnabled As Boolean
Dim Pos As Long
Dim QueueSignal As Long
Dim NextCmd As Long
Dim MyLink As defLink
Dim ClassIdx As Long
Dim RecallIdx As Long
'Dim PreparatoryIdx As Long

vbwProfiler.vbwExecuteLine 359
    If Idx > UBound(SignalAttributes) Then
vbwProfiler.vbwProcOut 27
vbwProfiler.vbwExecuteLine 360
        Exit Function
    End If
vbwProfiler.vbwExecuteLine 361 'B

'Check if Request requires Queueing or actioning
'If Fixed position and Position is in use
vbwProfiler.vbwExecuteLine 362
    With SignalAttributes(Idx)

vbwProfiler.vbwExecuteLine 363
        Select Case .Name
'vbwLine 364:        Case Is = "Finish"
        Case Is = IIf(vbwProfiler.vbwExecuteLine(364), VBWPROFILER_EMPTY, _
        "Finish")
'A Finish Requires Completely different handling, Only Raise the Link UP event
'Do not RaiseRequest to put Finish Flag up (would toggle finish command actions)
vbwProfiler.vbwExecuteLine 365
            If .Name = "Finish" Then
'A Finish always clocks the time and must give correct no Linked signals
vbwProfiler.vbwExecuteLine 366
               If Loading = False Then
vbwProfiler.vbwExecuteLine 367
                    Call FinishTime
'The linked signal requires Queueing, if currently flashing
vbwProfiler.vbwExecuteLine 368
                    Call LinkRequest(Idx)
                End If
vbwProfiler.vbwExecuteLine 369 'B
vbwProfiler.vbwProcOut 27
vbwProfiler.vbwExecuteLine 370
            Exit Function
            End If
vbwProfiler.vbwExecuteLine 371 'B
'vbwLine 372:        Case Is = "Recall", "General Recall"
        Case Is = IIf(vbwProfiler.vbwExecuteLine(372), VBWPROFILER_EMPTY, _
        "Recall"), "General Recall"
vbwProfiler.vbwExecuteLine 373
            Call LowerGroup(Idx)
        End Select
vbwProfiler.vbwExecuteLine 374 'B
'Debug.Print "RaiseReq " & .Name
vbwProfiler.vbwExecuteLine 375
        Pos = RC(.Flag.FixedRow, .Flag.FixedCol)
'If the Flag has a fixed position, check if any flag is already in this position
vbwProfiler.vbwExecuteLine 376
        If Pos > 0 And .Flag.Queue = True Then
vbwProfiler.vbwExecuteLine 377
            If Flags(Pos).Picture.Handle <> 0 Then
vbwProfiler.vbwExecuteLine 378
                QueueSignal = Idx
'Debug.Print "Q Flag is UP"
            End If
vbwProfiler.vbwExecuteLine 379 'B
        End If
vbwProfiler.vbwExecuteLine 380 'B

'Queues the the command if HoistTimer is running for this Group
'So linked sound signal not made as another flag will be raised on same col
vbwProfiler.vbwExecuteLine 381
        If .Group = LastHoist And .Flag.Queue Then
vbwProfiler.vbwExecuteLine 382
            QueueSignal = Idx
'Debug.Print "Q Timer On"
        End If
vbwProfiler.vbwExecuteLine 383 'B

'Queues Recall if Any ClassFlag is UP (Must Wait until Class Flag is dropped)
'If Recall is pressed within 10 seconds of dropping the Class flag
'It must not be queued as it is a recall for the Class that has just started.
vbwProfiler.vbwExecuteLine 384
        If .Group = "Recall" Then
vbwProfiler.vbwExecuteLine 385
            ClassIdx = GroupIdx("Class")
'Class Flag (may be another Class and NOT the one just started) is up and not
'within 10 secs of last start
vbwProfiler.vbwExecuteLine 386
            If ClassIdx > 0 And LastStart = 0 Then
vbwProfiler.vbwExecuteLine 387
                QueueSignal = Idx
'Debug.Print "Q Class Recall"
            End If
vbwProfiler.vbwExecuteLine 388 'B
        End If
vbwProfiler.vbwExecuteLine 389 'B

vbwProfiler.vbwExecuteLine 390
        If Loading = True Then
vbwProfiler.vbwExecuteLine 391
            QueueSignal = 0
        End If
vbwProfiler.vbwExecuteLine 392 'B

vbwProfiler.vbwExecuteLine 393
        If QueueSignal > 0 Then
'If ClassIdx > 0 Then
'                MyLink.Temp = True
'                MyLink.Flag = QueueSignal
'                MyLink.Raise = True
'                MyLink.Type = "DownLink"
'                Call CreateLink(ClassIdx, MyLink)
'                Commands(QueueSignal).BackColor = vbCyan
'            Else
vbwProfiler.vbwExecuteLine 394
                Call QueueRequest(QueueSignal)
'            End If
        Else
vbwProfiler.vbwExecuteLine 395 'B

'Put the Flag up
vbwProfiler.vbwExecuteLine 396
            Call RaiseFlag(Idx)

'Actions Linked Flag by calling LinkRequest (If not Queued)
vbwProfiler.vbwExecuteLine 397
            Call LinkRequest(Idx)

'Start HoistTimer for this Group, if not Queueable (Flags.Queue=False)
'So we dont Create a Second Sound signal
vbwProfiler.vbwExecuteLine 398
            If .Flag.Queue = False Then
vbwProfiler.vbwExecuteLine 399
                HoistTimer.Enabled = False
vbwProfiler.vbwExecuteLine 400
                HoistTimer.Enabled = True
vbwProfiler.vbwExecuteLine 401
                LastHoist = .Group
            End If
vbwProfiler.vbwExecuteLine 402 'B
        End If  'Not Queued
vbwProfiler.vbwExecuteLine 403 'B

vbwProfiler.vbwExecuteLine 404
    End With

vbwProfiler.vbwProcOut 27
vbwProfiler.vbwExecuteLine 405
End Function

'Called by RaiseRequest and to action the UP link
Public Function RaiseFlag(ByVal Idx As Long)
vbwProfiler.vbwProcIn 28
Dim Col As Long
Dim Row As Long

'Load Profile-Linked Signals with a higher idx will not have been created
'Debug.Print "Raise " & SignalAttributes(Idx).Name
'Action Command now
'Display Image first (if there is one for this Signal)
vbwProfiler.vbwExecuteLine 406
    With SignalAttributes(Idx)

vbwProfiler.vbwExecuteLine 407
        If Not .Image Is Nothing Then
vbwProfiler.vbwExecuteLine 408
            Call NextFreeGroupFlagPos(Idx)
vbwProfiler.vbwExecuteLine 409
            If .Flag.Col > 0 Then
vbwProfiler.vbwExecuteLine 410
                .Flag.Pos = RC(.Flag.Row, .Flag.Col)
            Else
vbwProfiler.vbwExecuteLine 411 'B
vbwProfiler.vbwExecuteLine 412
MsgBox "No free Flag positions", vbCritical, "RaiseFlag"
            End If
vbwProfiler.vbwExecuteLine 413 'B
        End If
vbwProfiler.vbwExecuteLine 414 'B

'If we have a flag position then create it (not set if no Image)
vbwProfiler.vbwExecuteLine 415
        If .Flag.Pos > 0 Then
vbwProfiler.vbwExecuteLine 416
            Flags(.Flag.Pos).Picture = .Image
'You have to set it to False becuase FlagVisibility only reacts to a change
vbwProfiler.vbwExecuteLine 417
            Flags(.Flag.Pos).Visible = False
'Must use flagvisibility to create controller event
vbwProfiler.vbwExecuteLine 418
            Call FlagVisibility(Idx, True)
vbwProfiler.vbwExecuteLine 419
            .Flag.Changed = True
        End If
vbwProfiler.vbwExecuteLine 420 'B

vbwProfiler.vbwExecuteLine 421
        Commands(Idx).BackColor = vbGreen
'May still be a timer even if no image to display
vbwProfiler.vbwExecuteLine 422
        If .TTL > 0 Then
vbwProfiler.vbwExecuteLine 423
            SignalTimer(Idx).Interval = .TTL
vbwProfiler.vbwExecuteLine 424
            SignalTimer(Idx).Enabled = True
        End If
vbwProfiler.vbwExecuteLine 425 'B
vbwProfiler.vbwExecuteLine 426
    End With
vbwProfiler.vbwExecuteLine 427
    Call ResetCols  'Resets Cols().Group & .Items from SignalAttributes

vbwProfiler.vbwProcOut 28
vbwProfiler.vbwExecuteLine 428
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
vbwProfiler.vbwProcIn 29
Dim NextCmd As Integer
Dim i As Long

'Load Profile-Linked Signals with a higher idx will not have been created
vbwProfiler.vbwExecuteLine 429
    If Idx > UBound(SignalAttributes) Then
vbwProfiler.vbwProcOut 29
vbwProfiler.vbwExecuteLine 430
        Exit Function
    End If
vbwProfiler.vbwExecuteLine 431 'B

vbwProfiler.vbwExecuteLine 432
    With SignalAttributes(Idx)
'Debug.Print "LowerReq " & SignalAttributes(Idx).Name

vbwProfiler.vbwExecuteLine 433
For i = 0 To UBound(CmdQ)
vbwProfiler.vbwExecuteLine 434
    If CmdQ(CInt(i)) <> 0 Then
vbwProfiler.vbwExecuteLine 435
Debug.Print "Queued(" & i & ")=" & CmdQ(CInt(i))
    End If
vbwProfiler.vbwExecuteLine 436 'B
vbwProfiler.vbwExecuteLine 437
Next i

 'Lower Flag then any below the Flag
vbwProfiler.vbwExecuteLine 438
        Call LowerFlag(Idx)
vbwProfiler.vbwExecuteLine 439
        Call LinkRequest(Idx)

'Overlapped Position Recall above ClassFlag Requesting Recall
'Dequeues Recall when Any Class Lowered by calling RaiseRequest
vbwProfiler.vbwExecuteLine 440
        If .Group = "Class" Then
vbwProfiler.vbwExecuteLine 441
            NextCmd = DequeCmd("Recall")
vbwProfiler.vbwExecuteLine 442
            If NextCmd <> 0 Then
vbwProfiler.vbwExecuteLine 443
                Call RaiseRequest(NextCmd)
            End If
vbwProfiler.vbwExecuteLine 444 'B
'Call CompressCols
        End If
vbwProfiler.vbwExecuteLine 445 'B

'Dequeues any Commands in the same Group Calling RaiseRequest
vbwProfiler.vbwExecuteLine 446
        NextCmd = DequeCmd(.Group)
vbwProfiler.vbwExecuteLine 447
        If NextCmd <> 0 Then
vbwProfiler.vbwExecuteLine 448
            Call RaiseRequest(NextCmd)
        End If
vbwProfiler.vbwExecuteLine 449 'B
vbwProfiler.vbwExecuteLine 450
    End With

vbwProfiler.vbwExecuteLine 451
    Call ResetCols
'    Call CommandColor

vbwProfiler.vbwProcOut 29
vbwProfiler.vbwExecuteLine 452
End Function

'Called by LowerRequest
'Lowers the Flag and any subservient flags
'Does not action any links
Private Function LowerFlag(ByVal Idx As Long)
vbwProfiler.vbwProcIn 30
Dim StartCol As Long
Dim StartRow As Long
Dim Group As String
Dim i As Long

'Debug.Print "LowerFlag " & SignalAttributes(Idx).Name

vbwProfiler.vbwExecuteLine 453
    With SignalAttributes(Idx)
vbwProfiler.vbwExecuteLine 454
        StartCol = .Flag.Col
vbwProfiler.vbwExecuteLine 455
        StartRow = .Flag.Row
vbwProfiler.vbwExecuteLine 456
        Group = .Group
vbwProfiler.vbwExecuteLine 457
    End With

'Calls LowerFlag to Lower any subservient flags WITHOUT actioning any link
'(Only action the Link of the TOP flag)
vbwProfiler.vbwExecuteLine 458
    For i = 1 To UBound(SignalAttributes)
vbwProfiler.vbwExecuteLine 459
        With SignalAttributes(i)
vbwProfiler.vbwExecuteLine 460
            If .Group = Group Then

'Stop first. otherwise Timer will fail when it calls FlagVisibility
vbwProfiler.vbwExecuteLine 461
                If SignalTimer(i).Enabled = True Then
vbwProfiler.vbwExecuteLine 462
                    SignalTimer(i).Enabled = False
                End If
vbwProfiler.vbwExecuteLine 463 'B

vbwProfiler.vbwExecuteLine 464
               If .Flag.Pos > 0 Then
'If in different col or lower row in same col remove
'                    If .Flag.Col <> StartCol Or .Flag.Row >= StartRow Then
'Change to only flags in same Col so Class Flags in different Cols are not dropped
vbwProfiler.vbwExecuteLine 465
                    If .Flag.Col = StartCol And .Flag.Row >= StartRow Then
'Clear the flag (if it exists)
vbwProfiler.vbwExecuteLine 466
                        If Flags(.Flag.Pos).Picture.Handle <> 0 Then
'If .Flag.Pos=0, FlagVisibility reports an error so must do first
vbwProfiler.vbwExecuteLine 467
                            Call FlagVisibility(i, False)
vbwProfiler.vbwExecuteLine 468
                            Flags(.Flag.Pos).Picture = Nothing
                        End If
vbwProfiler.vbwExecuteLine 469 'B
vbwProfiler.vbwExecuteLine 470
                        .Flag.Pos = 0
vbwProfiler.vbwExecuteLine 471
                        .Flag.Col = 0
vbwProfiler.vbwExecuteLine 472
                        .Flag.Row = 0
vbwProfiler.vbwExecuteLine 473
                        Commands(i).BackColor = cbDefault
'Stop Hoist Timer if last Flag up in this Group
vbwProfiler.vbwExecuteLine 474
                        If .Group = LastHoist Then
vbwProfiler.vbwExecuteLine 475
                            HoistTimer.Enabled = False
vbwProfiler.vbwExecuteLine 476
                            LastHoist = ""
vbwProfiler.vbwExecuteLine 477
Debug.Print "HoistTimer disabled"
                        End If
vbwProfiler.vbwExecuteLine 478 'B
'Keep the last start Flag for 10 secs
vbwProfiler.vbwExecuteLine 479
                        If .Group = "Class" Then
vbwProfiler.vbwExecuteLine 480
                            RecallTimer.Enabled = True
vbwProfiler.vbwExecuteLine 481
                            LastStart = Idx
vbwProfiler.vbwExecuteLine 482
Debug.Print "RecallTimer enabled"
                        End If
vbwProfiler.vbwExecuteLine 483 'B
                    End If
vbwProfiler.vbwExecuteLine 484 'B
                End If
vbwProfiler.vbwExecuteLine 485 'B
            End If
vbwProfiler.vbwExecuteLine 486 'B
vbwProfiler.vbwExecuteLine 487
        End With
vbwProfiler.vbwExecuteLine 488
    Next i

'    Call CompressCols

vbwProfiler.vbwProcOut 30
vbwProfiler.vbwExecuteLine 489
End Function

'Lower any flags that are up in this Group (without calling Linked flags)
Private Function LowerGroup(ByVal Idx As Long)
vbwProfiler.vbwProcIn 31
Dim i As Long
Dim Group As String

vbwProfiler.vbwExecuteLine 490
    Group = SignalAttributes(Idx).Group
vbwProfiler.vbwExecuteLine 491
    For i = 1 To UBound(SignalAttributes)
vbwProfiler.vbwExecuteLine 492
        With SignalAttributes(i)
vbwProfiler.vbwExecuteLine 493
            If .Group = Group And .Flag.Pos > 0 Then
vbwProfiler.vbwExecuteLine 494
                Call LowerRequest(i)
            End If
vbwProfiler.vbwExecuteLine 495 'B
vbwProfiler.vbwExecuteLine 496
        End With
vbwProfiler.vbwExecuteLine 497
    Next i
vbwProfiler.vbwProcOut 31
vbwProfiler.vbwExecuteLine 498
End Function

'Calling Flag must be positioned (Up or Down) before LinkRequest is Called
'If HoistTimer for this Group is running (LastHoist = IdxGroup) dont action Link
'If Queueable (Flags.Queue=True) there should not be a link

Private Function LinkRequest(ByVal Idx As Long)
vbwProfiler.vbwProcIn 32
Dim Lidx As Long
Dim MyLink As defLink

vbwProfiler.vbwExecuteLine 499
    With SignalAttributes(Idx)
vbwProfiler.vbwExecuteLine 500
        If IsArrayInitialised(.Links) Then
vbwProfiler.vbwExecuteLine 501
            For Lidx = 0 To UBound(.Links)
vbwProfiler.vbwExecuteLine 502
                MyLink = .Links(Lidx)
vbwProfiler.vbwExecuteLine 503
                If MyLink.Flag > 0 Then
'If MyLink.Flag = 4 Then Stop
vbwProfiler.vbwExecuteLine 504
                    If .Flag.Pos > 0 And MyLink.Type = "UpLink" Then
vbwProfiler.vbwExecuteLine 505
                        Call LinkExecute(Idx, MyLink)
                    End If
vbwProfiler.vbwExecuteLine 506 'B
vbwProfiler.vbwExecuteLine 507
                    If .Flag.Pos = 0 And MyLink.Type = "DownLink" Then
vbwProfiler.vbwExecuteLine 508
                        Call LinkExecute(Idx, MyLink)
                    End If
vbwProfiler.vbwExecuteLine 509 'B
                End If
vbwProfiler.vbwExecuteLine 510 'B
'Stop 'Link execute can delete a links index which causes a subscript error
'Change for to a loop with mo0re checking
vbwProfiler.vbwExecuteLine 511
            Next Lidx
        End If
vbwProfiler.vbwExecuteLine 512 'B
vbwProfiler.vbwExecuteLine 513
    End With
vbwProfiler.vbwProcOut 32
vbwProfiler.vbwExecuteLine 514
End Function


'byval
Private Function LinkExecute(Idx As Long, Link As defLink)
vbwProfiler.vbwProcIn 33
vbwProfiler.vbwExecuteLine 515
        With Link  'Raising Signal

'On ProfileLoad the linked flag may not have been created yet
'.Name is cleared when the Hoist Timer has finished its cycle (5 secs)
vbwProfiler.vbwExecuteLine 516
            If .Flag <> 0 And .Flag <= UBound(SignalAttributes) Then
vbwProfiler.vbwExecuteLine 517
Debug.Print "Link " & SignalAttributes(Idx).Name & " > " & SignalAttributes(.Flag).Name
vbwProfiler.vbwExecuteLine 518
                If SignalAttributes(Idx).Group <> LastHoist Then
vbwProfiler.vbwExecuteLine 519
                    If .Raise = True Then   'Raise Linked flag
vbwProfiler.vbwExecuteLine 520
                        Call RaiseRequest(.Flag)
                    Else
vbwProfiler.vbwExecuteLine 521 'B
vbwProfiler.vbwExecuteLine 522
                        Call LowerRequest(.Flag)   'Lower Linked flag
                    End If
vbwProfiler.vbwExecuteLine 523 'B
                Else
vbwProfiler.vbwExecuteLine 524 'B
vbwProfiler.vbwExecuteLine 525
Debug.Print "Suppressed"
                End If
vbwProfiler.vbwExecuteLine 526 'B
            Else
vbwProfiler.vbwExecuteLine 527 'B
'There are no Linked Flags to this Flag
'Debug.Print "Link " & SignalAttributes(Idx).Name & " > none"
            End If
vbwProfiler.vbwExecuteLine 528 'B
'If a temporary link delete it
vbwProfiler.vbwExecuteLine 529
        If Link.Temp = True Then
vbwProfiler.vbwExecuteLine 530
            Call LinkTempRemove(Idx, Link)
        End If
vbwProfiler.vbwExecuteLine 531 'B
vbwProfiler.vbwExecuteLine 532
        End With
vbwProfiler.vbwProcOut 33
vbwProfiler.vbwExecuteLine 533
End Function

Private Function RC(ByVal Row As Long, ByVal Col As Long) As Long
'Both must be valid as a pair
vbwProfiler.vbwProcIn 34
vbwProfiler.vbwExecuteLine 534
    If Row > 0 And Col > 0 Then
vbwProfiler.vbwExecuteLine 535
        RC = (Row - 1) * 10 + Col
    End If
vbwProfiler.vbwExecuteLine 536 'B
vbwProfiler.vbwProcOut 34
vbwProfiler.vbwExecuteLine 537
End Function

Private Function FlagRow(ByVal Pos As Long) As Long
vbwProfiler.vbwProcIn 35
vbwProfiler.vbwExecuteLine 538
    If Pos > 0 Then
vbwProfiler.vbwExecuteLine 539
        FlagRow = (Pos - 1) \ 10 + 1
    End If
vbwProfiler.vbwExecuteLine 540 'B
vbwProfiler.vbwProcOut 35
vbwProfiler.vbwExecuteLine 541
End Function
    
Private Function FlagCol(ByVal Pos As Long) As Long
vbwProfiler.vbwProcIn 36
vbwProfiler.vbwExecuteLine 542
    If Pos > 0 Then
vbwProfiler.vbwExecuteLine 543
        FlagCol = Pos - (FlagRow(Pos) - 1) * 10
    End If
vbwProfiler.vbwExecuteLine 544 'B
vbwProfiler.vbwProcOut 36
vbwProfiler.vbwExecuteLine 545
End Function

'Called when Raising Flag, SignalAttributes Col & Row = 0 if no Position available
Private Function NextFreeGroupFlagPos(ByVal Idx As Long)
vbwProfiler.vbwProcIn 37
Dim Col As Long
Dim Row As Long
Dim Pos As Long
Dim ClassIdx As Long

'If we do not have a set position see if this flag has a parent
'ie a 2 flag hoist and the parent flag is up

'    Call ResetCols
'If Idx = 9 Then Stop
vbwProfiler.vbwExecuteLine 546
   With SignalAttributes(Idx).Flag
'Get the Column first
vbwProfiler.vbwExecuteLine 547
        If .FixedCol > 0 Then
vbwProfiler.vbwExecuteLine 548
            .Col = .FixedCol
        End If
vbwProfiler.vbwExecuteLine 549 'B

'See if this flag wants placing in same col as the first Class Flag
'DONT REMOVE may want to use it later
#If False Then
            If .Col = 0 Then
            Select Case SignalAttributes(Idx).Group
            Case Is = "Preparatory", "Recall", "Shortened"
                   ClassIdx = GroupIdx("Class")
                    If ClassIdx > 0 Then
'Put flag in same col
Debug.Print "Top Row"
                        .Col = SignalAttributes(ClassIdx).Flag.Col
                        .Row = SignalAttributes(ClassIdx).Flag.Row
                        Call ShiftDown(.Row, .Col)
                    End If
            End Select
        End If
#End If
vbwProfiler.vbwExecuteLine 550
        If .Col = 0 Then
'See if we have a flag Raised in this Group with a spare Row available
vbwProfiler.vbwExecuteLine 551
            If Left$(SignalAttributes(Idx).Name, 6) <> "Class " Then
'Class Flags are always in separate cols (Keep in the same group)
vbwProfiler.vbwExecuteLine 552
                For Col = 1 To ColCountFree
vbwProfiler.vbwExecuteLine 553
                    If Cols(Col).Group = SignalAttributes(Idx).Group Then
vbwProfiler.vbwExecuteLine 554
                        If Cols(Col).Items < RowCount Then
vbwProfiler.vbwExecuteLine 555
                            .Col = Col
vbwProfiler.vbwExecuteLine 556
                            Exit For
                        End If
vbwProfiler.vbwExecuteLine 557 'B
                    End If
vbwProfiler.vbwExecuteLine 558 'B
vbwProfiler.vbwExecuteLine 559
                Next Col
            End If
vbwProfiler.vbwExecuteLine 560 'B
        End If
vbwProfiler.vbwExecuteLine 561 'B

'If no Col Group found, get First free col
vbwProfiler.vbwExecuteLine 562
        If .Col = 0 Then
vbwProfiler.vbwExecuteLine 563
            For Col = 1 To ColCountFree
vbwProfiler.vbwExecuteLine 564
                If Cols(Col).Items = 0 Then
'.Group is created by ResetCols
vbwProfiler.vbwExecuteLine 565
                    .Col = Col
vbwProfiler.vbwExecuteLine 566
                    Exit For
                End If
vbwProfiler.vbwExecuteLine 567 'B
vbwProfiler.vbwExecuteLine 568
            Next Col
        End If
vbwProfiler.vbwExecuteLine 569 'B

'If a Class flag see if we can place it in a free column but lower row
'Should only happen on initial load
vbwProfiler.vbwExecuteLine 570
        If .Col = 0 Then
vbwProfiler.vbwExecuteLine 571
            For Row = 2 To RowCount
vbwProfiler.vbwExecuteLine 572
                For Col = 1 To ColCountFree
vbwProfiler.vbwExecuteLine 573
                    If Cols(Col).Group = SignalAttributes(Idx).Group Then
vbwProfiler.vbwExecuteLine 574
                        If Cols(Col).Items < RowCount Then
vbwProfiler.vbwExecuteLine 575
                            .Col = Col
vbwProfiler.vbwExecuteLine 576
                            .Row = Row
vbwProfiler.vbwExecuteLine 577
                            Exit For
                        End If
vbwProfiler.vbwExecuteLine 578 'B
                    End If
vbwProfiler.vbwExecuteLine 579 'B
vbwProfiler.vbwExecuteLine 580
                If .Col > 0 Then
vbwProfiler.vbwExecuteLine 581
                     Exit For
                End If
vbwProfiler.vbwExecuteLine 582 'B
vbwProfiler.vbwExecuteLine 583
                Next Col
vbwProfiler.vbwExecuteLine 584
            If .Col > 0 Then
vbwProfiler.vbwExecuteLine 585
                 Exit For
            End If
vbwProfiler.vbwExecuteLine 586 'B
vbwProfiler.vbwExecuteLine 587
            Next Row
        End If
vbwProfiler.vbwExecuteLine 588 'B

'On initial load place in any free slot
vbwProfiler.vbwExecuteLine 589
        If .Col = 0 Then
vbwProfiler.vbwExecuteLine 590
            For Row = 1 To RowCount
vbwProfiler.vbwExecuteLine 591
                For Col = 1 To ColCount
vbwProfiler.vbwExecuteLine 592
                    If Cols(Col).Items < RowCount Then
vbwProfiler.vbwExecuteLine 593
                        .Col = Col
vbwProfiler.vbwExecuteLine 594
                        .Row = Cols(Col).Items + 1
vbwProfiler.vbwExecuteLine 595
                        Exit For
                    End If
vbwProfiler.vbwExecuteLine 596 'B
vbwProfiler.vbwExecuteLine 597
                If .Col > 0 Then
vbwProfiler.vbwExecuteLine 598
                     Exit For
                End If
vbwProfiler.vbwExecuteLine 599 'B
vbwProfiler.vbwExecuteLine 600
                Next Col
vbwProfiler.vbwExecuteLine 601
            If .Col > 0 Then
vbwProfiler.vbwExecuteLine 602
                 Exit For
            End If
vbwProfiler.vbwExecuteLine 603 'B
vbwProfiler.vbwExecuteLine 604
            Next Row
        End If
vbwProfiler.vbwExecuteLine 605 'B

vbwProfiler.vbwExecuteLine 606
        If .Col = 0 Then
vbwProfiler.vbwExecuteLine 607
MsgBox "No free Cols", vbCritical, "NextFreeGroupFlagPos"
vbwProfiler.vbwProcOut 37
vbwProfiler.vbwExecuteLine 608
            Exit Function
        End If
vbwProfiler.vbwExecuteLine 609 'B

vbwProfiler.vbwExecuteLine 610
        If .Row = 0 Then
vbwProfiler.vbwExecuteLine 611
            If .FixedRow > 0 Then
vbwProfiler.vbwExecuteLine 612
                .Row = .FixedRow
            Else
vbwProfiler.vbwExecuteLine 613 'B
vbwProfiler.vbwExecuteLine 614
                .Row = Cols(.Col).Items + 1
            End If
vbwProfiler.vbwExecuteLine 615 'B
        End If
vbwProfiler.vbwExecuteLine 616 'B
vbwProfiler.vbwExecuteLine 617
    End With
'Debug.Print "NextPos=" & NextFreeGroupFlagPos & " (" & Row & "," & Col & ")"
vbwProfiler.vbwProcOut 37
vbwProfiler.vbwExecuteLine 618
End Function

Private Function DequeCmd(Optional Group As String) As Integer
vbwProfiler.vbwProcIn 38
Dim i As Long
vbwProfiler.vbwExecuteLine 619
    For i = 0 To UBound(CmdQ)
vbwProfiler.vbwExecuteLine 620
        If CmdQ(i) <> 0 Then
vbwProfiler.vbwExecuteLine 621
            If Group = "" Or SignalAttributes(CmdQ(i)).Group = Group Then
vbwProfiler.vbwExecuteLine 622
                If DequeCmd = 0 Then
vbwProfiler.vbwExecuteLine 623
                    DequeCmd = CmdQ(i)
vbwProfiler.vbwExecuteLine 624
Debug.Print "Deque " & SignalAttributes(CmdQ(i)).Name & " (" & Group & ")"
vbwProfiler.vbwExecuteLine 625
                    Commands(CmdQ(i)).BackColor = cbDefault
vbwProfiler.vbwExecuteLine 626
                    CmdQ(i) = 0
                End If
vbwProfiler.vbwExecuteLine 627 'B
            End If
vbwProfiler.vbwExecuteLine 628 'B
        End If
vbwProfiler.vbwExecuteLine 629 'B
'Shift remaining commands up the queue
vbwProfiler.vbwExecuteLine 630
        If DequeCmd <> 0 Then
vbwProfiler.vbwExecuteLine 631
            If i = UBound(CmdQ) Then
vbwProfiler.vbwExecuteLine 632
                CmdQ(i) = 0
            Else
vbwProfiler.vbwExecuteLine 633 'B
vbwProfiler.vbwExecuteLine 634
                CmdQ(i) = CmdQ(i + 1)
            End If
vbwProfiler.vbwExecuteLine 635 'B
        End If
vbwProfiler.vbwExecuteLine 636 'B
vbwProfiler.vbwExecuteLine 637
    Next i

'Stop
vbwProfiler.vbwProcOut 38
vbwProfiler.vbwExecuteLine 638
End Function

Private Function QueueRequest(Idx As Long)
vbwProfiler.vbwProcIn 39
Dim i As Long

vbwProfiler.vbwExecuteLine 639
    For i = 0 To UBound(CmdQ)
vbwProfiler.vbwExecuteLine 640
        If CmdQ(i) = 0 Then
vbwProfiler.vbwExecuteLine 641
            CmdQ(i) = Idx
vbwProfiler.vbwExecuteLine 642
            Commands(Idx).BackColor = vbCyan
vbwProfiler.vbwExecuteLine 643
Debug.Print "Queue " & SignalAttributes(CmdQ(i)).Name
vbwProfiler.vbwProcOut 39
vbwProfiler.vbwExecuteLine 644
            Exit Function
        Else
vbwProfiler.vbwExecuteLine 645 'B
'Only q the same command once (must not queue Recall more than once)
vbwProfiler.vbwExecuteLine 646
            If CmdQ(i) = Idx Then
vbwProfiler.vbwProcOut 39
vbwProfiler.vbwExecuteLine 647
                 Exit Function
            End If
vbwProfiler.vbwExecuteLine 648 'B
        End If
vbwProfiler.vbwExecuteLine 649 'B
vbwProfiler.vbwExecuteLine 650
    Next i
'MsgBox "Command Queue is full (" & UBound(CmdQ) & ") maximum"
vbwProfiler.vbwProcOut 39
vbwProfiler.vbwExecuteLine 651
End Function

Private Function FinishTime()
vbwProfiler.vbwProcIn 40
vbwProfiler.vbwExecuteLine 652
    With mshFinish
'not the first (blank) row
vbwProfiler.vbwExecuteLine 653
        If .TextMatrix(.Rows - 1, 0) <> "" Then
vbwProfiler.vbwExecuteLine 654
            .Rows = .Rows + 1
        End If
vbwProfiler.vbwExecuteLine 655 'B
vbwProfiler.vbwExecuteLine 656
        .TextMatrix(.Rows - 1, 0) = .Rows - 1
vbwProfiler.vbwExecuteLine 657
        .TextMatrix(.Rows - 1, 1) = lblCurrTime.Caption
'Scroll to bottom
vbwProfiler.vbwExecuteLine 658
        .TopRow = .Rows - 1
vbwProfiler.vbwExecuteLine 659
End With

vbwProfiler.vbwProcOut 40
vbwProfiler.vbwExecuteLine 660
End Function

Private Function FlagVisibility(ByVal Idx As Long, Visible As Boolean)
vbwProfiler.vbwProcIn 41
Dim Pos As Long
Dim Cidx As Long
vbwProfiler.vbwExecuteLine 661
    Pos = SignalAttributes(Idx).Flag.Pos
'See if visiblility has changed (To generate Controller event)
vbwProfiler.vbwExecuteLine 662
    If Pos > 0 Then
vbwProfiler.vbwExecuteLine 663
        If Flags(Pos).Visible <> Visible Then
vbwProfiler.vbwExecuteLine 664
            Flags(Pos).Visible = Visible
vbwProfiler.vbwExecuteLine 665
            Cidx = SignalAttributes(Idx).Controller
vbwProfiler.vbwExecuteLine 666
            If Cidx <> -1 Then
vbwProfiler.vbwExecuteLine 667
                With Controllers(Cidx)
vbwProfiler.vbwExecuteLine 668
                    If Visible Then
'Debug.Print .Connection & "(" & Cidx & ")" & .On
vbwProfiler.vbwExecuteLine 669
                        If .Sound <> "" Then
vbwProfiler.vbwExecuteLine 670
                             Call PlayWav(.Sound)
                        End If
vbwProfiler.vbwExecuteLine 671 'B
'Call Beep(300, CInt(SignalAttributes(Idx).TTL))
                    Else
vbwProfiler.vbwExecuteLine 672 'B
'Debug.Print .Connection & "(" & Cidx & ")" & .Off
vbwProfiler.vbwExecuteLine 673
                        If .Sound <> "" Then
vbwProfiler.vbwExecuteLine 674
                             Call StopWav
                        End If
vbwProfiler.vbwExecuteLine 675 'B
                    End If
vbwProfiler.vbwExecuteLine 676 'B
vbwProfiler.vbwExecuteLine 677
                End With
            End If
vbwProfiler.vbwExecuteLine 678 'B
        End If
vbwProfiler.vbwExecuteLine 679 'B
    Else
vbwProfiler.vbwExecuteLine 680 'B
vbwProfiler.vbwExecuteLine 681
        MsgBox "Flag " & SignalAttributes(Idx).Name & " not Raised", vbCritical, "FlagVisibility"
    End If
vbwProfiler.vbwExecuteLine 682 'B
vbwProfiler.vbwProcOut 41
vbwProfiler.vbwExecuteLine 683
End Function

Private Function ResetCols()
vbwProfiler.vbwProcIn 42
Dim Idx As Long
Dim Col As Long

vbwProfiler.vbwExecuteLine 684
    ReDim Cols(ColCount)
vbwProfiler.vbwExecuteLine 685
    For Idx = 1 To UBound(SignalAttributes)
vbwProfiler.vbwExecuteLine 686
        With SignalAttributes(Idx)
vbwProfiler.vbwExecuteLine 687
            If .Flag.Col > 0 And .Flag.Row = 1 Then
vbwProfiler.vbwExecuteLine 688
                Cols(.Flag.Col).Group = .Group
            End If
vbwProfiler.vbwExecuteLine 689 'B
vbwProfiler.vbwExecuteLine 690
            If .Flag.FixedCol > 0 Then
vbwProfiler.vbwExecuteLine 691
                Cols(.Flag.FixedCol).Group = .Group
            End If
vbwProfiler.vbwExecuteLine 692 'B
vbwProfiler.vbwExecuteLine 693
            If SignalAttributes(Idx).Flag.Col > 0 Then
vbwProfiler.vbwExecuteLine 694
                Cols(.Flag.Col).Items = Cols(.Flag.Col).Items + 1
            End If
vbwProfiler.vbwExecuteLine 695 'B
vbwProfiler.vbwExecuteLine 696
        End With
vbwProfiler.vbwExecuteLine 697
    Next Idx
vbwProfiler.vbwProcOut 42
vbwProfiler.vbwExecuteLine 698
End Function

'Used to Check if a Class Flag is up when Recall is asked for
'If 2 Class flags are up it will select the lowest class (Idx is in class order)
Private Function GroupIdx(ByVal Group As String) As Long
vbwProfiler.vbwProcIn 43
Dim Idx As Long
vbwProfiler.vbwExecuteLine 699
    For Idx = 1 To UBound(SignalAttributes)
vbwProfiler.vbwExecuteLine 700
        With SignalAttributes(Idx)
vbwProfiler.vbwExecuteLine 701
            If .Group = Group And .Flag.Pos > 0 Then
vbwProfiler.vbwExecuteLine 702
                 GroupIdx = Idx
vbwProfiler.vbwExecuteLine 703
                Exit For
            End If
vbwProfiler.vbwExecuteLine 704 'B
vbwProfiler.vbwExecuteLine 705
        End With
vbwProfiler.vbwExecuteLine 706
    Next Idx
vbwProfiler.vbwProcOut 43
vbwProfiler.vbwExecuteLine 707
End Function


Private Function CompressCols()
vbwProfiler.vbwProcIn 44
Dim LowestFixedCol As Long
Dim Idx As Long
Dim Col As Long
Dim Row As Long
Dim Pos As Long
Dim PosFrom As Long
Dim PosTo As Long
'Ensure Cols() is correct
vbwProfiler.vbwExecuteLine 708
    Call ResetCols
vbwProfiler.vbwExecuteLine 709
    For Idx = 1 To UBound(SignalAttributes)
vbwProfiler.vbwExecuteLine 710
        With SignalAttributes(Idx).Flag
vbwProfiler.vbwExecuteLine 711
            If .FixedCol > 0 Then
vbwProfiler.vbwExecuteLine 712
                If LowestFixedCol = 0 Then
vbwProfiler.vbwExecuteLine 713
                     LowestFixedCol = .FixedCol
                End If
vbwProfiler.vbwExecuteLine 714 'B
vbwProfiler.vbwExecuteLine 715
                If LowestFixedCol > .FixedCol Then
vbwProfiler.vbwExecuteLine 716
                    LowestFixedCol = .FixedCol
                End If
vbwProfiler.vbwExecuteLine 717 'B
            End If
vbwProfiler.vbwExecuteLine 718 'B
vbwProfiler.vbwExecuteLine 719
        End With
vbwProfiler.vbwExecuteLine 720
    Next Idx
'Exit Function
'Stop

vbwProfiler.vbwExecuteLine 721
    For Col = 1 To LowestFixedCol - 2
vbwProfiler.vbwExecuteLine 722
        If Cols(Col).Items = 0 And Cols(Col + 1).Items > 0 Then
vbwProfiler.vbwExecuteLine 723
            For Idx = 1 To UBound(SignalAttributes)
vbwProfiler.vbwExecuteLine 724
                With SignalAttributes(Idx).Flag
vbwProfiler.vbwExecuteLine 725
                    If .Col = Col + 1 Then
'Move Flags(pos).Picture
vbwProfiler.vbwExecuteLine 726
                        For Row = 1 To RowCount
'GetPos of Empty Col
vbwProfiler.vbwExecuteLine 727
                            PosTo = RC(Row, Col)
vbwProfiler.vbwExecuteLine 728
                            PosFrom = RC(Row, Col + 1)
'                            Flags(CInt(Pos)).Picture = Flags(CInt(.Pos)).Picture
vbwProfiler.vbwExecuteLine 729
                            Flags(CInt(PosTo)) = Flags(CInt(PosFrom))
vbwProfiler.vbwExecuteLine 730
                            Flags(CInt(PosTo)).Visible = Flags(CInt(PosFrom)).Visible

vbwProfiler.vbwExecuteLine 731
                            Flags(CInt(PosFrom)).Picture = Nothing
vbwProfiler.vbwExecuteLine 732
                            Flags(CInt(PosFrom)).Visible = False
vbwProfiler.vbwExecuteLine 733
                            .Row = Row
vbwProfiler.vbwExecuteLine 734
                            .Col = Col
vbwProfiler.vbwExecuteLine 735
                            .Pos = PosTo
'Stop

'Reset Pos,Col,Row
vbwProfiler.vbwExecuteLine 736
                        Next Row
                    End If
vbwProfiler.vbwExecuteLine 737 'B
vbwProfiler.vbwExecuteLine 738
                End With
vbwProfiler.vbwExecuteLine 739
            Next Idx
        End If
vbwProfiler.vbwExecuteLine 740 'B
vbwProfiler.vbwExecuteLine 741
    Next Col
vbwProfiler.vbwExecuteLine 742
    Call ResetCols
vbwProfiler.vbwProcOut 44
vbwProfiler.vbwExecuteLine 743
End Function

'Shift This Column Down 1 from this Row
Private Function ShiftDown(ByVal Row As Long, ByVal Col As Long)
vbwProfiler.vbwProcIn 45
Dim Idx As Long
Dim PosFrom As Long
Dim PosTo As Long
vbwProfiler.vbwExecuteLine 744
    If Cols(Col).Items = RowCount Then
vbwProfiler.vbwExecuteLine 745
MsgBox "No Free Rows"
    End If
vbwProfiler.vbwExecuteLine 746 'B
vbwProfiler.vbwExecuteLine 747
    For Row = Cols(Col).Items To Row Step -1
vbwProfiler.vbwExecuteLine 748
        PosFrom = RC(Row, Col)
vbwProfiler.vbwExecuteLine 749
        For Idx = 1 To UBound(SignalAttributes)
vbwProfiler.vbwExecuteLine 750
            With SignalAttributes(Idx).Flag
vbwProfiler.vbwExecuteLine 751
                If .Pos = PosFrom Then
vbwProfiler.vbwExecuteLine 752
                    PosTo = RC(Row + 1, Col)
vbwProfiler.vbwExecuteLine 753
                    Flags(CInt(PosTo)) = Flags(CInt(PosFrom))
vbwProfiler.vbwExecuteLine 754
                    Flags(CInt(PosTo)).Visible = Flags(CInt(PosFrom)).Visible
vbwProfiler.vbwExecuteLine 755
                    Flags(CInt(PosFrom)).Picture = Nothing
vbwProfiler.vbwExecuteLine 756
                    Flags(CInt(PosFrom)).Visible = False
'Change Flag Position on SignalAttributes of flag were moving
vbwProfiler.vbwExecuteLine 757
                    .Row = Row + 1
vbwProfiler.vbwExecuteLine 758
                    .Col = Col
vbwProfiler.vbwExecuteLine 759
                    .Pos = PosTo
                End If
vbwProfiler.vbwExecuteLine 760 'B
vbwProfiler.vbwExecuteLine 761
            End With
vbwProfiler.vbwExecuteLine 762
        Next Idx
vbwProfiler.vbwExecuteLine 763
    Next Row

vbwProfiler.vbwProcOut 45
vbwProfiler.vbwExecuteLine 764
End Function
Private Function LinkTempRemove(ByVal Idx As Long, Link As defLink)
vbwProfiler.vbwProcIn 46
Dim Lidx As Long
Dim i As Long
vbwProfiler.vbwExecuteLine 765
    With SignalAttributes(Idx)
vbwProfiler.vbwExecuteLine 766
        If IsArrayInitialised(.Links) Then
vbwProfiler.vbwExecuteLine 767
            For Lidx = 0 To UBound(.Links)
vbwProfiler.vbwExecuteLine 768
                If .Links(Lidx).Temp = True _
                And .Links(Lidx).Flag = Link.Flag _
                And .Links(Lidx).Raise = Link.Raise _
                And .Links(Lidx).Type = Link.Type Then
'Shift down, if not at the bottom of the array
vbwProfiler.vbwExecuteLine 769
                    For i = Lidx To UBound(.Links) - 1
vbwProfiler.vbwExecuteLine 770
                        .Links(i) = .Links(i + 1)
vbwProfiler.vbwExecuteLine 771
                    Next i
'If a link has been removed the Redim the array, withoust the last element
vbwProfiler.vbwExecuteLine 772
                    ReDim Preserve .Links(UBound(.Links) - 1)
vbwProfiler.vbwProcOut 46
vbwProfiler.vbwExecuteLine 773
                Exit Function
                End If
vbwProfiler.vbwExecuteLine 774 'B
vbwProfiler.vbwExecuteLine 775
            Next Lidx
        End If
vbwProfiler.vbwExecuteLine 776 'B
vbwProfiler.vbwExecuteLine 777
    End With
vbwProfiler.vbwProcOut 46
vbwProfiler.vbwExecuteLine 778
End Function


