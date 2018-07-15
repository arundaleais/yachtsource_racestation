VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFlxGd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmMain"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   7665
   StartUpPosition =   3  'Windows Default
   WindowState     =   1  'Minimized
   Begin VB.CommandButton cmdPause 
      Caption         =   "Pause"
      Height          =   375
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   360
      Width           =   1095
   End
   Begin VB.Timer FinishTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6120
      Top             =   720
   End
   Begin VB.CommandButton Commands 
      Caption         =   "Commands"
      Enabled         =   0   'False
      Height          =   375
      Index           =   10
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Timer ClearFlagsTimer 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   7320
      Top             =   480
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
      Left            =   6480
      TabIndex        =   22
      Top             =   0
      Width           =   1095
   End
   Begin VB.Frame fraMain 
      Height          =   4455
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   7455
      Begin VB.Frame Frame3 
         Caption         =   "Elapsed Time"
         Height          =   855
         Left            =   2520
         TabIndex        =   27
         Top             =   1080
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
            TabIndex        =   28
            Top             =   240
            Width           =   960
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Start/Finish Times"
         Height          =   4215
         Left            =   5280
         TabIndex        =   18
         Top             =   120
         Width           =   2055
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshFinish 
            Height          =   3855
            Left            =   120
            TabIndex        =   19
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
         Left            =   2520
         TabIndex        =   17
         Top             =   3480
         Width           =   2655
         Begin VB.CommandButton Commands 
            Caption         =   "Commands"
            Enabled         =   0   'False
            Height          =   375
            Index           =   5
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame fraStart 
         Caption         =   "Start"
         Height          =   1575
         Left            =   2520
         TabIndex        =   15
         Top             =   1920
         Width           =   2655
         Begin VB.CommandButton Commands 
            Caption         =   "Commands"
            Enabled         =   0   'False
            Height          =   495
            Index           =   4
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   360
            Width           =   1095
         End
         Begin VB.CommandButton Commands 
            Caption         =   "Commands"
            Enabled         =   0   'False
            Height          =   375
            Index           =   3
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   1080
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Current Time"
         Height          =   975
         Left            =   2520
         TabIndex        =   13
         Top             =   120
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
            TabIndex        =   14
            Top             =   240
            Width           =   2460
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Horn"
         Height          =   855
         Left            =   120
         TabIndex        =   11
         Top             =   3480
         Width           =   2295
         Begin VB.CommandButton Commands 
            Caption         =   "Commands"
            Enabled         =   0   'False
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
         Left            =   120
         TabIndex        =   7
         Top             =   1920
         Width           =   2295
         Begin VB.Frame fraPostpone 
            Caption         =   "Minutes"
            Height          =   735
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   2055
            Begin VB.TextBox txtPostpone 
               Enabled         =   0   'False
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
               Text            =   "0"
               Top             =   240
               Width           =   735
            End
            Begin VB.Label lblCountDown 
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
               Left            =   120
               TabIndex        =   30
               Top             =   240
               Visible         =   0   'False
               Width           =   960
            End
         End
         Begin VB.CommandButton Commands 
            Caption         =   "Commands"
            Enabled         =   0   'False
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
         Height          =   1815
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   2295
         Begin VB.Frame Frame5 
            Caption         =   "Start Sequence"
            Height          =   735
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   2055
            Begin VB.ComboBox cboProfile 
               Height          =   315
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   21
               Top             =   240
               Width           =   1815
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "First Start Time"
            Height          =   735
            Left            =   120
            TabIndex        =   5
            Top             =   960
            Width           =   2055
            Begin VB.TextBox txtFirstStartTime 
               Enabled         =   0   'False
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
               Width           =   1215
            End
         End
      End
   End
   Begin VB.CommandButton Commands 
      Caption         =   "Commands"
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
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
      Begin VB.Label Label1 
         BackColor       =   &H80000009&
         Caption         =   "Please select a Start Sequence"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   1440
         TabIndex        =   26
         Top             =   240
         Width           =   3255
      End
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
      Top             =   6060
      Width           =   7665
      _ExtentX        =   13520
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   8334
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Bevel           =   0
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
Private BaseWidth As Single 'Form Width with no Commands
Private Cols() As defCols
Private NextFreeCol As Long
Private NextCommandTop As Single    'To set Command button positions in LoadProfile
Private FirstStartTime As Date  'To calculate elapsed time (MUST NOT BE CHANGED)
                                'because it will upset the Offset calculation and
                                'cause the catchup to not work
Private LastTimeOutput As Date  'Used for catch-up
'Private PausedSecs As Long
Private CmdQ(8) As Integer     'Idx of next signal (if timer)
Private LastHoist As String  'Group of last flag hoisted (Timer Suppresses sound signal)
Private ClassStart As Boolean   'True if there's been a Class Start at this Elapsed Time

Private FinishCount As Long     'No of Finished Clocked
Private FinishSignalCount As Long   'No of FinishSignals made
'Private PostponeIdx As Long 'The Current Postpone Class - changes at the start
'Private PostponeClass As Long   'The First Class that will be postponed, if the
                                'Postpone Flag is raised.
                                'It is the next class to start
Private RecallIdx As Long   'The Recall Class Flag Set when
Private EventTime As Long   'The time used for the current event we are processing
                            'This is passed to DoTimerEvents for EVERY second
                            'In between events (when a Command is clicked)
                            'this is the LastEventTime
Private PostponeCountDown As Long   'Seconds before Postpone Flag will be dropped

'Private Paused As Boolean

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
'            MsgBox "Please select a Start Sequence", vbExclamation, "No Start Sequence"
vbwProfiler.vbwExecuteLine 22
            WindowState = vbNormal
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

'Only used to clear all the flags off the display 3 secs after loading the profile
Private Sub ClearFlagsTimer_Timer()
vbwProfiler.vbwProcIn 3
vbwProfiler.vbwExecuteLine 28
    Loading = True
vbwProfiler.vbwExecuteLine 29
    Call DefaultsPreStartTimeSet
vbwProfiler.vbwExecuteLine 30
    Loading = False
'Must do after defaults to stop LastStart flag being set
vbwProfiler.vbwExecuteLine 31
    ClearFlagsTimer.Enabled = False
vbwProfiler.vbwProcOut 3
vbwProfiler.vbwExecuteLine 32
End Sub

Private Sub cmdEvents_Click()
vbwProfiler.vbwProcIn 4
vbwProfiler.vbwExecuteLine 33
    If frmEvents.Visible Then
vbwProfiler.vbwExecuteLine 34
        frmEvents.Visible = False
    Else
vbwProfiler.vbwExecuteLine 35 'B
vbwProfiler.vbwExecuteLine 36
        Call frmEvents.ListEvents
vbwProfiler.vbwExecuteLine 37
        frmEvents.Visible = True
    End If
vbwProfiler.vbwExecuteLine 38 'B
vbwProfiler.vbwProcOut 4
vbwProfiler.vbwExecuteLine 39
End Sub

'Private Sub cmdPause_Click()
'    With cmdPause
'        Select Case .BackColor
'        Case Is = cbDefault
'            .BackColor = vbGreen
'            Paused = True
'        Case Is = vbGreen
'            .BackColor = vbCyan     'Remove pause on next whole minute
'        End Select
'    End With
'End Sub

Private Sub Commands_Click(Index As Integer)
vbwProfiler.vbwProcIn 5
Dim Position As Long
Dim NextCommand As Long

vbwProfiler.vbwExecuteLine 40
Debug.Print "--- " & Commands(Index).Caption & " ---"
'Clear the message user may be playing with the flags

vbwProfiler.vbwExecuteLine 41
    Label1.Visible = False
vbwProfiler.vbwExecuteLine 42
    With SignalAttributes(Index)
'If this command is queued then just remove it (same as clicking when up)
'This must be done in the Click event because the user is making the request
'You cannot do it in RaiseRequest or LowerRequest because all queued events
'would get removed.
vbwProfiler.vbwExecuteLine 43
        If Commands(Index).BackColor = vbCyan Then
vbwProfiler.vbwExecuteLine 44
            NextCommand = DequeCmd(.Group)
vbwProfiler.vbwProcOut 5
vbwProfiler.vbwExecuteLine 45
            Exit Sub
        End If
vbwProfiler.vbwExecuteLine 46 'B
'If we have another commandButton queued in this group, remove this before
'actioning a raise request so we dont have 2 flags in same group queued
'This is important with Recall & General Recall
vbwProfiler.vbwExecuteLine 47
        If .Flag.Pos = 0 Then
vbwProfiler.vbwExecuteLine 48
            NextCommand = DequeCmd(.Group)
vbwProfiler.vbwExecuteLine 49
            Call RaiseRequest(CLng(Index))
vbwProfiler.vbwExecuteLine 50
            If Commands(Index).Caption = "Postpone" Then
'Causes StartTime to be validated if one has not yet been entered
'Which then causes Events to be reloaded and hence Postponed start time will be set
vbwProfiler.vbwExecuteLine 51
                If StartTimeValid = False Then
vbwProfiler.vbwExecuteLine 52
                    StartTimeValid = True
'                    Call ValidateStartTime

'                    txtFirstStartTime = Format$((DateAdd("n", CDbl(NulToZero(txtPostpone.Text)), FirstStartTime)), "hhnn")
                End If
vbwProfiler.vbwExecuteLine 53 'B
            End If
vbwProfiler.vbwExecuteLine 54 'B
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


Private Sub Commands_LostFocus(Index As Integer)
'    With Commands(Index)
'    Select Case .Caption
'        Case Is = "Postpone", "Recall", "General Recall", "Finish"
'            If .BackColor <> vbCyan Then
'                .BackColor = cbDefault
'            End If
'    End Select
'    End With
vbwProfiler.vbwProcIn 6
vbwProfiler.vbwProcOut 6
vbwProfiler.vbwExecuteLine 60
End Sub

Private Sub Flags_Click(Index As Integer)
'MsgBox Flags(Index).Picture.Handle
vbwProfiler.vbwProcIn 7
vbwProfiler.vbwProcOut 7
vbwProfiler.vbwExecuteLine 61
End Sub

Private Sub Form_Load()
vbwProfiler.vbwProcIn 8
Dim i As Long
Dim url As String
Dim Major As Long
Dim Minor As Long
Dim Revision As Long
Dim NewVersion As Boolean
Dim Cmd As CommandButton

vbwProfiler.vbwExecuteLine 62
    BaseWidth = Width
vbwProfiler.vbwExecuteLine 63
    For Each Cmd In Commands
vbwProfiler.vbwExecuteLine 64
    ReDim Preserve StaticCommands(Cmd.Index)
vbwProfiler.vbwExecuteLine 65
       StaticCommands(Cmd.Index) = True
vbwProfiler.vbwExecuteLine 66
    Next Cmd


vbwProfiler.vbwExecuteLine 67
    Caption = App.EXEName & " [" & App.Major & "." & App.Minor & "." _
    & App.Revision & "] "
vbwProfiler.vbwExecuteLine 68
    Caption = Replace(Caption, ".exe", "")
'Check if a later version exists
vbwProfiler.vbwExecuteLine 69
    url = "http://www.NmeaRouter.com/docs/ais/" & App.EXEName _
    & "_setup_"
vbwProfiler.vbwExecuteLine 70
    Major = App.Major
vbwProfiler.vbwExecuteLine 71
    Do
vbwProfiler.vbwExecuteLine 72
        If HTTPFileExists(url & Major & ".0.0.exe") = False Then
vbwProfiler.vbwExecuteLine 73
             Exit Do
        End If
vbwProfiler.vbwExecuteLine 74 'B
vbwProfiler.vbwExecuteLine 75
        Major = Major + 1
vbwProfiler.vbwExecuteLine 76
    Loop
vbwProfiler.vbwExecuteLine 77
    If Major > 0 Then 'Highest major that exists
vbwProfiler.vbwExecuteLine 78
         Major = Major - 1
    End If
vbwProfiler.vbwExecuteLine 79 'B

vbwProfiler.vbwExecuteLine 80
    url = url & Major & "."
vbwProfiler.vbwExecuteLine 81
    If Major = App.Major Then
vbwProfiler.vbwExecuteLine 82
        Minor = App.Minor
    Else
vbwProfiler.vbwExecuteLine 83 'B
vbwProfiler.vbwExecuteLine 84
        Minor = 0
    End If
vbwProfiler.vbwExecuteLine 85 'B
vbwProfiler.vbwExecuteLine 86
    Do
vbwProfiler.vbwExecuteLine 87
        If HTTPFileExists(url & Minor & ".0.exe") = False Then
vbwProfiler.vbwExecuteLine 88
             Exit Do
        End If
vbwProfiler.vbwExecuteLine 89 'B
vbwProfiler.vbwExecuteLine 90
        Minor = Minor + 1
vbwProfiler.vbwExecuteLine 91
    Loop
vbwProfiler.vbwExecuteLine 92
    If Minor > 0 Then
vbwProfiler.vbwExecuteLine 93
         Minor = Minor - 1
    End If
vbwProfiler.vbwExecuteLine 94 'B

vbwProfiler.vbwExecuteLine 95
    url = url & Minor & "."
vbwProfiler.vbwExecuteLine 96
    If Not (Major = App.Major And Minor = App.Minor) Then
vbwProfiler.vbwExecuteLine 97
        NewVersion = True
    End If
vbwProfiler.vbwExecuteLine 98 'B
'Only let a user get next revision if he is using a revision
'of his current version. Otherwise he goes up to the next minor version
vbwProfiler.vbwExecuteLine 99
    If NewVersion = False And App.Revision > 0 Then
vbwProfiler.vbwExecuteLine 100
        Revision = App.Revision
vbwProfiler.vbwExecuteLine 101
        Do
vbwProfiler.vbwExecuteLine 102
            If HTTPFileExists(url & Revision & ".exe") = False Then
vbwProfiler.vbwExecuteLine 103
                 Exit Do
            End If
vbwProfiler.vbwExecuteLine 104 'B
vbwProfiler.vbwExecuteLine 105
            Revision = Revision + 1
vbwProfiler.vbwExecuteLine 106
        Loop
vbwProfiler.vbwExecuteLine 107
        If Revision > 0 Then
vbwProfiler.vbwExecuteLine 108
             Revision = Revision - 1
        End If
vbwProfiler.vbwExecuteLine 109 'B
vbwProfiler.vbwExecuteLine 110
        If Revision < App.Revision Then
vbwProfiler.vbwExecuteLine 111
            NewVersion = True
        End If
vbwProfiler.vbwExecuteLine 112 'B
    End If
vbwProfiler.vbwExecuteLine 113 'B
vbwProfiler.vbwExecuteLine 114
    url = url & Revision & ".exe"

'If we are working on a higher version in VBE, don't try for newversion
vbwProfiler.vbwExecuteLine 115
    If App.Major * 2 ^ 8 + App.Minor * 2 ^ 4 + App.Revision > _
    Major * 2 ^ 8 + Minor * 2 ^ 4 + Revision Then
vbwProfiler.vbwExecuteLine 116
        NewVersion = False
    End If
vbwProfiler.vbwExecuteLine 117 'B
vbwProfiler.vbwExecuteLine 118
    If NewVersion = True Then
vbwProfiler.vbwExecuteLine 119
        Call frmDpyBox.DpyBox("A new update is available", 10, "New Version")
'Check we have internet access
vbwProfiler.vbwExecuteLine 120
        If HTTPFileExists(url) Then
vbwProfiler.vbwExecuteLine 121
            Call HttpSpawn(url)
        End If
vbwProfiler.vbwExecuteLine 122 'B
    End If
vbwProfiler.vbwExecuteLine 123 'B
'Position cursor at RHS of time displayed
vbwProfiler.vbwExecuteLine 124
    txtFirstStartTime.SelStart = Len(txtFirstStartTime)
vbwProfiler.vbwExecuteLine 125
    txtPostpone.SelStart = Len(txtPostpone)

vbwProfiler.vbwExecuteLine 126
    WindowState = vbNormal
vbwProfiler.vbwExecuteLine 127
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft")
vbwProfiler.vbwExecuteLine 128
    Me.Top = GetSetting(App.Title, "Settings", "MainTop")
'    Me.Width = GetSetting(App.Title, "Settings", "MainWidth")
'    Me.Height = GetSetting(App.Title, "Settings", "MainHeight")
vbwProfiler.vbwExecuteLine 129
    Visible = True

vbwProfiler.vbwExecuteLine 130
    With mshFinish
vbwProfiler.vbwExecuteLine 131
        .Width = 1795
vbwProfiler.vbwExecuteLine 132
        .FormatString = "<No|<Time"
vbwProfiler.vbwExecuteLine 133
        .ColWidth(0) = 500  'Position
vbwProfiler.vbwExecuteLine 134
        .ColWidth(1) = 1295  'Time
'        For i = 1 To 20
'            .Rows = i + 1
'            .TextMatrix(i, 0) = i
'        Next i
'        .TextMatrix(1, 1) = "13:22:45"
vbwProfiler.vbwExecuteLine 135
    End With


'Flags(0) exists - but not used
vbwProfiler.vbwExecuteLine 136
    RowCount = FlagRow(Flags.Count - 1)
vbwProfiler.vbwExecuteLine 137
    ColCount = FlagCol(Flags.Count - 1)
vbwProfiler.vbwExecuteLine 138
    ColCountFree = ColCount 'Reduces by number of Fixed cols
vbwProfiler.vbwExecuteLine 139
    ReDim Cols(1 To ColCount)
vbwProfiler.vbwExecuteLine 140
Visible = True
vbwProfiler.vbwExecuteLine 141
    StatusBar1.Panels(2).Width = 200
vbwProfiler.vbwExecuteLine 142
    StatusBar1.Panels(3).Width = 200
'Make the base index invisible as it is not used
vbwProfiler.vbwExecuteLine 143
    Commands(0).Enabled = False
vbwProfiler.vbwExecuteLine 144
    Commands(0).Visible = False
'Set up initial start time, LoadEvents not called
vbwProfiler.vbwExecuteLine 145
    FirstStartTime = Date & " " _
    & Format$(NulToZero(txtFirstStartTime), "00:00") & ":00"
vbwProfiler.vbwExecuteLine 146
Debug.Print Format$(FirstStartTime, "dd-mmm-yyyy")
vbwProfiler.vbwExecuteLine 147
Debug.Print Format$(FirstStartTime, "hh:mm:ss")

vbwProfiler.vbwExecuteLine 148
    Call LoadSequence

vbwProfiler.vbwProcOut 8
vbwProfiler.vbwExecuteLine 149
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Stop any events generated by DoTimerEvents interupting the Unloading
vbwProfiler.vbwProcIn 9
vbwProfiler.vbwExecuteLine 150
        RaceTimer.Enabled = False
'Clear any Controller Relays on when program is terminated (if connected)
vbwProfiler.vbwExecuteLine 151
        Unload frmDaventech
vbwProfiler.vbwExecuteLine 152
        Call EncryptFiles(Environ("AllUsersProfile") & "\Application Data\Arundale\RacingSignals\Sequences\", ".txt", ".ini")
'    End If
vbwProfiler.vbwProcOut 9
vbwProfiler.vbwExecuteLine 153
End Sub

Private Sub Form_Unload(Cancel As Integer)
vbwProfiler.vbwProcIn 10
    Dim i As Integer

'Must NOT reference frmDaventech as this will cause a reload and open of winsock
'Debug.Print "frmain.Unload " & frmDaventech.Winsock1.State
vbwProfiler.vbwExecuteLine 154
    Me.WindowState = vbNormal
vbwProfiler.vbwExecuteLine 155
    SaveSetting App.Title, "Settings", "MainLeft", Me.Left
vbwProfiler.vbwExecuteLine 156
    SaveSetting App.Title, "Settings", "MainTop", Me.Top
vbwProfiler.vbwExecuteLine 157
    SaveSetting App.Title, "Settings", "MainWidth", Me.Width
vbwProfiler.vbwExecuteLine 158
    SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    'close all sub forms
vbwProfiler.vbwExecuteLine 159
    For i = Forms.Count - 1 To 0 Step -1
vbwProfiler.vbwExecuteLine 160
        Unload Forms(i)
vbwProfiler.vbwExecuteLine 161
    Next
vbwProfiler.vbwExecuteLine 162
    End     'terminate program
vbwProfiler.vbwProcOut 10
vbwProfiler.vbwExecuteLine 163
End Sub

Private Sub HoistTimer_Timer()
vbwProfiler.vbwProcIn 11
vbwProfiler.vbwExecuteLine 164
    HoistTimer.Enabled = False
'Set last hoist to Blank so any subsequent hoist will action the sound signal
vbwProfiler.vbwExecuteLine 165
    LastHoist = ""
'Debug.Print "HoistTimer disabled"
vbwProfiler.vbwProcOut 11
vbwProfiler.vbwExecuteLine 166
End Sub


'Runs every 1 second
Private Sub RaceTimer_Timer()
vbwProfiler.vbwProcIn 12
Dim CurrTime As Date    'May be speeded up from Now() time for testing
Dim SecsSinceOutput As Long
Dim TimeToOutput As Date
Dim SecsToAdd As Long
Dim SecsSinceFirstStart As Long

vbwProfiler.vbwExecuteLine 167
    CurrTime = Now()
vbwProfiler.vbwExecuteLine 168
    lblCurrTime = Format$(CurrTime, "hh:mm:ss")

'Sets LastTimeOutput to to Current Time - 1 sec
vbwProfiler.vbwExecuteLine 169
    If LastTimeOutput = "00:00:00" Then
vbwProfiler.vbwExecuteLine 170
        Call ResetOutput(CurrTime)
    End If
vbwProfiler.vbwExecuteLine 171 'B

vbwProfiler.vbwExecuteLine 172
    SecsSinceOutput = DateDiff("s", LastTimeOutput, CurrTime)
'No Output due yet
vbwProfiler.vbwExecuteLine 173
    If SecsSinceOutput = 0 Then
'Debug.Print "Skip"
vbwProfiler.vbwProcOut 12
vbwProfiler.vbwExecuteLine 174
        Exit Sub
    End If
vbwProfiler.vbwExecuteLine 175 'B

'Goes round this loop once for each second between LastTimeOutput and CurrentTime
vbwProfiler.vbwExecuteLine 176
    Do
vbwProfiler.vbwExecuteLine 177
        TimeToOutput = DateAdd("s", 1, LastTimeOutput)
vbwProfiler.vbwExecuteLine 178
        If TimerOutput(TimeToOutput) = True Then
vbwProfiler.vbwExecuteLine 179
             LastTimeOutput = TimeToOutput
        End If
vbwProfiler.vbwExecuteLine 180 'B
vbwProfiler.vbwExecuteLine 181
        SecsSinceOutput = DateDiff("s", LastTimeOutput, CurrTime)
vbwProfiler.vbwExecuteLine 182
        SecsSinceFirstStart = DateDiff("s", FirstStartTime, CurrTime)
vbwProfiler.vbwExecuteLine 183
        EventTime = SecsSinceFirstStart - SecsSinceOutput
'Debug.Print ElapsedTime - SecsSinceOutput

vbwProfiler.vbwExecuteLine 184
        If StartTimeValid Then '(EventTime)
vbwProfiler.vbwExecuteLine 185
             Call DoTimerEvents
        End If
vbwProfiler.vbwExecuteLine 186 'B
'        lblElapsedTime = aSecToElapsed(ElapsedTime)
'display of ElapsedTime has catch-up secs taken off and PausedSecs
'        lblElapsedTime = aSecToElapsed(ElapsedTime - SecsSinceOutput) ' - PausedSecs)
vbwProfiler.vbwExecuteLine 187
If SecsSinceOutput > 0 Then
vbwProfiler.vbwExecuteLine 188
     Debug.Print "Catch-up " & SecsSinceOutput
End If
vbwProfiler.vbwExecuteLine 189 'B
vbwProfiler.vbwExecuteLine 190
    Loop Until SecsSinceOutput = 0  'Always execute once
vbwProfiler.vbwProcOut 12
vbwProfiler.vbwExecuteLine 191
End Sub


Private Sub ResetOutput(StartTime As Date)
vbwProfiler.vbwProcIn 13
vbwProfiler.vbwExecuteLine 192
        LastTimeOutput = DateAdd("s", -1, StartTime)
vbwProfiler.vbwProcOut 13
vbwProfiler.vbwExecuteLine 193
End Sub

Private Function aSecToElapsed(ByVal Secs As Long) As String
vbwProfiler.vbwProcIn 14
Dim hms As defhms
Dim Sign As Long
Dim aSign As String

'Secs = 3600& * 100&
vbwProfiler.vbwExecuteLine 194
    Sign = Sgn(Secs)    '-1 = -ve, 0 = 0 , +1 = +ve
vbwProfiler.vbwExecuteLine 195
    If Sign = -1 Then
vbwProfiler.vbwExecuteLine 196
        Secs = Secs * Sign 'force +ve
vbwProfiler.vbwExecuteLine 197
        aSign = "-"
    Else
vbwProfiler.vbwExecuteLine 198 'B
vbwProfiler.vbwExecuteLine 199
        aSign = " "
    End If
vbwProfiler.vbwExecuteLine 200 'B
vbwProfiler.vbwExecuteLine 201
    hms.Hour = Int(Secs / 3600&)
vbwProfiler.vbwExecuteLine 202
    Secs = Secs - hms.Hour * 3600&
vbwProfiler.vbwExecuteLine 203
    hms.Min = Int(Secs / 60&)
vbwProfiler.vbwExecuteLine 204
    Secs = Secs - hms.Min * 60&
vbwProfiler.vbwExecuteLine 205
    hms.Sec = Secs
vbwProfiler.vbwExecuteLine 206
    aSecToElapsed = aSign & Format$(hms.Hour, "###")
vbwProfiler.vbwExecuteLine 207
    If Abs(hms.Hour) >= 1 Then
vbwProfiler.vbwExecuteLine 208
         aSecToElapsed = aSecToElapsed & ":"
    End If
vbwProfiler.vbwExecuteLine 209 'B
vbwProfiler.vbwExecuteLine 210
    aSecToElapsed = aSecToElapsed & Format$(hms.Min, "00") _
    & ":" & Format$(hms.Sec, "00")
vbwProfiler.vbwProcOut 14
vbwProfiler.vbwExecuteLine 211
End Function

Private Sub ReloadTimer_Timer()
vbwProfiler.vbwProcIn 15
vbwProfiler.vbwExecuteLine 212
    ReloadTimer.Enabled = False
vbwProfiler.vbwExecuteLine 213
    Call LoadProfile
vbwProfiler.vbwProcOut 15
vbwProfiler.vbwExecuteLine 214
End Sub

Private Sub SignalTimer_Timer(Index As Integer)
vbwProfiler.vbwProcIn 16
Dim FlagIdx  As Long
Dim kb As String
Dim CyclesCompleted As Long
Dim LinkedFlagPos As Long

vbwProfiler.vbwExecuteLine 215
    With SignalAttributes(Index)
vbwProfiler.vbwExecuteLine 216
kb = SignalTimer(Index).Enabled
'Debug.Print Flags(FlagIdx).Visible
'A cycle is completed every time a flag is turned off AFTER it has been on

vbwProfiler.vbwExecuteLine 217
        If .Flag.Pos Then
vbwProfiler.vbwExecuteLine 218
            If Flags(.Flag.Pos).Visible = True Then
vbwProfiler.vbwExecuteLine 219
                .OnCycles = .OnCycles + 1
vbwProfiler.vbwExecuteLine 220
                SignalTimer(Index).Interval = .TTD
            Else
vbwProfiler.vbwExecuteLine 221 'B
vbwProfiler.vbwExecuteLine 222
                SignalTimer(Index).Interval = .TTL
vbwProfiler.vbwExecuteLine 223
                CyclesCompleted = .OnCycles

            End If
vbwProfiler.vbwExecuteLine 224 'B
        Else
vbwProfiler.vbwExecuteLine 225 'B
vbwProfiler.vbwExecuteLine 226
            .OnCycles = .OnCycles + 1
'Terminate Timer & Lower flag
vbwProfiler.vbwExecuteLine 227
            CyclesCompleted = .OnCycles
'MsgBox "Signal(" & Index & ")." & .Name & " has no associated Flag", vbCritical, "SignalTimer_Timer"
        End If
vbwProfiler.vbwExecuteLine 228 'B
'Debug.Print CyclesCompleted & "(" & Index & ")"

'Continuous
vbwProfiler.vbwExecuteLine 229
        If .CyclesRequired = 0 Then
vbwProfiler.vbwExecuteLine 230
            If Loading = False Then
vbwProfiler.vbwExecuteLine 231
                 CyclesCompleted = -1
            End If
vbwProfiler.vbwExecuteLine 232 'B
        End If
vbwProfiler.vbwExecuteLine 233 'B

vbwProfiler.vbwExecuteLine 234
        If Loading And CyclesCompleted > 5 Then
vbwProfiler.vbwExecuteLine 235
             CyclesCompleted = .CyclesRequired
        End If
vbwProfiler.vbwExecuteLine 236 'B
vbwProfiler.vbwExecuteLine 237
        Select Case CyclesCompleted
'The timer has started but we do not want the Signal Off
'In fact we should not have started it in the first place
'vbwLine 238:        Case Is >= .CyclesRequired
        Case Is >= IIf(vbwProfiler.vbwExecuteLine(238), VBWPROFILER_EMPTY, _
        .CyclesRequired)
'This only occurs when the flag is about to be made invisible
'Turn off Signal, before disabling the timer
'Otherwise MakeSignals will start it again
'Click the command button (set to True) to put the flag down
'Only disable if not Continuous
vbwProfiler.vbwExecuteLine 239
            SignalTimer(Index).Enabled = False
'Must be after timer is disabled
vbwProfiler.vbwExecuteLine 240
            .OnCycles = 0
vbwProfiler.vbwExecuteLine 241
            Call LowerRequest(Index)
'        Commands(Index).Value = True
'Click the command button
'kb = SignalTimer(Index).Enabled    'Must be turned off
'Do this last, so if the timer is called again
'another off will be generated, and the timer will
'not re-start
'Remove this from the queue and re-enable with next signal (if any)
'        Call DequeTimer(Index)
'vbwLine 242:        Case Is > .CyclesRequired
        Case Is > IIf(vbwProfiler.vbwExecuteLine(242), VBWPROFILER_EMPTY, _
        .CyclesRequired)
'Continuous
'vbwLine 243:        Case Is < .CyclesRequired
        Case Is < IIf(vbwProfiler.vbwExecuteLine(243), VBWPROFILER_EMPTY, _
        .CyclesRequired)
'Reverse the Visibility of this flag and do another cycle
'No linked Flags are activated
vbwProfiler.vbwExecuteLine 244
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
vbwProfiler.vbwExecuteLine 245 'B
vbwProfiler.vbwExecuteLine 246
    End With
vbwProfiler.vbwProcOut 16
vbwProfiler.vbwExecuteLine 247
End Sub

Private Sub txtFirstStartTime_Change()
vbwProfiler.vbwProcIn 17
vbwProfiler.vbwExecuteLine 248
    If txtFirstStartTime.Enabled = True Then
'Manually set
vbwProfiler.vbwExecuteLine 249
        If ValidateStartTime = True Then
vbwProfiler.vbwExecuteLine 250
            Label1.Visible = False
vbwProfiler.vbwExecuteLine 251
            Call DefaultsStartTimeSet
        End If
vbwProfiler.vbwExecuteLine 252 'B
    Else
vbwProfiler.vbwExecuteLine 253 'B
'Automatic set (by postpone)
vbwProfiler.vbwExecuteLine 254
        Label1.Visible = False
vbwProfiler.vbwExecuteLine 255
        Call DefaultsStartTimeSet
    End If
vbwProfiler.vbwExecuteLine 256 'B
vbwProfiler.vbwProcOut 17
vbwProfiler.vbwExecuteLine 257
End Sub

Private Sub txtPostpone_Change()
vbwProfiler.vbwProcIn 18
vbwProfiler.vbwExecuteLine 258
    Call ValidatePostponeTime
vbwProfiler.vbwProcOut 18
vbwProfiler.vbwExecuteLine 259
End Sub

Private Function ValidateStartTime() As Boolean
vbwProfiler.vbwProcIn 19
Dim MyElapsedTime As Long

vbwProfiler.vbwExecuteLine 260
    On Error GoTo ValidateStartTime_error
vbwProfiler.vbwExecuteLine 261
    If txtFirstStartTime = "" Then
vbwProfiler.vbwExecuteLine 262
        txtFirstStartTime.BackColor = vbRed
    Else
vbwProfiler.vbwExecuteLine 263 'B
vbwProfiler.vbwExecuteLine 264
        txtFirstStartTime.BackColor = vbWhite
    End If
vbwProfiler.vbwExecuteLine 265 'B
vbwProfiler.vbwExecuteLine 266
    If Len(txtFirstStartTime) = 4 _
        And CLng(NulToZero(txtFirstStartTime)) >= 0 _
        And CLng(NulToZero(txtFirstStartTime)) <= 2400 _
        And IsNumeric(NulToZero(txtFirstStartTime)) = True Then
vbwProfiler.vbwExecuteLine 267
            FirstStartTime = Date & " " _
            & Format$(NulToZero(txtFirstStartTime), "00:00") & ":00"
vbwProfiler.vbwExecuteLine 268
            On Error GoTo 0
vbwProfiler.vbwExecuteLine 269
            MyElapsedTime = DateDiff("s", FirstStartTime, Now())
vbwProfiler.vbwExecuteLine 270
            If IsEvtsInitialised(Evts) Then
vbwProfiler.vbwExecuteLine 271
                If MyElapsedTime > Evts(0).ElapsedTime Then
vbwProfiler.vbwExecuteLine 272
                    GoTo ValidateStartTime_error
                End If
vbwProfiler.vbwExecuteLine 273 'B
            End If
vbwProfiler.vbwExecuteLine 274 'B
vbwProfiler.vbwExecuteLine 275
            txtFirstStartTime.ForeColor = vbBlack
vbwProfiler.vbwExecuteLine 276
                StatusBar1.Panels(1).Text = ""
'Must not only reset the flags because once the start sequence
'has commenced the whole profile should be reloaded
'        Call ResetFlags
vbwProfiler.vbwExecuteLine 277
            ValidateStartTime = True
vbwProfiler.vbwExecuteLine 278
            StartTimeValid = True
vbwProfiler.vbwProcOut 19
vbwProfiler.vbwExecuteLine 279
            Exit Function
    End If
vbwProfiler.vbwExecuteLine 280 'B
ValidateStartTime_error:
vbwProfiler.vbwExecuteLine 281
    StatusBar1.Panels(1).Text = "Start time invalid"
vbwProfiler.vbwExecuteLine 282
    StartTimeValid = False  'Suppress Time Events
vbwProfiler.vbwExecuteLine 283
    txtFirstStartTime.ForeColor = vbRed
vbwProfiler.vbwProcOut 19
vbwProfiler.vbwExecuteLine 284
End Function

Private Function ValidatePostponeTime() As Boolean
vbwProfiler.vbwProcIn 20

vbwProfiler.vbwExecuteLine 285
    On Error GoTo ValidatePostponeTime_error
vbwProfiler.vbwExecuteLine 286
    If txtPostpone = "" Then
vbwProfiler.vbwExecuteLine 287
        txtPostpone.BackColor = vbRed
    Else
vbwProfiler.vbwExecuteLine 288 'B
vbwProfiler.vbwExecuteLine 289
        txtPostpone.BackColor = vbWhite
    End If
vbwProfiler.vbwExecuteLine 290 'B
vbwProfiler.vbwExecuteLine 291
    If IsNumeric(NulToZero(txtPostpone)) = True Then
vbwProfiler.vbwExecuteLine 292
        txtPostpone.ForeColor = vbBlack
vbwProfiler.vbwExecuteLine 293
        ValidatePostponeTime = True
vbwProfiler.vbwProcOut 20
vbwProfiler.vbwExecuteLine 294
        Exit Function
    End If
vbwProfiler.vbwExecuteLine 295 'B
ValidatePostponeTime_error:
vbwProfiler.vbwExecuteLine 296
    txtPostpone.ForeColor = vbRed
vbwProfiler.vbwProcOut 20
vbwProfiler.vbwExecuteLine 297
End Function

Public Function DebugFlagsCheck()
vbwProfiler.vbwProcIn 21
Dim MyImage As Image
Dim Idx As Long
Dim Count As Long
Dim NoImageCount As Long

vbwProfiler.vbwExecuteLine 298
        For Idx = 1 To UBound(SignalAttributes)
vbwProfiler.vbwExecuteLine 299
            If SignalAttributes(Idx).Flag.Pos > 0 Then
'RecallClass may not have been set when check is called
vbwProfiler.vbwExecuteLine 300
                If Not SignalAttributes(Idx).Image Is Nothing Then

vbwProfiler.vbwExecuteLine 301
                    If SignalAttributes(Idx).Image <> Flags(SignalAttributes(Idx).Flag.Pos).Picture Then

'                    Stop
                    End If
vbwProfiler.vbwExecuteLine 302 'B
                End If
vbwProfiler.vbwExecuteLine 303 'B
            End If
vbwProfiler.vbwExecuteLine 304 'B
vbwProfiler.vbwExecuteLine 305
        Next Idx

vbwProfiler.vbwProcOut 21
vbwProfiler.vbwExecuteLine 306
Exit Function
vbwProfiler.vbwExecuteLine 307
    For Each MyImage In frmMain.Flags
vbwProfiler.vbwExecuteLine 308
        Count = 0
vbwProfiler.vbwExecuteLine 309
        For Idx = 1 To UBound(SignalAttributes)
vbwProfiler.vbwExecuteLine 310
            If SignalAttributes(Idx).Flag.Pos = MyImage.Index Then
vbwProfiler.vbwExecuteLine 311
                Count = Count + 1
            End If
vbwProfiler.vbwExecuteLine 312 'B
vbwProfiler.vbwExecuteLine 313
        Next Idx
vbwProfiler.vbwExecuteLine 314
        If MyImage.Picture.Handle > 0 Then
vbwProfiler.vbwExecuteLine 315
            If Count <> 1 Then
vbwProfiler.vbwExecuteLine 316
                 Stop
            End If
vbwProfiler.vbwExecuteLine 317 'B
        Else
vbwProfiler.vbwExecuteLine 318 'B
vbwProfiler.vbwExecuteLine 319
            NoImageCount = Count
        End If
vbwProfiler.vbwExecuteLine 320 'B
vbwProfiler.vbwExecuteLine 321
    Next MyImage

vbwProfiler.vbwProcOut 21
vbwProfiler.vbwExecuteLine 322
End Function

#If False Then
Public Function ResetEvents()
    Set CurrEvent = Nothing
End Function
#End If

'Called when a profile is Loaded
Public Function ResetProfile()
'Set up new profile
vbwProfiler.vbwProcIn 22
vbwProfiler.vbwExecuteLine 323
    Caption = App.EXEName & " [" & App.Major & "." & App.Minor & "." _
    & App.Revision & "] " & cboProfile.List(cboProfile.ListIndex)
vbwProfiler.vbwExecuteLine 324
    Label1.Visible = False
vbwProfiler.vbwExecuteLine 325
    txtPostpone.Enabled = False
'Start a fresh Profile
vbwProfiler.vbwExecuteLine 326
    RaceTimer.Enabled = False
'    PausedSecs = 0
 '   PostponeIdx = 0
vbwProfiler.vbwExecuteLine 327
    RecallIdx = 0
vbwProfiler.vbwExecuteLine 328
    FirstStartTime = Date
vbwProfiler.vbwExecuteLine 329
    PostponeCountDown = 0
vbwProfiler.vbwExecuteLine 330
    Call ResetCommands
vbwProfiler.vbwExecuteLine 331
    Call ResetFinish
vbwProfiler.vbwExecuteLine 332
    Call ResetSignalTimers
vbwProfiler.vbwExecuteLine 333
    Call ResetFlags
vbwProfiler.vbwProcOut 22
vbwProfiler.vbwExecuteLine 334
End Function

Private Function ResetFlags()
vbwProfiler.vbwProcIn 23
Dim MyImage As Image

vbwProfiler.vbwExecuteLine 335
    For Each MyImage In frmMain.Flags
vbwProfiler.vbwExecuteLine 336
        MyImage.Picture = Nothing
vbwProfiler.vbwExecuteLine 337
    Next
'Must set so that when loading profile the Queue Check for Recall does not fail
'With error
vbwProfiler.vbwExecuteLine 338
    RecallSignalIdx = 0
'    LastStart = 0
vbwProfiler.vbwProcOut 23
vbwProfiler.vbwExecuteLine 339
End Function

Private Function ResetCommands()
vbwProfiler.vbwProcIn 24
Dim MyCommand As CommandButton
vbwProfiler.vbwExecuteLine 340
    Width = BaseWidth
vbwProfiler.vbwExecuteLine 341
    For Each MyCommand In Commands
vbwProfiler.vbwExecuteLine 342
        If MyCommand.Index <> 0 Then
vbwProfiler.vbwExecuteLine 343
            MyCommand.Enabled = True
vbwProfiler.vbwExecuteLine 344
            MyCommand.Visible = True
vbwProfiler.vbwExecuteLine 345
            MyCommand.BackColor = cbDefault
        End If
vbwProfiler.vbwExecuteLine 346 'B
vbwProfiler.vbwExecuteLine 347
    Next MyCommand
vbwProfiler.vbwExecuteLine 348
    NextCommandTop = 0
vbwProfiler.vbwExecuteLine 349
    txtFirstStartTime = "0000"
vbwProfiler.vbwExecuteLine 350
    txtFirstStartTime.ForeColor = vbBlack
vbwProfiler.vbwExecuteLine 351
    txtFirstStartTime.BackColor = vbWhite
vbwProfiler.vbwExecuteLine 352
    txtFirstStartTime.Enabled = True
vbwProfiler.vbwExecuteLine 353
    txtPostpone.Enabled = True
vbwProfiler.vbwProcOut 24
vbwProfiler.vbwExecuteLine 354
End Function

Private Function ResetFinish()
vbwProfiler.vbwProcIn 25
Dim Row As Long
Dim Col As Long

vbwProfiler.vbwExecuteLine 355
    FinishCount = 0
vbwProfiler.vbwExecuteLine 356
    FinishSignalCount = 0
vbwProfiler.vbwExecuteLine 357
    With mshFinish
'Clear rows (except 1)
vbwProfiler.vbwExecuteLine 358
        For Row = 2 To .Rows - 1
vbwProfiler.vbwExecuteLine 359
            .RemoveItem 1
vbwProfiler.vbwExecuteLine 360
        Next Row
'Clear Row 1
vbwProfiler.vbwExecuteLine 361
        For Col = 0 To .Cols - 1
vbwProfiler.vbwExecuteLine 362
            .TextMatrix(1, Col) = ""
vbwProfiler.vbwExecuteLine 363
        Next Col
vbwProfiler.vbwExecuteLine 364
    End With
vbwProfiler.vbwProcOut 25
vbwProfiler.vbwExecuteLine 365
End Function

Private Function ResetSignalTimers()
vbwProfiler.vbwProcIn 26
Dim MySignalTimer As Timer
Dim i As Long
vbwProfiler.vbwExecuteLine 366
    For Each MySignalTimer In frmMain.SignalTimer
vbwProfiler.vbwExecuteLine 367
        If MySignalTimer.Index > 0 Then  'Dont delete SignalTimer(0)
vbwProfiler.vbwExecuteLine 368
            Unload MySignalTimer
        End If
vbwProfiler.vbwExecuteLine 369 'B
vbwProfiler.vbwExecuteLine 370
    Next
vbwProfiler.vbwExecuteLine 371
    HoistTimer.Enabled = False
vbwProfiler.vbwExecuteLine 372
    LastHoist = ""
'Debug.Print "HoistTimer disabled"
vbwProfiler.vbwProcOut 26
vbwProfiler.vbwExecuteLine 373
End Function

'Sets the default settings when the Splash screen timer has finished and before
'a valid start time has been entered
Private Function DefaultsPreStartTimeSet()
vbwProfiler.vbwProcIn 27
Dim Idx As Long
vbwProfiler.vbwExecuteLine 374
    RecallSignalIdx = 0  'Remove so that logic for recall is not applied when
                        'the flags are lowered
vbwProfiler.vbwExecuteLine 375
    For Idx = 1 To UBound(SignalAttributes)
vbwProfiler.vbwExecuteLine 376
        With SignalAttributes(Idx)
vbwProfiler.vbwExecuteLine 377
            If SignalAttributes(Idx).Flag.Pos <> 0 Then
vbwProfiler.vbwExecuteLine 378
                Call LowerFlag(Idx)
            End If
vbwProfiler.vbwExecuteLine 379 'B
vbwProfiler.vbwExecuteLine 380
        End With
vbwProfiler.vbwExecuteLine 381
    Next Idx
'Reset
vbwProfiler.vbwExecuteLine 382
    RecallSignalIdx = SignalFromName("Recall Class")
vbwProfiler.vbwExecuteLine 383
    Call DisplayStartTimes
vbwProfiler.vbwExecuteLine 384
    Label1.Caption = "Please Set a Start Time"
vbwProfiler.vbwExecuteLine 385
    Label1.Visible = True
'    txtPostpone.Enabled = False
vbwProfiler.vbwExecuteLine 386
    HoistTimer.Enabled = False
vbwProfiler.vbwExecuteLine 387
    LastHoist = ""  'Group
'    Call ResetRecall
vbwProfiler.vbwExecuteLine 388
    Call ResetCols

vbwProfiler.vbwProcOut 27
vbwProfiler.vbwExecuteLine 389
End Function
'Set the Buttons enabled/disabled when the start time has been set
'Thse will be the settings until the First Event is triggered
Public Function DefaultsStartTimeSet()
vbwProfiler.vbwProcIn 28

'Allow user to enter the Postpone Mins
vbwProfiler.vbwExecuteLine 390
    txtPostpone.Enabled = True

vbwProfiler.vbwExecuteLine 391
    With Commands(CommandFromCaption("Postpone"))
vbwProfiler.vbwExecuteLine 392
        .BackColor = vbGreen
vbwProfiler.vbwExecuteLine 393
        .Enabled = True
vbwProfiler.vbwExecuteLine 394
        .SetFocus
vbwProfiler.vbwExecuteLine 395
        End With
vbwProfiler.vbwExecuteLine 396
    With Commands(CommandFromCaption("Recall"))
vbwProfiler.vbwExecuteLine 397
        .BackColor = cbDefault
vbwProfiler.vbwExecuteLine 398
        .Enabled = False
vbwProfiler.vbwExecuteLine 399
    End With
vbwProfiler.vbwExecuteLine 400
    With Commands(CommandFromCaption("General Recall"))
vbwProfiler.vbwExecuteLine 401
        .BackColor = cbDefault
vbwProfiler.vbwExecuteLine 402
        .Enabled = False
vbwProfiler.vbwExecuteLine 403
    End With
vbwProfiler.vbwExecuteLine 404
    With Commands(CommandFromCaption("Finish"))
vbwProfiler.vbwExecuteLine 405
        .BackColor = cbDefault
vbwProfiler.vbwExecuteLine 406
        .Enabled = False
vbwProfiler.vbwExecuteLine 407
    End With
vbwProfiler.vbwExecuteLine 408
    Call DisplayStartTimes
vbwProfiler.vbwProcOut 28
vbwProfiler.vbwExecuteLine 409
End Function

'Called by DoTimerEvents immediately before first event is triggered
'And when Postpone is Raised
Public Function DefaultsFirstEvent()
vbwProfiler.vbwProcIn 29
Dim Eidx As Long
'When the first event is carried out disable the start time
vbwProfiler.vbwExecuteLine 410
    If txtFirstStartTime.Enabled = True Then
vbwProfiler.vbwExecuteLine 411
        txtFirstStartTime.Enabled = False
    End If
vbwProfiler.vbwExecuteLine 412 'B
vbwProfiler.vbwExecuteLine 413
    Call DefaultsStartTimeSet

vbwProfiler.vbwProcOut 29
vbwProfiler.vbwExecuteLine 414
Exit Function
vbwProfiler.vbwExecuteLine 415
    frmMain.Commands(0).Visible = True
vbwProfiler.vbwExecuteLine 416
    frmMain.Commands(0).Enabled = True
vbwProfiler.vbwExecuteLine 417
    frmMain.Commands(0).SetFocus
vbwProfiler.vbwExecuteLine 418
    frmMain.Commands(0).Visible = False

vbwProfiler.vbwProcOut 29
vbwProfiler.vbwExecuteLine 419
End Function
'Requires MSINET.OCX
'See http://officeone.mvps.org/vba/http_file_exists.html
Public Function HTTPFileExists(ByVal url As String) As Boolean
vbwProfiler.vbwProcIn 30
    Dim S As String
    Dim Exists As Boolean
vbwProfiler.vbwExecuteLine 420
    On Error GoTo Inet1_Error
vbwProfiler.vbwExecuteLine 421
    With Inet1
vbwProfiler.vbwExecuteLine 422
        .RequestTimeout = 5
vbwProfiler.vbwExecuteLine 423
        .Protocol = icHTTP
vbwProfiler.vbwExecuteLine 424
        .url = url
vbwProfiler.vbwExecuteLine 425
        .Execute
'see http://support.microsoft.com/kb/182152 =True doesnt work
'vbwLine 426:        Do While .StillExecuting <> False
        Do While vbwProfiler.vbwExecuteLine(426) Or .StillExecuting <> False
vbwProfiler.vbwExecuteLine 427
            DoEvents
vbwProfiler.vbwExecuteLine 428
        Loop
vbwProfiler.vbwExecuteLine 429
        S = UCase(.GetHeader())
vbwProfiler.vbwExecuteLine 430
        Exists = (InStr(1, S, "200 OK") > 0)
vbwProfiler.vbwExecuteLine 431
        .Cancel 'close therequest
vbwProfiler.vbwExecuteLine 432
    End With
vbwProfiler.vbwExecuteLine 433
    HTTPFileExists = Exists
vbwProfiler.vbwProcOut 30
vbwProfiler.vbwExecuteLine 434
    Exit Function
Inet1_Error:
vbwProfiler.vbwExecuteLine 435
    Select Case Err.Number
'vbwLine 436:    Case Is = icConnectFailed 'No internet connection
    Case Is = IIf(vbwProfiler.vbwExecuteLine(436), VBWPROFILER_EMPTY, _
        icConnectFailed )'No internet connection
    End Select
vbwProfiler.vbwExecuteLine 437 'B
vbwProfiler.vbwExecuteLine 438
    Inet1.Cancel
vbwProfiler.vbwProcOut 30
vbwProfiler.vbwExecuteLine 439
End Function

Public Function HttpSpawn(url As String)
vbwProfiler.vbwProcIn 31
Dim r As Long
Dim Command As String

vbwProfiler.vbwExecuteLine 440
If Environ("windir") <> "" Then
vbwProfiler.vbwExecuteLine 441
    r = ShellExecute(0, "open", url, 0, 0, 1)
Else
vbwProfiler.vbwExecuteLine 442 'B
'try for linux compatibility
vbwProfiler.vbwExecuteLine 443
    Command = "winebrowser " & url & " ""%1"""

vbwProfiler.vbwExecuteLine 444
    Shell (Command)
End If
vbwProfiler.vbwExecuteLine 445 'B
vbwProfiler.vbwProcOut 31
vbwProfiler.vbwExecuteLine 446
End Function

Public Function PositionCommand(Idx As Long)
'You dont need these unless testing this module in VBE
'If you have a break set frmMain is minimised and
'the Scale values will be 0
'Dont leave a blank gap
vbwProfiler.vbwProcIn 32
Dim BaseTop As Single   'Top of first Command

vbwProfiler.vbwExecuteLine 447
    BaseTop = 0 'fraMain.Top
vbwProfiler.vbwExecuteLine 448
    With Commands(Idx)
vbwProfiler.vbwExecuteLine 449
        .Caption = .Caption & "(" & Idx & ")"
vbwProfiler.vbwExecuteLine 450
        If .Visible = True Then
'This will be overwritten with the Name from SignalAttributes
'Align first command with top of main frame
vbwProfiler.vbwExecuteLine 451
            If .Width > ScaleWidth - fraMain.Width Then
vbwProfiler.vbwExecuteLine 452
                Width = Width + .Width
            End If
vbwProfiler.vbwExecuteLine 453 'B
vbwProfiler.vbwExecuteLine 454
            WindowState = vbNormal  'Scale will be 0 in VBE (window is minimized)
vbwProfiler.vbwExecuteLine 455
            .Top = ScaleTop + BaseTop + NextCommandTop
'            If .Top + .Height > BaseTop + fraMain.Height Then
vbwProfiler.vbwExecuteLine 456
            If .Top + .Height > StatusBar1.Top Then
vbwProfiler.vbwExecuteLine 457
                NextCommandTop = 0
vbwProfiler.vbwExecuteLine 458
                Width = Width + .Width
vbwProfiler.vbwExecuteLine 459
                WindowState = vbNormal  'Scale will be 0 in VBE (window is minimized)
vbwProfiler.vbwExecuteLine 460
                .Top = ScaleTop + BaseTop + NextCommandTop
            End If
vbwProfiler.vbwExecuteLine 461 'B
vbwProfiler.vbwExecuteLine 462
            WindowState = vbNormal  'Scale will be 0 in VBE (window is minimized)
vbwProfiler.vbwExecuteLine 463
            .Left = ScaleWidth - .Width
vbwProfiler.vbwExecuteLine 464
            NextCommandTop = NextCommandTop + .Height
        End If
vbwProfiler.vbwExecuteLine 465 'B
vbwProfiler.vbwExecuteLine 466
    End With
vbwProfiler.vbwProcOut 32
vbwProfiler.vbwExecuteLine 467
End Function

'Called by the a Command to Raise a flag
'Must called by the Link (Sound may be clicked, with Sound still running)
'Queues if fixed position and Fixed position in use)
'Queues the the command if HoistTimer is running for this Group
'Queues Recall if ClassFlag is UP
'Actions Linked Flag by calling LinkRequest (If not Queued)
'Starts HoistTimer for this Group, if not Queueable (Flags.Queue=False)

Public Function RaiseRequest(ByVal Idx As Long)
vbwProfiler.vbwProcIn 33
Dim SoundEnabled As Boolean
Dim Pos As Long
Dim QueueSignal As Long
Dim NextCmd As Long
Dim MyLink As defLink
Dim ClassIdx As Long
Dim MyImage As Image
Dim i As Long
Dim PostponeIdx As Long

'Dim PreparatoryIdx As Long

vbwProfiler.vbwExecuteLine 468
    If Idx > UBound(SignalAttributes) Then
vbwProfiler.vbwProcOut 33
vbwProfiler.vbwExecuteLine 469
        Exit Function
    End If
vbwProfiler.vbwExecuteLine 470 'B

'Check if Request requires Queueing or actioning
'If Fixed position and Position is in use
vbwProfiler.vbwExecuteLine 471
    With SignalAttributes(Idx)

vbwProfiler.vbwExecuteLine 472
        If Loading = False Then
vbwProfiler.vbwExecuteLine 473
            Select Case .Name
'vbwLine 474:            Case Is = "Finish"
            Case Is = IIf(vbwProfiler.vbwExecuteLine(474), VBWPROFILER_EMPTY, _
        "Finish")
'A Finish Requires Completely different handling, Only Raise the Link UP event
'Do not RaiseRequest to put Finish Flag up (would toggle finish command actions)
vbwProfiler.vbwExecuteLine 475
            If .Name = "Finish" Then
'A Finish always clocks the time and must give correct no Linked signals
vbwProfiler.vbwExecuteLine 476
                Call FinishTime
vbwProfiler.vbwProcOut 33
vbwProfiler.vbwExecuteLine 477
                Exit Function
            End If
vbwProfiler.vbwExecuteLine 478 'B

'Used After the start if the Recall is changed to General Recall
'I now don't think the user should be allowed to change this as it would
'Cause confusing signals (Recall followed by General Recall)
'It is required before the start to Deque the previous request if the user
'changes the type of recall
'            Case Is = "Recall", "General Recall"
'               Call LowerGroup(.Group)
'Dont exit, we need to action this signal
'Now I think if a recall is called the Other Recall should be disabled
'Providing it is not queued so it is moved to the end
            End Select
vbwProfiler.vbwExecuteLine 479 'B
        End If
vbwProfiler.vbwExecuteLine 480 'B
'Debug.Print "RaiseReq " & .Name
vbwProfiler.vbwExecuteLine 481
        Pos = RC(.Flag.FixedRow, .Flag.FixedCol)
'If the Flag has a fixed position, check if any flag is already in this position
vbwProfiler.vbwExecuteLine 482
        If Pos > 0 And .Flag.Queue = True Then
vbwProfiler.vbwExecuteLine 483
            If Flags(Pos).Picture.Handle <> 0 Then
vbwProfiler.vbwExecuteLine 484
                QueueSignal = Idx
'Debug.Print "Q Flag is UP"
            End If
vbwProfiler.vbwExecuteLine 485 'B
        End If
vbwProfiler.vbwExecuteLine 486 'B

'Queues the the command if HoistTimer is running for this Group
'So linked sound signal not made as another flag will be raised on same col
vbwProfiler.vbwExecuteLine 487
        If .Group = LastHoist And .Flag.Queue Then
vbwProfiler.vbwExecuteLine 488
            QueueSignal = Idx
'Debug.Print "Q Timer On"
        End If
vbwProfiler.vbwExecuteLine 489 'B

'Decide if we need to Queue the Recall
'First though Clear any existing recall
vbwProfiler.vbwExecuteLine 490
        If .Group = "Recall" Then
'You cannot action a recall without a RecallIdx (ie Class to recall)
'It is set when Recall is enabled
vbwProfiler.vbwExecuteLine 491
            If RecallIdx = 0 Then
'Before the start sequence Recall is enabled, so set it if we can from the Class Flag
'This enabled testing manually the recall code
vbwProfiler.vbwExecuteLine 492
                Call RecallSetSignal
'There is no Recall Class Flag set up so exit
vbwProfiler.vbwExecuteLine 493
                If RecallIdx = 0 Then
vbwProfiler.vbwProcOut 33
vbwProfiler.vbwExecuteLine 494
                     Exit Function
                End If
vbwProfiler.vbwExecuteLine 495 'B
            End If
vbwProfiler.vbwExecuteLine 496 'B
vbwProfiler.vbwExecuteLine 497
            If .Name <> "Recall Class" Then
vbwProfiler.vbwExecuteLine 498
                Call RecallChange(Idx)  'Check other Recall not currently raised
            End If
vbwProfiler.vbwExecuteLine 499 'B

vbwProfiler.vbwExecuteLine 500
            If RecallIdx = GetPostponeIdx Then
vbwProfiler.vbwExecuteLine 501
                QueueSignal = Idx
            Else
vbwProfiler.vbwExecuteLine 502 'B

'If we are acioning a General Recall we need to suspend the Start Sequence
'If an Individual Recall the sequence continues
vbwProfiler.vbwExecuteLine 503
                If .Name = "General Recall" Then
'if before start Next Class flag event will not have been actioned
'If after start next class will have been actioned
'If before start sequence ??
'Stop
                End If
vbwProfiler.vbwExecuteLine 504 'B
            End If
vbwProfiler.vbwExecuteLine 505 'B

        End If  'Of Recall
vbwProfiler.vbwExecuteLine 506 'B

'If GeneralRecall is when postpone is raised, queue Postpone
vbwProfiler.vbwExecuteLine 507
        If .Group = "Postpone" Then
vbwProfiler.vbwExecuteLine 508
            If Loading = False Then
vbwProfiler.vbwExecuteLine 509
                If SignalAttributes(SignalFromName("General Recall")).Flag.Pos > 0 Then
vbwProfiler.vbwExecuteLine 510
                    QueueSignal = Idx
                End If
vbwProfiler.vbwExecuteLine 511 'B
            End If
vbwProfiler.vbwExecuteLine 512 'B
        End If
vbwProfiler.vbwExecuteLine 513 'B

vbwProfiler.vbwExecuteLine 514
        If Loading = True Then
vbwProfiler.vbwExecuteLine 515
            QueueSignal = 0
        End If
vbwProfiler.vbwExecuteLine 516 'B

vbwProfiler.vbwExecuteLine 517
        If QueueSignal > 0 Then
vbwProfiler.vbwExecuteLine 518
                Call QueueCmd(QueueSignal)
        Else
vbwProfiler.vbwExecuteLine 519 'B

'Put the Flag up (if not up)
vbwProfiler.vbwExecuteLine 520
            If .Flag.Pos = 0 Then

vbwProfiler.vbwExecuteLine 521
                Call RaiseFlag(Idx)

'Actions Linked Flag by calling LinkRequest (If not Queued)
vbwProfiler.vbwExecuteLine 522
                Call LinkRequest(Idx)

'Start HoistTimer for this Group, if not Queueable (Flags.Queue=False)
'So we dont Create a Second Sound signal
vbwProfiler.vbwExecuteLine 523
                If .Flag.Queue = False Then
vbwProfiler.vbwExecuteLine 524
                    HoistTimer.Enabled = False
vbwProfiler.vbwExecuteLine 525
                    HoistTimer.Enabled = True
vbwProfiler.vbwExecuteLine 526
                    LastHoist = .Group
'Debug.Print "HoistTimer Enabled"
                End If
vbwProfiler.vbwExecuteLine 527 'B

            End If  'Flag was not already up
vbwProfiler.vbwExecuteLine 528 'B
        End If  'Not Queued
vbwProfiler.vbwExecuteLine 529 'B

vbwProfiler.vbwExecuteLine 530
    End With

vbwProfiler.vbwProcOut 33
vbwProfiler.vbwExecuteLine 531
End Function

'Called by RaiseRequest and to action the UP link
Public Function RaiseFlag(ByVal Idx As Long)
vbwProfiler.vbwProcIn 34
Dim Col As Long
Dim Row As Long

'Load Profile-Linked Signals with a higher idx will not have been created
'Debug.Print "Raise " & SignalAttributes(Idx).Name
'Action Command now
'Display Image first (if there is one for this Signal)
vbwProfiler.vbwExecuteLine 532
    With SignalAttributes(Idx)

vbwProfiler.vbwExecuteLine 533
        If Not .Image Is Nothing Then
vbwProfiler.vbwExecuteLine 534
            Call NextFreeGroupFlagPos(Idx)
vbwProfiler.vbwExecuteLine 535
            If .Flag.Pos = 0 And Loading = False Then
vbwProfiler.vbwExecuteLine 536
MsgBox "No free Flag positions", vbCritical, "RaiseFlag"
            End If
vbwProfiler.vbwExecuteLine 537 'B
        End If
vbwProfiler.vbwExecuteLine 538 'B

'If we have a flag position then create it (not set if no Image)
vbwProfiler.vbwExecuteLine 539
        If .Flag.Pos > 0 Then
vbwProfiler.vbwExecuteLine 540
            Flags(.Flag.Pos).Picture = .Image
'You have to set it to False becuase FlagVisibility only reacts to a change
vbwProfiler.vbwExecuteLine 541
            Flags(.Flag.Pos).Visible = False
'Must use flagvisibility to create controller event
vbwProfiler.vbwExecuteLine 542
            Call FlagVisibility(Idx, True)
vbwProfiler.vbwExecuteLine 543
            .Flag.Changed = True

vbwProfiler.vbwExecuteLine 544
            Commands(Idx).BackColor = vbGreen
'May still be a timer even if no image to display
vbwProfiler.vbwExecuteLine 545
            If .TTL > 0 Then
vbwProfiler.vbwExecuteLine 546
                SignalTimer(Idx).Interval = .TTL
vbwProfiler.vbwExecuteLine 547
                SignalTimer(Idx).Enabled = True
            End If
vbwProfiler.vbwExecuteLine 548 'B
        End If
vbwProfiler.vbwExecuteLine 549 'B
vbwProfiler.vbwExecuteLine 550
    End With

vbwProfiler.vbwExecuteLine 551
    Call ResetCols  'Resets Cols().Group & .Items from SignalAttributes

vbwProfiler.vbwProcOut 34
vbwProfiler.vbwExecuteLine 552
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
vbwProfiler.vbwProcIn 35
Dim NextCmd As Integer
Dim i As Long
Dim Pos As Long     '> 0 if flag was up

'Load Profile-Linked Signals with a higher idx will not have been created
vbwProfiler.vbwExecuteLine 553
    If Idx > UBound(SignalAttributes) Then
vbwProfiler.vbwProcOut 35
vbwProfiler.vbwExecuteLine 554
        Exit Function
    End If
vbwProfiler.vbwExecuteLine 555 'B

vbwProfiler.vbwExecuteLine 556
    With SignalAttributes(Idx)
'Keep whether flag was actually up. because this can ve called (by Evts Recall Down)
'even when it was not up
vbwProfiler.vbwExecuteLine 557
        Pos = .Flag.Pos
'Debug.Print "LowerReq " & SignalAttributes(Idx).Name

vbwProfiler.vbwExecuteLine 558
For i = 0 To UBound(CmdQ)
vbwProfiler.vbwExecuteLine 559
    If CmdQ(CInt(i)) <> 0 Then
'Debug.Print "Queued(" & i & ")=" & CmdQ(CInt(i))
    End If
vbwProfiler.vbwExecuteLine 560 'B
vbwProfiler.vbwExecuteLine 561
Next i

'We cant ResetClassStart here beacause it is called even when the flag is not up
'Reset the ElapsedTime and remove any class flags

 'Lower Flag (if Up) then any below the Flag
vbwProfiler.vbwExecuteLine 562
        If SignalAttributes(Idx).Flag.Pos > 0 Then

'This is the entry point when a class is started
vbwProfiler.vbwExecuteLine 563
            If .Group = "Class" Then
'Before the Start is actioned we check if Postpone is UP
vbwProfiler.vbwExecuteLine 564
                If SignalAttributes(SignalFromName("Postpone")).Flag.Pos > 0 Then
'If so we Remove an recall that may be queued
vbwProfiler.vbwExecuteLine 565
                    Call DequeCmd("Recall")
'Remove any Start+Warning Flags up for any classes not yet started
vbwProfiler.vbwExecuteLine 566
                    For i = 1 To UBound(Classes)
vbwProfiler.vbwExecuteLine 567
                        If Classes(i).State = 0 Then
'Check if Class Flag is actually up
vbwProfiler.vbwExecuteLine 568
                            If SignalAttributes(Classes(i).Signal).Flag.Pos > 0 Then
'This is silent
vbwProfiler.vbwExecuteLine 569
                                Call LowerFlag(Classes(i).Signal)
                            End If
vbwProfiler.vbwExecuteLine 570 'B
                        End If
vbwProfiler.vbwExecuteLine 571 'B
vbwProfiler.vbwExecuteLine 572
                    Next i
                End If
vbwProfiler.vbwExecuteLine 573 'B
'Stop
            End If
vbwProfiler.vbwExecuteLine 574 'B
'This appears to be a ISAF one off. If Abandon is a 2 flag hoist then dont sound the LowerSignal
vbwProfiler.vbwExecuteLine 575
            If SignalAttributes(Idx).Group = "Abandon" And Cols(SignalAttributes(Idx).Flag.Col).Items > 1 Then
vbwProfiler.vbwExecuteLine 576
                Call LowerFlag(Idx)
            Else
vbwProfiler.vbwExecuteLine 577 'B
vbwProfiler.vbwExecuteLine 578
                Call LowerFlag(Idx)
vbwProfiler.vbwExecuteLine 579
                Call LinkRequest(Idx)
            End If
vbwProfiler.vbwExecuteLine 580 'B
        End If
vbwProfiler.vbwExecuteLine 581 'B

'Overlapped Position Recall above ClassFlag Requesting Recall
'Dequeues Recall when Any Class Lowered by calling RaiseRequest
vbwProfiler.vbwExecuteLine 582
        If .Group = "Class" Then
vbwProfiler.vbwExecuteLine 583
            NextCmd = DequeCmd("Recall")
'            Call ResetRecall
vbwProfiler.vbwExecuteLine 584
            If NextCmd <> 0 Then
vbwProfiler.vbwExecuteLine 585
                If SignalAttributes(NextCmd).Group = "Recall" Then
vbwProfiler.vbwExecuteLine 586
                    Call RaiseRequest(NextCmd)
                End If
vbwProfiler.vbwExecuteLine 587 'B
            End If
vbwProfiler.vbwExecuteLine 588 'B
'Call CompressCols
        End If
vbwProfiler.vbwExecuteLine 589 'B

'If Postpone is Queued when we Lower General Recall then release it
vbwProfiler.vbwExecuteLine 590
        If .Name = "General Recall" Then
vbwProfiler.vbwExecuteLine 591
            NextCmd = DequeCmd("Postpone")    'Must be class
'            Call ResetRecall
vbwProfiler.vbwExecuteLine 592
            If NextCmd <> 0 Then
vbwProfiler.vbwExecuteLine 593
                If SignalAttributes(NextCmd).Name = "Postpone" Then
vbwProfiler.vbwExecuteLine 594
                    Call RaiseRequest(NextCmd)
                End If
vbwProfiler.vbwExecuteLine 595 'B
            End If
vbwProfiler.vbwExecuteLine 596 'B
'Call CompressCols
        End If
vbwProfiler.vbwExecuteLine 597 'B

'Dequeues any Commands in the same Group Calling RaiseRequest
vbwProfiler.vbwExecuteLine 598
        NextCmd = DequeCmd(.Group)
vbwProfiler.vbwExecuteLine 599
        If NextCmd <> 0 Then
vbwProfiler.vbwExecuteLine 600
            Call RaiseRequest(NextCmd)
        End If
vbwProfiler.vbwExecuteLine 601 'B

vbwProfiler.vbwExecuteLine 602
    End With
vbwProfiler.vbwProcOut 35
vbwProfiler.vbwExecuteLine 603
End Function

'Called by LowerRequest
'Lowers the Flag and any subservient flags
'Does not action any links
Private Function LowerFlag(ByVal Idx As Long)
vbwProfiler.vbwProcIn 36
Dim StartCol As Long
Dim StartRow As Long
Dim Group As String
Dim i As Long
Dim Remove As Boolean
Dim PostponeIdx As Long
Dim PostponeClass As Long

'Debug.Print "LowerFlag " & SignalAttributes(Idx).Name

vbwProfiler.vbwExecuteLine 604
    With SignalAttributes(Idx)
vbwProfiler.vbwExecuteLine 605
        StartCol = .Flag.Col
vbwProfiler.vbwExecuteLine 606
        StartRow = .Flag.Row
vbwProfiler.vbwExecuteLine 607
        Group = .Group
'When any Recall Flag is being lowered Clear the signal
'Only disable when called in DoEvents (to faciltate testing before a start sequence)
'LowerFlag is called when loading to drop all the Splash flags
'We must have lowered the Postpone Flag before we calculate the Restart Offsets
'Because ClassRestart need to know whether to add 1 second which it will do
'if the Restart is at the NEXT second

vbwProfiler.vbwExecuteLine 608
        If Loading = False And StartTimeValid = True Then
vbwProfiler.vbwExecuteLine 609
            If .Name = "Postpone" Then     'stops pause
'When we lower the Flag the Nest Event will be 1 second on from the Current EventTime
'                PostponeClass = NextClassStart(EventTime)
'                Classes(PostponeClass).Postpone = False
'Stop the timer if its dropped early by clicking the Command button
vbwProfiler.vbwExecuteLine 610
                PostponeCountDown = 0
'                PostponeClass = NextClassStart(EventTime + 1)
'                Call ClassRestart(PostponeClass, 0)
            End If
vbwProfiler.vbwExecuteLine 611 'B
        End If
vbwProfiler.vbwExecuteLine 612 'B

'Must do this after GeneralRecall stops pause because it needs the RecallIdx
vbwProfiler.vbwExecuteLine 613
        If .Group = "Recall" Then
vbwProfiler.vbwExecuteLine 614
            If RecallIdx > 0 Then
'Clears RecallClassFlag image and Idx, disables Recall buttons
vbwProfiler.vbwExecuteLine 615
                Call RecallClearSignal
            End If
vbwProfiler.vbwExecuteLine 616 'B
        End If
vbwProfiler.vbwExecuteLine 617 'B
vbwProfiler.vbwExecuteLine 618
   End With

'Calls LowerFlag to Lower any subservient flags WITHOUT actioning any link
'(Only the Link of the TOP flag is actioned by LowerRequest)
vbwProfiler.vbwExecuteLine 619
    For i = 1 To UBound(SignalAttributes)
'If i = 36 Then Stop
vbwProfiler.vbwExecuteLine 620
        With SignalAttributes(i)
vbwProfiler.vbwExecuteLine 621
            If .Group = Group Or (Group = "Class" And .Group = "Preparatory") Then
'If in different col or lower row in same col remove
vbwProfiler.vbwExecuteLine 622
                If .Flag.Col = StartCol And .Flag.Row >= StartRow Then

'Stop first. otherwise Timer will fail when it calls FlagVisibility
vbwProfiler.vbwExecuteLine 623
                    If SignalTimer(i).Enabled = True Then
vbwProfiler.vbwExecuteLine 624
                        SignalTimer(i).Enabled = False
                    End If
vbwProfiler.vbwExecuteLine 625 'B

'Clear the flag (if it exists)
vbwProfiler.vbwExecuteLine 626
                    If Flags(.Flag.Pos).Picture.Handle <> 0 Then
'If .Flag.Pos=0, FlagVisibility reports an error so must do first
vbwProfiler.vbwExecuteLine 627
                        Call FlagVisibility(i, False)
vbwProfiler.vbwExecuteLine 628
                        Flags(.Flag.Pos).Picture = Nothing
                    End If
vbwProfiler.vbwExecuteLine 629 'B
vbwProfiler.vbwExecuteLine 630
                    .Flag.Pos = 0
vbwProfiler.vbwExecuteLine 631
                    .Flag.Col = 0
vbwProfiler.vbwExecuteLine 632
                    .Flag.Row = 0
vbwProfiler.vbwExecuteLine 633
                    Commands(i).BackColor = cbDefault
'Must call the link to remove any Lights linked to non class flags (Postpone)
vbwProfiler.vbwExecuteLine 634
                    .Silent = True
vbwProfiler.vbwExecuteLine 635
                    Call LinkRequest(i)
vbwProfiler.vbwExecuteLine 636
                    .Silent = False
                End If
vbwProfiler.vbwExecuteLine 637 'B
            End If
vbwProfiler.vbwExecuteLine 638 'B
vbwProfiler.vbwExecuteLine 639
        End With
vbwProfiler.vbwExecuteLine 640
    Next i


'Stop Hoist Timer if last Flag up in this Group
vbwProfiler.vbwExecuteLine 641
    If Group = LastHoist Then
vbwProfiler.vbwExecuteLine 642
        HoistTimer.Enabled = False
vbwProfiler.vbwExecuteLine 643
        LastHoist = ""
'Debug.Print "HoistTimer disabled"
    End If
vbwProfiler.vbwExecuteLine 644 'B

vbwProfiler.vbwExecuteLine 645
    Call ResetCols  'Resets Cols().Group & .Items from SignalAttributes

vbwProfiler.vbwProcOut 36
vbwProfiler.vbwExecuteLine 646
End Function

'Calling Flag must be positioned (Up or Down) before LinkRequest is Called
'If HoistTimer for this Group is running (LastHoist = IdxGroup) dont action Link
'If Queueable (Flags.Queue=True) there should not be a link

Private Function LinkRequest(ByVal Idx As Long)
vbwProfiler.vbwProcIn 37
Dim Lidx As Long
Dim MyLink As defLink
Dim Suppress As Boolean

vbwProfiler.vbwExecuteLine 647
    With SignalAttributes(Idx)
vbwProfiler.vbwExecuteLine 648
        If IsLinksInitialised(.Links) Then
vbwProfiler.vbwExecuteLine 649
            For Lidx = 0 To UBound(.Links)
vbwProfiler.vbwExecuteLine 650
                MyLink = .Links(Lidx)
vbwProfiler.vbwExecuteLine 651
                If MyLink.Flag > 0 Then
vbwProfiler.vbwExecuteLine 652
                    Suppress = False
'If MyLink.Flag = 4 Then Stop
vbwProfiler.vbwExecuteLine 653
                    If .Flag.Pos > 0 And MyLink.Type = "UpLink" Then
vbwProfiler.vbwExecuteLine 654
                        Call LinkExecute(Idx, MyLink)
                    End If
vbwProfiler.vbwExecuteLine 655 'B
vbwProfiler.vbwExecuteLine 656
                    If .Flag.Pos = 0 And MyLink.Type = "DownLink" Then
vbwProfiler.vbwExecuteLine 657
                        If SignalAttributes(MyLink.Flag).Group = "Sound" _
                        And .Name = "Finish" And FinishCount > 1 _
                        And SoundOnAllFinishers = False Then
vbwProfiler.vbwExecuteLine 658
                            Suppress = True
                        End If
vbwProfiler.vbwExecuteLine 659 'B
vbwProfiler.vbwExecuteLine 660
                        If Suppress = True Then
vbwProfiler.vbwExecuteLine 661
Debug.Print SignalAttributes(MyLink.Flag).Name & " suppressed"
                        Else
vbwProfiler.vbwExecuteLine 662 'B
vbwProfiler.vbwExecuteLine 663
                            Call LinkExecute(Idx, MyLink)
                        End If
vbwProfiler.vbwExecuteLine 664 'B
                    End If
vbwProfiler.vbwExecuteLine 665 'B
                End If
vbwProfiler.vbwExecuteLine 666 'B
'Stop 'Link execute can delete a links index which causes a subscript error
'Change for to a loop with mo0re checking
vbwProfiler.vbwExecuteLine 667
            Next Lidx
        End If
vbwProfiler.vbwExecuteLine 668 'B
vbwProfiler.vbwExecuteLine 669
    End With
vbwProfiler.vbwProcOut 37
vbwProfiler.vbwExecuteLine 670
End Function

'IDx is the Signal containing the Link to Link
Private Function LinkExecute(Idx As Long, Link As defLink)
vbwProfiler.vbwProcIn 38
Dim LinkRejected As String
Dim Silent As Boolean

vbwProfiler.vbwExecuteLine 671
        With Link  'Raising Signal

'On ProfileLoad the linked flag may not have been created yet
'.Name is cleared when the Hoist Timer has finished its cycle (5 secs)
vbwProfiler.vbwExecuteLine 672
            If .Flag <> 0 And .Flag <= UBound(SignalAttributes) Then
vbwProfiler.vbwExecuteLine 673
Debug.Print "Link " & SignalAttributes(Idx).Name & " > " & SignalAttributes(.Flag).Name
'Only on raise because we have to action the downlink (White) when postpone is dropped
'within 10 secs
vbwProfiler.vbwExecuteLine 674
                If SignalAttributes(Idx).Group = LastHoist And .Raise = True Then
vbwProfiler.vbwExecuteLine 675
                    LinkRejected = "Suppressed, LastHoist(" & LastHoist & ")"
                End If
vbwProfiler.vbwExecuteLine 676 'B
vbwProfiler.vbwExecuteLine 677
                If SignalAttributes(Idx).Silent = True And SignalAttributes(.Flag).Group = "Sound" Then
vbwProfiler.vbwExecuteLine 678
                    LinkRejected = "Silenced"
                End If
vbwProfiler.vbwExecuteLine 679 'B
vbwProfiler.vbwExecuteLine 680
                If LinkRejected = "" Then
vbwProfiler.vbwExecuteLine 681
                    If .Raise = True Then   'Raise Linked flag
vbwProfiler.vbwExecuteLine 682
                        Call RaiseRequest(.Flag)
                    Else
vbwProfiler.vbwExecuteLine 683 'B
vbwProfiler.vbwExecuteLine 684
                        Call LowerRequest(.Flag)   'Lower Linked flag
                    End If
vbwProfiler.vbwExecuteLine 685 'B
                Else
vbwProfiler.vbwExecuteLine 686 'B
vbwProfiler.vbwExecuteLine 687
Debug.Print LinkRejected
                End If
vbwProfiler.vbwExecuteLine 688 'B
            Else
vbwProfiler.vbwExecuteLine 689 'B
'There are no Linked Flags to this Flag
'Debug.Print "Link " & SignalAttributes(Idx).Name & " > none"
            End If
vbwProfiler.vbwExecuteLine 690 'B
vbwProfiler.vbwExecuteLine 691
        End With
vbwProfiler.vbwProcOut 38
vbwProfiler.vbwExecuteLine 692
End Function

Private Function RC(ByVal Row As Long, ByVal Col As Long) As Long
'Both must be valid as a pair
vbwProfiler.vbwProcIn 39
vbwProfiler.vbwExecuteLine 693
    If Row > 0 And Col > 0 Then
vbwProfiler.vbwExecuteLine 694
        RC = (Row - 1) * 10 + Col
    End If
vbwProfiler.vbwExecuteLine 695 'B
vbwProfiler.vbwProcOut 39
vbwProfiler.vbwExecuteLine 696
End Function

Private Function FlagRow(ByVal Pos As Long) As Long
vbwProfiler.vbwProcIn 40
vbwProfiler.vbwExecuteLine 697
    If Pos > 0 Then
vbwProfiler.vbwExecuteLine 698
        FlagRow = (Pos - 1) \ 10 + 1
    End If
vbwProfiler.vbwExecuteLine 699 'B
vbwProfiler.vbwProcOut 40
vbwProfiler.vbwExecuteLine 700
End Function
    
Private Function FlagCol(ByVal Pos As Long) As Long
vbwProfiler.vbwProcIn 41
vbwProfiler.vbwExecuteLine 701
    If Pos > 0 Then
vbwProfiler.vbwExecuteLine 702
        FlagCol = Pos - (FlagRow(Pos) - 1) * 10
    End If
vbwProfiler.vbwExecuteLine 703 'B
vbwProfiler.vbwProcOut 41
vbwProfiler.vbwExecuteLine 704
End Function

'Called when Raising Flag, SignalAttributes Col & Row = 0 if no Position available
Private Function NextFreeGroupFlagPos(ByVal Idx As Long)
vbwProfiler.vbwProcIn 42
Dim Col As Long
Dim Row As Long
Dim Pos As Long
Dim ClassIdx As Long

'If we do not have a set position see if this flag has a parent
'ie a 2 flag hoist and the parent flag is up

'    Call ResetCols
'If Idx = 9 Then Stop
vbwProfiler.vbwExecuteLine 705
   With SignalAttributes(Idx).Flag
'Get the Column first
vbwProfiler.vbwExecuteLine 706
        If .FixedCol > 0 Then
vbwProfiler.vbwExecuteLine 707
            .Col = .FixedCol
        End If
vbwProfiler.vbwExecuteLine 708 'B

'See if this flag wants placing in same col as the first Class Flag
'DONT REMOVE may want to use it later
vbwProfiler.vbwExecuteLine 709
            If .Col = 0 Then
vbwProfiler.vbwExecuteLine 710
            Select Case SignalAttributes(Idx).Group
'vbwLine 711:            Case Is = "Preparatory", "Shortened"
            Case Is = IIf(vbwProfiler.vbwExecuteLine(711), VBWPROFILER_EMPTY, _
        "Preparatory"), "Shortened"
'Not Recall as next Class flag may be up
vbwProfiler.vbwExecuteLine 712
                   ClassIdx = GroupIdx("Class")
vbwProfiler.vbwExecuteLine 713
                    If ClassIdx > 0 Then
'Put flag in same col
vbwProfiler.vbwExecuteLine 714
Debug.Print "Top Row"
vbwProfiler.vbwExecuteLine 715
                        .Col = SignalAttributes(ClassIdx).Flag.Col
vbwProfiler.vbwExecuteLine 716
                        .Row = Cols(.Col).Items + 1   '1st free row
'                        Call ShiftDown(.Row, .Col)
                    End If
vbwProfiler.vbwExecuteLine 717 'B
            End Select
vbwProfiler.vbwExecuteLine 718 'B
        End If
vbwProfiler.vbwExecuteLine 719 'B

vbwProfiler.vbwExecuteLine 720
        If .Col = 0 Then
'See if we have a flag Raised in this Group with a spare Row available
vbwProfiler.vbwExecuteLine 721
            If Left$(SignalAttributes(Idx).Name, 6) <> "Class " Then
'Class Flags are always in separate cols (Keep in the same group)
vbwProfiler.vbwExecuteLine 722
                For Col = 1 To ColCountFree
vbwProfiler.vbwExecuteLine 723
                    If Cols(Col).Group = SignalAttributes(Idx).Group Then
vbwProfiler.vbwExecuteLine 724
                        If Cols(Col).Items < RowCount Then
vbwProfiler.vbwExecuteLine 725
                            .Col = Col
vbwProfiler.vbwExecuteLine 726
                            Exit For
                        End If
vbwProfiler.vbwExecuteLine 727 'B
                    End If
vbwProfiler.vbwExecuteLine 728 'B
vbwProfiler.vbwExecuteLine 729
                Next Col
            End If
vbwProfiler.vbwExecuteLine 730 'B
        End If
vbwProfiler.vbwExecuteLine 731 'B

'If no Col Group found, get First free col
vbwProfiler.vbwExecuteLine 732
        If .Col = 0 Then
vbwProfiler.vbwExecuteLine 733
            For Col = 1 To ColCountFree
vbwProfiler.vbwExecuteLine 734
                If Cols(Col).Items = 0 Then
'.Group is created by ResetCols
vbwProfiler.vbwExecuteLine 735
                    .Col = Col
vbwProfiler.vbwExecuteLine 736
                    Exit For
                End If
vbwProfiler.vbwExecuteLine 737 'B
vbwProfiler.vbwExecuteLine 738
            Next Col
        End If
vbwProfiler.vbwExecuteLine 739 'B

'If a Class flag see if we can place it in a free column but lower row
'Should only happen on initial load
vbwProfiler.vbwExecuteLine 740
        If .Col = 0 Then
vbwProfiler.vbwExecuteLine 741
            For Col = 1 To ColCountFree
vbwProfiler.vbwExecuteLine 742
                If Cols(Col).Items < RowCount Then  'This Col is full
vbwProfiler.vbwExecuteLine 743
                    If Cols(Col).Group = SignalAttributes(Idx).Group Then
vbwProfiler.vbwExecuteLine 744
                        For Row = Cols(Col).Items + 1 To RowCount
vbwProfiler.vbwExecuteLine 745
                            .Col = Col
vbwProfiler.vbwExecuteLine 746
                            .Row = Row
vbwProfiler.vbwExecuteLine 747
                            Exit For
vbwProfiler.vbwExecuteLine 748
                        Next Row
                    End If
vbwProfiler.vbwExecuteLine 749 'B
vbwProfiler.vbwExecuteLine 750
                If .Col > 0 Then
vbwProfiler.vbwExecuteLine 751
                     Exit For
                End If
vbwProfiler.vbwExecuteLine 752 'B
                End If
vbwProfiler.vbwExecuteLine 753 'B
vbwProfiler.vbwExecuteLine 754
            If .Col > 0 Then
vbwProfiler.vbwExecuteLine 755
                 Exit For
            End If
vbwProfiler.vbwExecuteLine 756 'B
vbwProfiler.vbwExecuteLine 757
            Next Col
        End If
vbwProfiler.vbwExecuteLine 758 'B

'On initial load place in any free slot
vbwProfiler.vbwExecuteLine 759
        If .Col = 0 Then
vbwProfiler.vbwExecuteLine 760
            For Row = 1 To RowCount
vbwProfiler.vbwExecuteLine 761
                For Col = 1 To ColCount
vbwProfiler.vbwExecuteLine 762
                    If Cols(Col).Items < RowCount Then
vbwProfiler.vbwExecuteLine 763
                        .Col = Col
vbwProfiler.vbwExecuteLine 764
                        If Row < RowCount Then
vbwProfiler.vbwExecuteLine 765
                            .Row = Cols(.Col).Items + 1
vbwProfiler.vbwExecuteLine 766
                            Exit For
                        Else
vbwProfiler.vbwExecuteLine 767 'B
vbwProfiler.vbwExecuteLine 768
MsgBox "No free Rows", vbCritical, "NextFreeGroupFlagPos"
                        End If
vbwProfiler.vbwExecuteLine 769 'B
                    End If
vbwProfiler.vbwExecuteLine 770 'B
vbwProfiler.vbwExecuteLine 771
                If .Col > 0 Then
vbwProfiler.vbwExecuteLine 772
                     Exit For
                End If
vbwProfiler.vbwExecuteLine 773 'B
vbwProfiler.vbwExecuteLine 774
                Next Col
vbwProfiler.vbwExecuteLine 775
            If .Col > 0 Then
vbwProfiler.vbwExecuteLine 776
                 Exit For
            End If
vbwProfiler.vbwExecuteLine 777 'B
vbwProfiler.vbwExecuteLine 778
            Next Row
        End If
vbwProfiler.vbwExecuteLine 779 'B

vbwProfiler.vbwExecuteLine 780
        If .Col = 0 Then
vbwProfiler.vbwExecuteLine 781
MsgBox "No free Cols", vbCritical, "NextFreeGroupFlagPos"
vbwProfiler.vbwProcOut 42
vbwProfiler.vbwExecuteLine 782
            Exit Function
        End If
vbwProfiler.vbwExecuteLine 783 'B

vbwProfiler.vbwExecuteLine 784
        If .Row = 0 Then
vbwProfiler.vbwExecuteLine 785
            If .FixedRow > 0 Then
vbwProfiler.vbwExecuteLine 786
                .Row = .FixedRow
            Else
vbwProfiler.vbwExecuteLine 787 'B
vbwProfiler.vbwExecuteLine 788
                If Row < RowCount Then
vbwProfiler.vbwExecuteLine 789
                    .Row = Cols(.Col).Items + 1
                Else
vbwProfiler.vbwExecuteLine 790 'B
vbwProfiler.vbwExecuteLine 791
MsgBox "No free Rows", vbCritical, "NextFreeGroupFlagPos"
                End If
vbwProfiler.vbwExecuteLine 792 'B
            End If
vbwProfiler.vbwExecuteLine 793 'B
        End If
vbwProfiler.vbwExecuteLine 794 'B

'Check position is actually free
vbwProfiler.vbwExecuteLine 795
        Pos = RC(.Row, .Col)
vbwProfiler.vbwExecuteLine 796
        If Flags(Pos).Picture = 0 Then
'Debug check (before .Pos is Set)
vbwProfiler.vbwExecuteLine 797
            Call DebugFlagsCheck
vbwProfiler.vbwExecuteLine 798
            .Pos = Pos
        Else
vbwProfiler.vbwExecuteLine 799 'B
'This will happen when SplashScreen is loaded (multiple flags in fixed Positions)
vbwProfiler.vbwExecuteLine 800
            If Loading = False Then
vbwProfiler.vbwExecuteLine 801
MsgBox "Signal(" & Idx & ") " & SignalAttributes(Idx).Name & vbCrLf & "Flags(" & Pos & ") not empty", vbCritical, "NextFreeGroupFlagPos"
            End If
vbwProfiler.vbwExecuteLine 802 'B
vbwProfiler.vbwExecuteLine 803
            .Col = 0
vbwProfiler.vbwExecuteLine 804
            .Row = 0
        End If
vbwProfiler.vbwExecuteLine 805 'B

vbwProfiler.vbwExecuteLine 806
    End With
'Debug.Print "NextPos=" & NextFreeGroupFlagPos & " (" & Row & "," & Col & ")"
vbwProfiler.vbwProcOut 42
vbwProfiler.vbwExecuteLine 807
End Function

Private Function DequeCmd(Optional Group As String) As Integer
vbwProfiler.vbwProcIn 43
Dim i As Long
vbwProfiler.vbwExecuteLine 808
    For i = 0 To UBound(CmdQ)
vbwProfiler.vbwExecuteLine 809
        If CmdQ(i) <> 0 Then
vbwProfiler.vbwExecuteLine 810
            If Group = "" Or SignalAttributes(CmdQ(i)).Group = Group Then
vbwProfiler.vbwExecuteLine 811
                If DequeCmd = 0 Then
vbwProfiler.vbwExecuteLine 812
                    DequeCmd = CmdQ(i)
vbwProfiler.vbwExecuteLine 813
Debug.Print "Deque " & SignalAttributes(CmdQ(i)).Name & " (" & Group & ")"
vbwProfiler.vbwExecuteLine 814
                    Commands(CmdQ(i)).BackColor = cbDefault
#If False Then
'When a recall and the command is cancelled, put focus & colour on the other recall command
'Must be done here because only deque is called not Lower Flag
                    Select Case SignalAttributes(CmdQ(i)).Name
                    Case Is = "Recall"
'                        Commands(CommandFromCaption("General Recall")).BackColor = vbGreen
'                        Commands(CommandFromCaption("General Recall")).SetFocus
                    Case Is = "General Recall"
'                        Commands(CommandFromCaption("Recall")).BackColor = vbGreen
'                        Commands(CommandFromCaption("Recall")).SetFocus
                    End Select
#End If
vbwProfiler.vbwExecuteLine 815
                    CmdQ(i) = 0
                End If
vbwProfiler.vbwExecuteLine 816 'B
            End If
vbwProfiler.vbwExecuteLine 817 'B
        End If
vbwProfiler.vbwExecuteLine 818 'B
'Shift remaining commands up the queue
vbwProfiler.vbwExecuteLine 819
        If DequeCmd <> 0 Then
vbwProfiler.vbwExecuteLine 820
            If i = UBound(CmdQ) Then
vbwProfiler.vbwExecuteLine 821
                CmdQ(i) = 0
            Else
vbwProfiler.vbwExecuteLine 822 'B
vbwProfiler.vbwExecuteLine 823
                CmdQ(i) = CmdQ(i + 1)
            End If
vbwProfiler.vbwExecuteLine 824 'B
        End If
vbwProfiler.vbwExecuteLine 825 'B
vbwProfiler.vbwExecuteLine 826
    Next i

'Stop
vbwProfiler.vbwProcOut 43
vbwProfiler.vbwExecuteLine 827
End Function

Private Function QueueCmd(Idx As Long)
vbwProfiler.vbwProcIn 44
Dim i As Long

vbwProfiler.vbwExecuteLine 828
    For i = 0 To UBound(CmdQ)
vbwProfiler.vbwExecuteLine 829
        If CmdQ(i) = 0 Then
vbwProfiler.vbwExecuteLine 830
            CmdQ(i) = Idx
vbwProfiler.vbwExecuteLine 831
            Commands(Idx).BackColor = vbCyan
'Debug.Print "Queue " & SignalAttributes(CmdQ(i)).Name
vbwProfiler.vbwProcOut 44
vbwProfiler.vbwExecuteLine 832
            Exit Function
        Else
vbwProfiler.vbwExecuteLine 833 'B
'Only q the same command once (must not queue Recall more than once)
vbwProfiler.vbwExecuteLine 834
            If CmdQ(i) = Idx Then
vbwProfiler.vbwProcOut 44
vbwProfiler.vbwExecuteLine 835
                 Exit Function
            End If
vbwProfiler.vbwExecuteLine 836 'B
        End If
vbwProfiler.vbwExecuteLine 837 'B
vbwProfiler.vbwExecuteLine 838
    Next i
'MsgBox "Command Queue is full (" & UBound(CmdQ) & ") maximum"
vbwProfiler.vbwProcOut 44
vbwProfiler.vbwExecuteLine 839
End Function

Private Function DisplayStartTimes()
vbwProfiler.vbwProcIn 45
Dim Csidx As Long
Dim FirstStartSecs As Long
Dim kb As String


'Total Secs at first start time
vbwProfiler.vbwExecuteLine 840
    FirstStartSecs = DateDiff("s", Date, FirstStartTime)
'    FirstStartSecs = FirstStartSecs - EventTime
vbwProfiler.vbwExecuteLine 841
    For Csidx = 1 To UBound(Classes)
vbwProfiler.vbwExecuteLine 842
        With mshFinish
vbwProfiler.vbwExecuteLine 843
            If Csidx > .Rows - .FixedRows Then
vbwProfiler.vbwExecuteLine 844
                .Rows = Csidx + .FixedRows
            End If
vbwProfiler.vbwExecuteLine 845 'B
vbwProfiler.vbwExecuteLine 846
            .TextMatrix(Csidx, 0) = "C" & Csidx
vbwProfiler.vbwExecuteLine 847
            .TextMatrix(Csidx, 1) = aSecToElapsed(FirstStartSecs + Classes(Csidx).Start + Classes(Csidx).Offset)
vbwProfiler.vbwExecuteLine 848
        End With
vbwProfiler.vbwExecuteLine 849
    Next Csidx

vbwProfiler.vbwProcOut 45
vbwProfiler.vbwExecuteLine 850
End Function
Private Function StartTime(ByVal Class As Long)
vbwProfiler.vbwProcIn 46
vbwProfiler.vbwExecuteLine 851
    With mshFinish
'not the first (blank) row
vbwProfiler.vbwExecuteLine 852
        If Class > .Rows - .FixedRows Then
vbwProfiler.vbwExecuteLine 853
            .Rows = Class + .FixedRows
        End If
vbwProfiler.vbwExecuteLine 854 'B
vbwProfiler.vbwExecuteLine 855
        .TextMatrix(Class, 0) = "C" & Class
vbwProfiler.vbwExecuteLine 856
        .TextMatrix(Class, 1) = lblCurrTime.Caption
'Scroll to bottom
vbwProfiler.vbwExecuteLine 857
        .TopRow = .Rows - 1
vbwProfiler.vbwExecuteLine 858
    End With
vbwProfiler.vbwProcOut 46
vbwProfiler.vbwExecuteLine 859
End Function

'The finish time must be taken immediately
Private Function FinishTime()
vbwProfiler.vbwProcIn 47

vbwProfiler.vbwExecuteLine 860
    With mshFinish
'not the first (blank) row
vbwProfiler.vbwExecuteLine 861
        FinishCount = FinishCount + 1
vbwProfiler.vbwExecuteLine 862
        If .TextMatrix(.Rows - 1, 0) <> "" Then
vbwProfiler.vbwExecuteLine 863
            .Rows = .Rows + 1
        End If
vbwProfiler.vbwExecuteLine 864 'B
vbwProfiler.vbwExecuteLine 865
        .TextMatrix(.Rows - 1, 0) = FinishCount
vbwProfiler.vbwExecuteLine 866
        .TextMatrix(.Rows - 1, 1) = lblCurrTime.Caption
'Scroll to bottom
vbwProfiler.vbwExecuteLine 867
        .TopRow = .Rows - 1
vbwProfiler.vbwExecuteLine 868
    End With
'Check is there is a linked signal still visible
vbwProfiler.vbwExecuteLine 869
    Call FinishSignalRequest
vbwProfiler.vbwProcOut 47
vbwProfiler.vbwExecuteLine 870
End Function

'A Finish signal must be made for each finisher, so they must
'be queued, if the previous signal has not yet finished
Private Function FinishSignalRequest()
vbwProfiler.vbwProcIn 48
Dim Idx As Long

vbwProfiler.vbwExecuteLine 871
    Idx = SignalFromName("Finish")
vbwProfiler.vbwExecuteLine 872
    If LinkedSignalVisible(Idx) = False Then
'Make the linked signals immediately
vbwProfiler.vbwExecuteLine 873
        FinishSignalCount = FinishSignalCount + 1
vbwProfiler.vbwExecuteLine 874
        Call LinkRequest(Idx)
    End If
vbwProfiler.vbwExecuteLine 875 'B

vbwProfiler.vbwExecuteLine 876
    If FinishSignalCount >= FinishCount Then
'all outstanding signals have been made
vbwProfiler.vbwExecuteLine 877
        FinishTimer.Enabled = False
    Else
vbwProfiler.vbwExecuteLine 878 'B
'Make the signal later - try again in 1 sec
vbwProfiler.vbwExecuteLine 879
        FinishTimer.Enabled = True
    End If
vbwProfiler.vbwExecuteLine 880 'B

vbwProfiler.vbwProcOut 48
vbwProfiler.vbwExecuteLine 881
End Function

'Check if there is a finish signal in progress
Private Function LinkedSignalVisible(ByVal Idx As Long) As Boolean
vbwProfiler.vbwProcIn 49
Dim Lidx As Long
Dim MyLink As defLink

vbwProfiler.vbwExecuteLine 882
    With SignalAttributes(Idx)
vbwProfiler.vbwExecuteLine 883
        If IsLinksInitialised(.Links) Then
vbwProfiler.vbwExecuteLine 884
            For Lidx = 0 To UBound(.Links)
vbwProfiler.vbwExecuteLine 885
                MyLink = .Links(Lidx)
vbwProfiler.vbwExecuteLine 886
                If MyLink.Flag > 0 Then
vbwProfiler.vbwExecuteLine 887
                    If SignalAttributes(MyLink.Flag).Flag.Pos > 0 Then
vbwProfiler.vbwExecuteLine 888
                        LinkedSignalVisible = True
vbwProfiler.vbwProcOut 49
vbwProfiler.vbwExecuteLine 889
                        Exit Function
                    End If
vbwProfiler.vbwExecuteLine 890 'B
                End If
vbwProfiler.vbwExecuteLine 891 'B
vbwProfiler.vbwExecuteLine 892
            Next Lidx
        End If
vbwProfiler.vbwExecuteLine 893 'B
vbwProfiler.vbwExecuteLine 894
    End With
vbwProfiler.vbwProcOut 49
vbwProfiler.vbwExecuteLine 895
End Function

'Keeps running until no outstanding finish signals to make
Private Sub FinishTimer_Timer()
vbwProfiler.vbwProcIn 50
vbwProfiler.vbwExecuteLine 896
    Call FinishSignalRequest
vbwProfiler.vbwProcOut 50
vbwProfiler.vbwExecuteLine 897
End Sub

Private Function FlagVisibility(ByVal Idx As Long, Visible As Boolean)
vbwProfiler.vbwProcIn 51
Dim Pos As Long
Dim Cidx As Long
vbwProfiler.vbwExecuteLine 898
    Pos = SignalAttributes(Idx).Flag.Pos
'See if visiblility has changed (To generate Controller event)
vbwProfiler.vbwExecuteLine 899
    If Pos > 0 Then
vbwProfiler.vbwExecuteLine 900
        If Flags(Pos).Visible <> Visible Then
vbwProfiler.vbwExecuteLine 901
            Flags(Pos).Visible = Visible
vbwProfiler.vbwExecuteLine 902
            Cidx = SignalAttributes(Idx).Controller
vbwProfiler.vbwExecuteLine 903
            If Cidx <> -1 Then
vbwProfiler.vbwExecuteLine 904
                With Controllers(Cidx)
vbwProfiler.vbwExecuteLine 905
                    If Visible Then
'Debug.Print .Connection & "(" & Cidx & ")" & .On
vbwProfiler.vbwExecuteLine 906
                        If .Sound <> "" Then
vbwProfiler.vbwExecuteLine 907
                             Call PlayWav
                        End If
vbwProfiler.vbwExecuteLine 908 'B
vbwProfiler.vbwExecuteLine 909
                        If .On <> "" Then
vbwProfiler.vbwExecuteLine 910
                            Call frmDaventech.OpenAndSend(.On)
vbwProfiler.vbwExecuteLine 911
                            .State = True
                        End If
vbwProfiler.vbwExecuteLine 912 'B
                    Else
vbwProfiler.vbwExecuteLine 913 'B
'Debug.Print .Connection & "(" & Cidx & ")" & .Off
vbwProfiler.vbwExecuteLine 914
                        If .Sound <> "" Then
vbwProfiler.vbwExecuteLine 915
                             Call PauseWav
                        End If
vbwProfiler.vbwExecuteLine 916 'B
vbwProfiler.vbwExecuteLine 917
                        If .Off <> "" Then
vbwProfiler.vbwExecuteLine 918
                            Call frmDaventech.OpenAndSend(.Off)
vbwProfiler.vbwExecuteLine 919
                            .State = False
                        End If
vbwProfiler.vbwExecuteLine 920 'B
                    End If
vbwProfiler.vbwExecuteLine 921 'B
vbwProfiler.vbwExecuteLine 922
                End With
            End If
vbwProfiler.vbwExecuteLine 923 'B
        End If
vbwProfiler.vbwExecuteLine 924 'B
    Else
vbwProfiler.vbwExecuteLine 925 'B
vbwProfiler.vbwExecuteLine 926
        MsgBox "Flag " & SignalAttributes(Idx).Name & " not Raised", vbCritical, "FlagVisibility"
    End If
vbwProfiler.vbwExecuteLine 927 'B
vbwProfiler.vbwProcOut 51
vbwProfiler.vbwExecuteLine 928
End Function

Private Function ResetCols()
vbwProfiler.vbwProcIn 52
Dim Idx As Long
Dim Col As Long
vbwProfiler.vbwExecuteLine 929
    ReDim Cols(ColCount)
vbwProfiler.vbwExecuteLine 930
    For Idx = 1 To UBound(SignalAttributes)
vbwProfiler.vbwExecuteLine 931
        With SignalAttributes(Idx)
vbwProfiler.vbwExecuteLine 932
            If .Flag.Col > 0 And .Flag.Row = 1 Then
vbwProfiler.vbwExecuteLine 933
                Cols(.Flag.Col).Group = .Group
            End If
vbwProfiler.vbwExecuteLine 934 'B
vbwProfiler.vbwExecuteLine 935
            If .Flag.FixedCol > 0 Then
vbwProfiler.vbwExecuteLine 936
                Cols(.Flag.FixedCol).Group = .Group
            End If
vbwProfiler.vbwExecuteLine 937 'B
vbwProfiler.vbwExecuteLine 938
            If SignalAttributes(Idx).Flag.Col > 0 Then
vbwProfiler.vbwExecuteLine 939
                Cols(.Flag.Col).Items = Cols(.Flag.Col).Items + 1
            End If
vbwProfiler.vbwExecuteLine 940 'B
vbwProfiler.vbwExecuteLine 941
        End With
vbwProfiler.vbwExecuteLine 942
    Next Idx
vbwProfiler.vbwProcOut 52
vbwProfiler.vbwExecuteLine 943
End Function

'Used to Check if a Class Flag is up when Recall is asked for
'If 2 Class flags are up it will select the lowest class (Idx is in class order)
Private Function GroupIdx(ByVal Group As String) As Long
vbwProfiler.vbwProcIn 53
Dim Idx As Long
vbwProfiler.vbwExecuteLine 944
    For Idx = 1 To UBound(SignalAttributes)
vbwProfiler.vbwExecuteLine 945
        With SignalAttributes(Idx)
vbwProfiler.vbwExecuteLine 946
            If .Group = Group And .Flag.Pos > 0 Then
vbwProfiler.vbwExecuteLine 947
                 GroupIdx = Idx
vbwProfiler.vbwExecuteLine 948
                Exit For
            End If
vbwProfiler.vbwExecuteLine 949 'B
vbwProfiler.vbwExecuteLine 950
        End With
vbwProfiler.vbwExecuteLine 951
    Next Idx
vbwProfiler.vbwProcOut 53
vbwProfiler.vbwExecuteLine 952
End Function

'Return the Command Button IDX, as we should find it within the first 6 buttons (Fixed)
Public Function CommandFromCaption(ByVal CbName As String) As Integer
vbwProfiler.vbwProcIn 54
Dim Index As Integer
vbwProfiler.vbwExecuteLine 953
    For Index = 0 To Commands.Count
vbwProfiler.vbwExecuteLine 954
        If Commands(Index).Caption = CbName Then
vbwProfiler.vbwExecuteLine 955
            CommandFromCaption = Index
vbwProfiler.vbwProcOut 54
vbwProfiler.vbwExecuteLine 956
            Exit Function
        End If
vbwProfiler.vbwExecuteLine 957 'B
vbwProfiler.vbwExecuteLine 958
    Next Index
'Stop
vbwProfiler.vbwProcOut 54
vbwProfiler.vbwExecuteLine 959
End Function

'The EventTime has the PausedTime taken off
'Called once a second
Public Function DoTimerEvents() '(ByVal EventTime As Long)
vbwProfiler.vbwProcIn 55
Dim Eidx As Long
Dim Sidx As Long
Dim Bidx As Long
Dim Csidx As Long
Dim Pause As Boolean
Dim NextClass As Long   'to Start
Dim LastClass As Long   'to Start

'Timer is enabled to show the time while splash screen is displayed
vbwProfiler.vbwExecuteLine 960
    If ClearFlagsTimer.Enabled = True Then
vbwProfiler.vbwProcOut 55
vbwProfiler.vbwExecuteLine 961
        Exit Function
    End If
vbwProfiler.vbwExecuteLine 962 'B

'This must be called immediately prior to the first event
'To ensure focus is put on the Horn button
vbwProfiler.vbwExecuteLine 963
    If EventTime = Evts(0).ElapsedTime + Classes(Evts(0).Class).Offset Then
vbwProfiler.vbwExecuteLine 964
        Call DefaultsFirstEvent
'Dont allow the User to postpone once the first class has started
'May need changing to allow subsequent classes to be postponed
'Now allow changes after start        txtPostpone.Enabled = False
    End If
vbwProfiler.vbwExecuteLine 965 'B


vbwProfiler.vbwExecuteLine 966
    lblElapsedTime = aSecToElapsed(EventTime)


vbwProfiler.vbwExecuteLine 967
    ClassStart = False  'To suppress the sound signal if a Class Warning is
                        'Raised after a class start at the same elapsed second

'If this event is a class start, these will return the Class before this class start
'Debug.Print EventTime & "-" & LastClassStart(EventTime)
'Debug.Print EventTime & "-" & NextClassStart(EventTime)

'This must be done before each event as the Postpone may have been raised
'since the last event
vbwProfiler.vbwExecuteLine 968
    If SignalAttributes(SignalFromName("Postpone")).Flag.Pos > 0 Then
'If counting down reduce by 1
vbwProfiler.vbwExecuteLine 969
        If PostponeCountDown > 0 Then
vbwProfiler.vbwExecuteLine 970
            PostponeCountDown = PostponeCountDown - 1
vbwProfiler.vbwExecuteLine 971
            If PostponeCountDown = 0 Then
'This is the last CountDown, so the Postpone has timed out
vbwProfiler.vbwExecuteLine 972
                Call LowerFlag(SignalFromName("Postpone"))
vbwProfiler.vbwExecuteLine 973
                lblCountDown = aSecToElapsed(PostponeCountDown)
            End If
vbwProfiler.vbwExecuteLine 974 'B
        End If
vbwProfiler.vbwExecuteLine 975 'B

'Get the Class we are postponing
vbwProfiler.vbwExecuteLine 976
        NextClass = NextClassStart(EventTime)
vbwProfiler.vbwExecuteLine 977
        If NextClass = 0 Then
'This will happen if the StartTiume has not been set by the user when Postpone is
'Clicked as the EventTime calculated by RaceTimer_Timer will be the elapsed time since
'Midnignt. By calling ClassRestart, the FirstStartTime will be computed from
'the current Time & then when we call NextClassStart will return 1 because the Offsets
'will have been added into Evts as the current time + the back off to the first event
'for the class
vbwProfiler.vbwExecuteLine 978
            Call ClassRestart(1)
vbwProfiler.vbwExecuteLine 979
            NextClass = NextClassStart(EventTime)
'        Stop
        End If
vbwProfiler.vbwExecuteLine 980 'B

'If the first event after the Postpone flag has been raised, set Postpone on the class
vbwProfiler.vbwExecuteLine 981
        If Classes(NextClass).Postpone = False Then
vbwProfiler.vbwExecuteLine 982
            If IsNumeric(txtPostpone) Then
vbwProfiler.vbwExecuteLine 983
                PostponeCountDown = txtPostpone * 60 / Multiplier
            End If
vbwProfiler.vbwExecuteLine 984 'B
vbwProfiler.vbwExecuteLine 985
            Call ClassClear     'Clear any class flags that are raised (no sound)
vbwProfiler.vbwExecuteLine 986
            Call ClassRestart(NextClass)    'Reset the StartTime by adding an offset

'Set the postpone flag on the first class that is being postponed
vbwProfiler.vbwExecuteLine 987
            Classes(NextClass).Postpone = True
        End If
vbwProfiler.vbwExecuteLine 988 'B
    End If
vbwProfiler.vbwExecuteLine 989 'B
'End of Postpone flag is up

'General Recall
vbwProfiler.vbwExecuteLine 990
    If SignalAttributes(SignalFromName("General Recall")).Flag.Pos > 0 Then
'If the first event after the Postpone flag has been raised, set GeneralRecall on the class
'Check ClassRestart has not already been called
vbwProfiler.vbwExecuteLine 991
        NextClass = NextClassStart(EventTime)
vbwProfiler.vbwExecuteLine 992
        If Classes(NextClass).GeneralRecall = False Then
'Check this is the First Event after Recall has been raised
vbwProfiler.vbwExecuteLine 993
            LastClass = LastClassStart(EventTime)
vbwProfiler.vbwExecuteLine 994
            If Classes(LastClass).GeneralRecall = False Then
'Reset the Class Start Times to the LastClassStarted
vbwProfiler.vbwExecuteLine 995
                LastClass = LastClassStart(EventTime)
vbwProfiler.vbwExecuteLine 996
Debug.Print "GR " & LastClass
vbwProfiler.vbwExecuteLine 997
If SkipClassOnRecall = False Then
vbwProfiler.vbwExecuteLine 998
                Call ClassClear
vbwProfiler.vbwExecuteLine 999
                Call ClassRestart(LastClass)
'NextClass is now LastClass
End If
vbwProfiler.vbwExecuteLine 1000 'B
vbwProfiler.vbwExecuteLine 1001
                Classes(LastClass).GeneralRecall = True
            End If
vbwProfiler.vbwExecuteLine 1002 'B
        End If
vbwProfiler.vbwExecuteLine 1003 'B
    End If
vbwProfiler.vbwExecuteLine 1004 'B

'Call the events for each class in turn
vbwProfiler.vbwExecuteLine 1005
    For Csidx = 0 To UBound(Classes)

vbwProfiler.vbwExecuteLine 1006
        With Classes(Csidx)

'If Postpone Flag is up Postpone any class not started
vbwProfiler.vbwExecuteLine 1007
            If .Postpone = True Then
vbwProfiler.vbwExecuteLine 1008
                If SignalAttributes(SignalFromName("Postpone")).Flag.Pos = 0 Then
'The Postpone Flag has been lowered manually (because .Postpone is still true)
vbwProfiler.vbwExecuteLine 1009
                    .Postpone = False
vbwProfiler.vbwExecuteLine 1010
                    NextClass = NextClassStart(EventTime)
vbwProfiler.vbwExecuteLine 1011
                    Call ClassRestart(NextClass)
                Else
vbwProfiler.vbwExecuteLine 1012 'B
'The Postpone flag is still up
'The Initial Countdown Time is included in the Offset so we do not pause if its running
vbwProfiler.vbwExecuteLine 1013
                    If PostponeCountDown = 0 Then
vbwProfiler.vbwExecuteLine 1014
                        Pause = True
                    End If
vbwProfiler.vbwExecuteLine 1015 'B
                End If
vbwProfiler.vbwExecuteLine 1016 'B
            End If
vbwProfiler.vbwExecuteLine 1017 'B

'If General Recall Flag is up Pause any class not started
vbwProfiler.vbwExecuteLine 1018
            If .GeneralRecall = True And SkipClassOnRecall = False Then
vbwProfiler.vbwExecuteLine 1019
                If SignalAttributes(SignalFromName("General Recall")).Flag.Pos = 0 Then
'The General Recall Flag has been lowered manually (because .GeneralRecall is still true)
vbwProfiler.vbwExecuteLine 1020
                    .GeneralRecall = False
vbwProfiler.vbwExecuteLine 1021
                    NextClass = NextClassStart(EventTime)
vbwProfiler.vbwExecuteLine 1022
                    Call ClassRestart(NextClass)
                Else
vbwProfiler.vbwExecuteLine 1023 'B
'The General Recall flag is still up
'The Initial Countdown Time is included in the Offset so we do not pause if its running
'The postpone could have been raised while the General Recall Flag is still up
vbwProfiler.vbwExecuteLine 1024
                    If PostponeCountDown = 0 Then
vbwProfiler.vbwExecuteLine 1025
                        Pause = True
                    End If
vbwProfiler.vbwExecuteLine 1026 'B
                End If
vbwProfiler.vbwExecuteLine 1027 'B
            End If
vbwProfiler.vbwExecuteLine 1028 'B

'This class and all subsequent Classes must be paused
vbwProfiler.vbwExecuteLine 1029
            If Pause = True Then
vbwProfiler.vbwExecuteLine 1030
                .Offset = .Offset + 1
            End If
vbwProfiler.vbwExecuteLine 1031 'B

vbwProfiler.vbwExecuteLine 1032
        End With

vbwProfiler.vbwExecuteLine 1033
        For Eidx = 0 To UBound(Evts)
vbwProfiler.vbwExecuteLine 1034
            If Evts(Eidx).Class = Csidx Then
vbwProfiler.vbwExecuteLine 1035
                If EventTime = Evts(Eidx).ElapsedTime + Classes(Evts(Eidx).Class).Offset Then
vbwProfiler.vbwExecuteLine 1036
                    Call ProcessEvent(Eidx)
                End If  'This Event
vbwProfiler.vbwExecuteLine 1037 'B
            End If
vbwProfiler.vbwExecuteLine 1038 'B
vbwProfiler.vbwExecuteLine 1039
        Next Eidx
vbwProfiler.vbwExecuteLine 1040
    Next Csidx
vbwProfiler.vbwExecuteLine 1041
    Call DisplayStartTimes

'Display Countdown timer rather than Postpone Text Box
vbwProfiler.vbwExecuteLine 1042
    If PostponeCountDown > 0 Then
vbwProfiler.vbwExecuteLine 1043
        lblCountDown = aSecToElapsed(PostponeCountDown)
vbwProfiler.vbwExecuteLine 1044
        txtPostpone.Visible = False
vbwProfiler.vbwExecuteLine 1045
        lblCountDown.Visible = True
vbwProfiler.vbwExecuteLine 1046
        fraPostpone.Caption = "Time to Go"
    Else
vbwProfiler.vbwExecuteLine 1047 'B
vbwProfiler.vbwExecuteLine 1048
        lblCountDown.Visible = False
vbwProfiler.vbwExecuteLine 1049
        txtPostpone.Visible = True
vbwProfiler.vbwExecuteLine 1050
        fraPostpone.Caption = "Minutes"
    End If
vbwProfiler.vbwExecuteLine 1051 'B
vbwProfiler.vbwProcOut 55
vbwProfiler.vbwExecuteLine 1052
End Function

'This processes 1 event
Private Function ProcessEvent(ByVal Eidx As Long)
vbwProfiler.vbwProcIn 56
Dim Sidx As Long
Dim Bidx As Long

'When the first event is carried out disable the start time
'            If Eidx = 0 Then
'                txtFirstStartTime.Enabled = False
'Dont allow the User to postpone once the first class has started
'May need changing to allow subsequent classes to be postponed
'                txtPostpone.Enabled = False
'            End If
vbwProfiler.vbwExecuteLine 1053
            If Left$(Evts(Eidx).Message, 1) <> "~" Then
vbwProfiler.vbwExecuteLine 1054
                StatusBar1.Panels(1).Text = Evts(Eidx).Message
            End If
vbwProfiler.vbwExecuteLine 1055 'B
vbwProfiler.vbwExecuteLine 1056
            If IsSignalsInitialised(Evts(Eidx).Signals) Then
vbwProfiler.vbwExecuteLine 1057
                For Sidx = 0 To UBound(Evts(Eidx).Signals)
vbwProfiler.vbwExecuteLine 1058
                    With Evts(Eidx).Signals(Sidx)
'Silent on SignalAttributes is only used by LinkRquest and is set temporarily
'for this call only
vbwProfiler.vbwExecuteLine 1059
                        If .Silent = "True" Then
vbwProfiler.vbwExecuteLine 1060
                             SignalAttributes(.Signal).Silent = True
                        End If
vbwProfiler.vbwExecuteLine 1061 'B
'Dont make sound signal if another class flag is raised at the same
'time as a Class is started
vbwProfiler.vbwExecuteLine 1062
                            If SignalAttributes(.Signal).Group = "Class" And ClassStart = True Then
vbwProfiler.vbwExecuteLine 1063
                                SignalAttributes(.Signal).Silent = True
                            End If
vbwProfiler.vbwExecuteLine 1064 'B

'Must be explictly asked for
vbwProfiler.vbwExecuteLine 1065
                        If .Raise = True Then
vbwProfiler.vbwExecuteLine 1066
                            If SignalAttributes(.Signal).Flag.Pos = 0 Then
vbwProfiler.vbwExecuteLine 1067
                                Call frmMain.RaiseRequest(.Signal)
                            End If
vbwProfiler.vbwExecuteLine 1068 'B
                        End If
vbwProfiler.vbwExecuteLine 1069 'B
'Must be explictly asked for
vbwProfiler.vbwExecuteLine 1070
                        If .Raise = False Then
vbwProfiler.vbwExecuteLine 1071
                            If SignalAttributes(.Signal).Flag.Pos > 0 Then
vbwProfiler.vbwExecuteLine 1072
                                Call frmMain.LowerRequest(.Signal)
vbwProfiler.vbwExecuteLine 1073
                                If SignalAttributes(.Signal).Group = "Class" Then
vbwProfiler.vbwExecuteLine 1074
                                    ClassStart = True
'Add on MshFinish
vbwProfiler.vbwExecuteLine 1075
                                    Call StartTime(SignalAttributes(.Signal).Class)
                                End If
vbwProfiler.vbwExecuteLine 1076 'B
                            End If
vbwProfiler.vbwExecuteLine 1077 'B

'This is not required as it is now actioned by LowerRequest
'If a recall we must Remove this RecallIdx explicitly, this prevents the recall
'being actioned without a flag
'                            If SignalAttributes(.Signal).Group = "Recall" Then
'                                If SignalAttributes(.Signal).Flag.Pos = 0 Then
'Stop
'                                    Call RecallClearSignal
'                                End If
'                            End If
                        End If
vbwProfiler.vbwExecuteLine 1078 'B

'Reset (may have been set to silent temporarily for this event above)
vbwProfiler.vbwExecuteLine 1079
                        SignalAttributes(.Signal).Silent = False
vbwProfiler.vbwExecuteLine 1080
                    End With
vbwProfiler.vbwExecuteLine 1081
                Next Sidx
            End If
vbwProfiler.vbwExecuteLine 1082 'B
vbwProfiler.vbwExecuteLine 1083
            If IsButtonsInitialised(Evts(Eidx).Buttons) Then
vbwProfiler.vbwExecuteLine 1084
                For Bidx = 0 To UBound(Evts(Eidx).Buttons)
vbwProfiler.vbwExecuteLine 1085
                    With Evts(Eidx).Buttons(Bidx)
'Stop
'The Command button properties can be set immediately if any
vbwProfiler.vbwExecuteLine 1086
                        If .Enabled <> "" Then
'If the flag is up, you may not disable the command button because the user
'must be able to manually drop the flag
'This is important with recalls and probably postpone
vbwProfiler.vbwExecuteLine 1087
                            If SignalAttributes(.Button).Flag.Pos = 0 Or Commands(.Button).Enabled = False Then
vbwProfiler.vbwExecuteLine 1088
                                Commands(.Button).Enabled = AtoBool(.Enabled)
                            End If
vbwProfiler.vbwExecuteLine 1089 'B

'If we disable any command we must clear the colour
vbwProfiler.vbwExecuteLine 1090
                            If Commands(.Button).Enabled = False Then
vbwProfiler.vbwExecuteLine 1091
                                Commands(.Button).BackColor = cbDefault
                            End If
vbwProfiler.vbwExecuteLine 1092 'B

vbwProfiler.vbwExecuteLine 1093
                            If SignalAttributes(.Button).Group = "Recall" Then
vbwProfiler.vbwExecuteLine 1094
                                If Commands(.Button).Enabled = True Then
'Add the Recall Class Flag when the Recall button is enabled
'Also Sets the NextStart Signal on the RecallClass
'RecallClearSignal is called when the Signal is dropped manually or on timeout
vbwProfiler.vbwExecuteLine 1095
                                    Call RecallSetSignal
                                End If
vbwProfiler.vbwExecuteLine 1096 'B
                            End If
vbwProfiler.vbwExecuteLine 1097 'B

                        End If
vbwProfiler.vbwExecuteLine 1098 'B
vbwProfiler.vbwExecuteLine 1099
                    End With
vbwProfiler.vbwExecuteLine 1100
                Next Bidx
            End If
vbwProfiler.vbwExecuteLine 1101 'B
vbwProfiler.vbwExecuteLine 1102
            If Evts(Eidx).Focus > 0 Then
'Must be enabled & visible to put focus on it
vbwProfiler.vbwExecuteLine 1103
                Commands(Evts(Eidx).Focus).Enabled = True
vbwProfiler.vbwExecuteLine 1104
                Commands(Evts(Eidx).Focus).Visible = True
vbwProfiler.vbwExecuteLine 1105
                Commands(Evts(Eidx).Focus).SetFocus
vbwProfiler.vbwExecuteLine 1106
                Commands(Evts(Eidx).Focus).BackColor = vbGreen
            End If
vbwProfiler.vbwExecuteLine 1107 'B
'This is the Commands(0) button
vbwProfiler.vbwExecuteLine 1108
            If Evts(Eidx).Focus = 0 Then
vbwProfiler.vbwExecuteLine 1109
                Commands(Evts(Eidx).Focus).Enabled = True
vbwProfiler.vbwExecuteLine 1110
                Commands(Evts(Eidx).Focus).Visible = True
vbwProfiler.vbwExecuteLine 1111
                Commands(Evts(Eidx).Focus).SetFocus
vbwProfiler.vbwExecuteLine 1112
                Commands(Evts(Eidx).Focus).Enabled = False
vbwProfiler.vbwExecuteLine 1113
                Commands(Evts(Eidx).Focus).Visible = False
            End If
vbwProfiler.vbwExecuteLine 1114 'B
vbwProfiler.vbwProcOut 56
vbwProfiler.vbwExecuteLine 1115
End Function

'Called when Postpone or GeneralRecall is raised
'Removes all raised ClassFlags (any any below the class flag)
Private Function ClassClear()
vbwProfiler.vbwProcIn 57
Dim Idx As Long

vbwProfiler.vbwExecuteLine 1116
Debug.Print "ClassClear"
vbwProfiler.vbwExecuteLine 1117
    For Idx = 1 To UBound(SignalAttributes)
vbwProfiler.vbwExecuteLine 1118
        With SignalAttributes(Idx)
vbwProfiler.vbwExecuteLine 1119
            If .Group = "Class" And .Flag.Pos > 0 Then
vbwProfiler.vbwExecuteLine 1120
                    .Silent = True
vbwProfiler.vbwExecuteLine 1121
                    Call LowerFlag(Idx)
vbwProfiler.vbwExecuteLine 1122
                    Call LinkRequest(Idx)
vbwProfiler.vbwExecuteLine 1123
                    .Silent = False
            End If
vbwProfiler.vbwExecuteLine 1124 'B
vbwProfiler.vbwExecuteLine 1125
        End With
vbwProfiler.vbwExecuteLine 1126
    Next Idx
vbwProfiler.vbwProcOut 57
vbwProfiler.vbwExecuteLine 1127
End Function

'Called when Postpone Raised or GeneralRecall is Raised
'Class will be 0 if called before the start sequence has started
Private Function ClassRestart(ByVal Class As Long, Optional NextSecond As Long)
'Dim ElapsedTime As Long
vbwProfiler.vbwProcIn 58
Dim Offset As Long
Dim Csidx As Long

'The Offset is that required at the NextEvent when the flag
'Secs we have to start sequence before the start
'NextSecond=0 when called from DoEvents
'NextSecond=1 when called by LowerFlag
vbwProfiler.vbwExecuteLine 1128
        Offset = EventTime - FirstEventTime(Class) + NextSecond

vbwProfiler.vbwExecuteLine 1129
Call DisplayStartTimes
'This is required when postpone is before the firststart
vbwProfiler.vbwExecuteLine 1130
        Offset = Offset + PostponeCountDown
'This will happen if the class can be manually postponed longer than then
'earliest the postponed class can start
vbwProfiler.vbwExecuteLine 1131
        If PostponeCountDown > Offset Then
vbwProfiler.vbwExecuteLine 1132
            Offset = PostponeCountDown
        End If
vbwProfiler.vbwExecuteLine 1133 'B
    'time for each class, including this class
vbwProfiler.vbwExecuteLine 1134
Debug.Print "Restart C" & Class & "(" & Trim$(aSecToElapsed(EventTime)) & ")"
vbwProfiler.vbwExecuteLine 1135
        For Csidx = Class To UBound(Classes)
vbwProfiler.vbwExecuteLine 1136
            Classes(Csidx).Offset = Offset
vbwProfiler.vbwExecuteLine 1137
        Next Csidx
vbwProfiler.vbwExecuteLine 1138
Call DisplayStartTimes
vbwProfiler.vbwProcOut 58
vbwProfiler.vbwExecuteLine 1139
End Function
'Sets the ClassFlagImage on the RecallClassIdx
'Sets the RecallIdx
'Called when the Recall Buttons are enabled
'Also called when Recall is clicked before Start Sequence and there is no recallIdx
'to help testing)
Private Function RecallSetSignal()
vbwProfiler.vbwProcIn 59
Dim i As Long
Dim Idx As Long
vbwProfiler.vbwExecuteLine 1140
    i = SignalFromName("Recall Class")
vbwProfiler.vbwExecuteLine 1141
    If i > 0 Then   'When Loading the Recall Class is not set
vbwProfiler.vbwExecuteLine 1142
        RecallIdx = GetPostponeIdx  'The Class we will recall
'If there is no Class Flag up There will be no Postpone Idx so we
'cannot set a Class to recall
vbwProfiler.vbwExecuteLine 1143
        If RecallIdx > 0 Then
vbwProfiler.vbwExecuteLine 1144
            Set SignalAttributes(i).Image = SignalAttributes(RecallIdx).Image
vbwProfiler.vbwExecuteLine 1145
    Commands(CommandFromCaption("Recall")).Enabled = True
vbwProfiler.vbwExecuteLine 1146
    Commands(CommandFromCaption("General Recall")).Enabled = True
        End If
vbwProfiler.vbwExecuteLine 1147 'B
    End If
vbwProfiler.vbwExecuteLine 1148 'B
vbwProfiler.vbwProcOut 59
vbwProfiler.vbwExecuteLine 1149
End Function

'Cleared when the Recall Flag is Lowered. The Recall Flag is linked to the Recall
'Class Flag
Private Function RecallClearSignal()
vbwProfiler.vbwProcIn 60
Dim Idx As Long

vbwProfiler.vbwExecuteLine 1150
    Idx = SignalFromName("Recall Class")
vbwProfiler.vbwExecuteLine 1151
    Set SignalAttributes(Idx).Image = Nothing
vbwProfiler.vbwExecuteLine 1152
    RecallIdx = 0
vbwProfiler.vbwExecuteLine 1153
    Commands(CommandFromCaption("Recall")).Enabled = False
vbwProfiler.vbwExecuteLine 1154
    Commands(CommandFromCaption("General Recall")).Enabled = False
vbwProfiler.vbwProcOut 60
vbwProfiler.vbwExecuteLine 1155
End Function

'Called when a Recall Flag is actually raised, which can be after the start
'Idx is the Recall Flag we are raising
Private Function RecallChange(ByVal Idx As Long)
vbwProfiler.vbwProcIn 61
Dim OtherIdx As Long
Dim SaveRecallIdx As Long
Dim i As Long

vbwProfiler.vbwExecuteLine 1156
    Call DequeCmd("Recall") 'May be queued and this request is to Deque it
vbwProfiler.vbwExecuteLine 1157
    Select Case SignalAttributes(Idx).Name
'vbwLine 1158:    Case Is = "Recall"
    Case Is = IIf(vbwProfiler.vbwExecuteLine(1158), VBWPROFILER_EMPTY, _
        "Recall")
vbwProfiler.vbwExecuteLine 1159
        OtherIdx = SignalFromName("General Recall")
'vbwLine 1160:    Case Is = "General Recall"
    Case Is = IIf(vbwProfiler.vbwExecuteLine(1160), VBWPROFILER_EMPTY, _
        "General Recall")
vbwProfiler.vbwExecuteLine 1161
        OtherIdx = SignalFromName("Recall")
    Case Else
vbwProfiler.vbwExecuteLine 1162 'B
'This is the RecallClass flag being raised
vbwProfiler.vbwExecuteLine 1163
Stop
vbwProfiler.vbwProcOut 61
vbwProfiler.vbwExecuteLine 1164
        Exit Function
    End Select
vbwProfiler.vbwExecuteLine 1165 'B

'Check if other recall flag is up
vbwProfiler.vbwExecuteLine 1166
    If SignalAttributes(OtherIdx).Flag.Pos > 0 Then
'We need to save the class flag Idx because LowerFlag will remove it
vbwProfiler.vbwExecuteLine 1167
        SaveRecallIdx = RecallIdx
vbwProfiler.vbwExecuteLine 1168
        Call LowerFlag(OtherIdx)
vbwProfiler.vbwExecuteLine 1169
        Call LinkRequest(OtherIdx)  'Drop the Linked signal (Fl White)
vbwProfiler.vbwExecuteLine 1170
        RecallIdx = SaveRecallIdx
vbwProfiler.vbwExecuteLine 1171
        i = SignalFromName("Recall Class")
vbwProfiler.vbwExecuteLine 1172
        Set SignalAttributes(i).Image = SignalAttributes(RecallIdx).Image
vbwProfiler.vbwExecuteLine 1173
        Commands(Idx).Enabled = True
vbwProfiler.vbwExecuteLine 1174
        Commands(OtherIdx).Enabled = True
    End If
vbwProfiler.vbwExecuteLine 1175 'B

'Clear Green off Recall if General Recall is clicked before the start
'as well as when the Recall is changed
vbwProfiler.vbwExecuteLine 1176
    Commands(OtherIdx).BackColor = cbDefault

vbwProfiler.vbwProcOut 61
vbwProfiler.vbwExecuteLine 1177
End Function

'Returns the Lowest Class Flag currently UP. With SYC there can be 2 Class Flags
'UP at the same time, The Lowest will be the next one to start
'Returns 0 if no class flag up
Private Function GetPostponeIdx() As Long
vbwProfiler.vbwProcIn 62
Dim Csidx As Long
Dim Idx As Long

vbwProfiler.vbwExecuteLine 1178
    For Csidx = 1 To UBound(Classes)
vbwProfiler.vbwExecuteLine 1179
        If SignalAttributes(Classes(Csidx).Signal).Flag.Pos > 0 Then
vbwProfiler.vbwExecuteLine 1180
            Idx = Classes(Csidx).Signal
vbwProfiler.vbwExecuteLine 1181
            Exit For
        End If
vbwProfiler.vbwExecuteLine 1182 'B
vbwProfiler.vbwExecuteLine 1183
    Next Csidx
vbwProfiler.vbwExecuteLine 1184
    GetPostponeIdx = Idx
vbwProfiler.vbwProcOut 62
vbwProfiler.vbwExecuteLine 1185
End Function

Private Function NextClassStart(ByVal EventTime As Long) As Long
vbwProfiler.vbwProcIn 63
Dim Eidx As Long
Dim Sidx As Long
Dim Silence As Boolean
Dim i As Long

'If no start time has yet been entered the EventTime will the current elapsed time
'but the events will not have yet been set up
'The NextClassStart will be 0 as the EventTime will be after the Latest Event
'This occurs when the Postpone flag is raised before the starttime is entered.
'DoTimerEvents must then call ClassRestart to set up the Class Offsets and then
'NextClassStart again to get the Class
vbwProfiler.vbwExecuteLine 1186
    For Eidx = 0 To UBound(Evts)
vbwProfiler.vbwExecuteLine 1187
        If Evts(Eidx).ElapsedTime + Classes(Evts(Eidx).Class).Offset >= EventTime Then
vbwProfiler.vbwExecuteLine 1188
            If IsSignalsInitialised(Evts(Eidx).Signals) Then
vbwProfiler.vbwExecuteLine 1189
                For Sidx = 0 To UBound(Evts(Eidx).Signals)
vbwProfiler.vbwExecuteLine 1190
                    If SignalAttributes(Evts(Eidx).Signals(Sidx).Signal).Group = "Class" Then
vbwProfiler.vbwExecuteLine 1191
                        If Evts(Eidx).Signals(Sidx).Raise = "False" Then
vbwProfiler.vbwExecuteLine 1192
                            NextClassStart = Evts(Eidx).Class
vbwProfiler.vbwProcOut 63
vbwProfiler.vbwExecuteLine 1193
                            Exit Function
                        End If
vbwProfiler.vbwExecuteLine 1194 'B
                    End If
vbwProfiler.vbwExecuteLine 1195 'B
vbwProfiler.vbwExecuteLine 1196
                Next Sidx
            End If
vbwProfiler.vbwExecuteLine 1197 'B
        End If
vbwProfiler.vbwExecuteLine 1198 'B
vbwProfiler.vbwExecuteLine 1199
    Next Eidx
vbwProfiler.vbwExecuteLine 1200
Call frmEvents.ListEvents
'Stop
vbwProfiler.vbwProcOut 63
vbwProfiler.vbwExecuteLine 1201
End Function

Private Function LastClassStart(ByVal EventTime As Long) As Long
vbwProfiler.vbwProcIn 64
Dim Eidx As Long
Dim Sidx As Long
Dim Silence As Boolean
Dim i As Long

'If no start time has yet been entered the EventTime will the current elapsed time
'but the events will not have yet been set up
'The NextClassStart will be 0 as the EventTime will be after the Latest Event
'This occurs when the Postpone flag is raised before the starttime is entered.
'DoTimerEvents must then call ClassRestart to set up the Class Offsets and then
'NextClassStart again to get the Class
vbwProfiler.vbwExecuteLine 1202
Call frmEvents.ListEvents
vbwProfiler.vbwExecuteLine 1203
    For Eidx = 0 To UBound(Evts)
'Wants testing to see if it should be >=
vbwProfiler.vbwExecuteLine 1204
        If Evts(Eidx).ElapsedTime + Classes(Evts(Eidx).Class).Offset > EventTime Then
vbwProfiler.vbwExecuteLine 1205
            Exit For
        End If
vbwProfiler.vbwExecuteLine 1206 'B
vbwProfiler.vbwExecuteLine 1207
        If IsSignalsInitialised(Evts(Eidx).Signals) Then
vbwProfiler.vbwExecuteLine 1208
            For Sidx = 0 To UBound(Evts(Eidx).Signals)
vbwProfiler.vbwExecuteLine 1209
                If SignalAttributes(Evts(Eidx).Signals(Sidx).Signal).Group = "Class" Then
vbwProfiler.vbwExecuteLine 1210
                    If Evts(Eidx).Signals(Sidx).Raise = "False" Then
vbwProfiler.vbwExecuteLine 1211
                        LastClassStart = Evts(Eidx).Class
                    End If
vbwProfiler.vbwExecuteLine 1212 'B
                End If
vbwProfiler.vbwExecuteLine 1213 'B
vbwProfiler.vbwExecuteLine 1214
            Next Sidx
        End If
vbwProfiler.vbwExecuteLine 1215 'B
vbwProfiler.vbwExecuteLine 1216
    Next Eidx
vbwProfiler.vbwExecuteLine 1217
Call frmEvents.ListEvents
'Stop
vbwProfiler.vbwProcOut 64
vbwProfiler.vbwExecuteLine 1218
End Function

Private Function FirstEventTime(ByVal Class As Long) As Long
vbwProfiler.vbwProcIn 65
Dim Eidx As Long
Dim Sidx As Long
Dim Silence As Boolean
Dim i As Long

'If no valid date has yet been entered the the next class must be 1
'The EventTime that will have been passed will be the no of secs since midnight
'If class 0 was returned the next DoEvents would think all classes had started
'    If txtFirstStartTime.Enabled = True Then
'Stop
'        Exit Function
'    End If
vbwProfiler.vbwExecuteLine 1219
    For Eidx = 0 To UBound(Evts)
vbwProfiler.vbwExecuteLine 1220
        If Evts(Eidx).Class = Class Then
vbwProfiler.vbwExecuteLine 1221
            FirstEventTime = Evts(Eidx).ElapsedTime ' + Classes(Class).Offset
vbwProfiler.vbwProcOut 65
vbwProfiler.vbwExecuteLine 1222
            Exit Function
        End If
vbwProfiler.vbwExecuteLine 1223 'B
vbwProfiler.vbwExecuteLine 1224
    Next Eidx
vbwProfiler.vbwProcOut 65
vbwProfiler.vbwExecuteLine 1225
End Function

Private Function SetPostpone()
vbwProfiler.vbwProcIn 66

vbwProfiler.vbwProcOut 66
vbwProfiler.vbwExecuteLine 1226
End Function



