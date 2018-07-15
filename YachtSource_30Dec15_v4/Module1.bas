Attribute VB_Name = "modMain"
Option Explicit

Public Const cbDefault = &H8000000F

Public Declare Function ShellExecute _
                            Lib "SHELL32.DLL" _
                            Alias "ShellExecuteA" ( _
                            ByVal hwnd As Long, _
                            ByVal lpOperation As String, _
                            ByVal lpFile As String, _
                            ByVal lpParameters As String, _
                            ByVal lpDirectory As String, _
                            ByVal nShowCmd As Long) _
                            As Long

Public cbOrange As Long
Public Type defLink
    Flag As Long    'The Link Flag that is associated
    Raise As Boolean    'True = RaiseFlag, False=LowerFlag
End Type
    
Public Type defGroup
    GroupName As String
    Queue As Boolean    'Process signals sequentially
End Type
Private Type defSignalAttribute 'Initialled loaded from .ini file [Signal] section
'These are defined again as they are used once the timer is
'running - they are loaded from the UpDown
'These are the same for OFF and ON
    FlagIndex As Long   'Must be the same as the frmMain.Flags(Index) image
                        'A frmMain.SignalTimer(Index) is created for every Flags(Index)
    Type As String  'Class, Finish, Sound, Recall
    Name As String  'Name of the Signal  Class Flag 1
    Image As Picture    'GIF image
    Position As Long    'Flags(Position)
    Group As String  'Flag is positioned below any UP Flag in this Group
    TTL As Long     'Time this flag is displayed in Millisecs
                    'It will be off for the same Period (if more than 1 cycle)
    CyclesRequired As Long  'No of On cycles by timer before creating OFF event
    OnCycles As Long    'Count of on cycles, completed after next off(when timer is enabled)
    TTD As Long         'Time Off
    ImageFilePath As String 'Flag Image
                            'timer must be unique
'If Flag is Raised 2 is used
    Link(1 To 2) As defLink   '1 is used when signal is Raised(UP) & 2 when Lowered(Down)
End Type

Private IniNewEvent As clsEvent    'This is used to keep the variables require to setup
'a New Event from the .ini file

Public IniFileName As String

Public SignalAttributes() As defSignalAttribute
Public Myprofile As clsProfile
Public FixedGroups() As defGroup    'Group predefined for this column
Public ElapsedTime As Long
Public Multiplier As Long
Public RecallIdx As Long    'Keep to remove necessity of looking up at end of time cycle
Private SignalImageFilePath As String

Sub Main()
'    Action.Load (Environ("AllUsersProfile") & "\Application Data\Arundale\RacingSignals\Sequences\" & "ScarboroughMultiple.csv")
    cbOrange = RGB(255, 102, 0)
    
    Multiplier = 20
    Load frmMain
    If frmMain.cboProfile.ListCount = 0 Then
        Unload frmMain  'exit program
    Else
        frmMain.Show
    End If
       
    
'    If IniFileName <> "" Then
'        Call LoadProfile(frmMain.cboProfile.List(frmMain.cboProfile.ListIndex))
'        frmMain.RaceTimer.Enabled = True
'    End If
End Sub
'False if it fails

Public Function TimerOutput(OutputTime As Date) As Boolean
Debug.Print Format$(OutputTime, "hh:mm:ss")
    TimerOutput = True
End Function

'Converts a Nul string to "0"
Public Function NulToZero(TxtIn As String) As String
    If TxtIn = "" Then
        NulToZero = "0"
    Else
        NulToZero = TxtIn
    End If
End Function

Public Function DoTimerEvents(ElapsedTime As Long) As Boolean
Dim MyEvent As clsEvent
'Dim LastEvent As clsEvent
'Dim FirstEvent As clsEvent
Dim NextStartIdx As Long
Stop
    If Myprofile.IsEventDue(ElapsedTime) = False Then
'If ElapsedTime = 0 Then Stop
        Exit Function
    End If
                                    
    DoTimerEvents = True    'Start sequence has started
    For Each MyEvent In Myprofile
'If MyEvent.Signal = 0 Then Stop
        If MyEvent.Second = ElapsedTime Then
Debug.Print MyEvent.Second & " " & MyEvent.Signal & " " & MyEvent.Raised
'MakeSignals will generated any linked signals
'And use the SignalTimer if required
            Call MakeSignals(MyEvent.Signal, MyEvent.Raised)
        End If
    Next MyEvent
    
'    Set FirstEvent = myProfile.FirstEvent
'    Set LastEvent = myProfile.LastEvent

'If Timer Events have started and not finished you cannot Postpone
    With frmMain
        Select Case ElapsedTime
        Case Is >= Myprofile.LastEvent.Second
'Start Sequence finished
#If False Then
            Select Case .cmdRecall.BackColor
            Case Is = vbRed
'Start recall timer (if not running)
                If .SignalTimer(SignalIdx("Recall")).Enabled = False Then
                    Call MakeSignals(SignalIdx("Recall"), True)
                    Call MakeSignals(SignalIdx("Sound"), True)
                End If
            Case Is = vbGreen
'Recall not pressed at moment of start
                .cmdRecall.BackColor = cbDefault
                .cmdFinish.BackColor = vbGreen
                .cmdFinish.SetFocus
            Case Else
'If recall never been pressed
            End Select
#End If
        Case Is >= Myprofile.FirstEvent.Second
'Start Sequence Running
'            .cmdPostpone.BackColor = cbDefault
'Only set it once
#If False Then
            If .cmdRecall.BackColor = cbDefault Then
                .cmdRecall.BackColor = vbGreen
                .cmdRecall.SetFocus
            End If
#End If
        Case Else
'Start Sequence not started
 '           .cmdPostpone.BackColor = vbGreen
 '           .cmdPostpone.Enabled = True
 '           .cmdPostpone.SetFocus
        End Select
    End With
'    Set FirstEvent = Nothing
'    Set LastEvent = Nothing
    
    NextStartIdx = NextStartSignalIdx
    
    If NextStartIdx Then
        frmMain.StatusBar1.Panels(1).Text = "Next Start " & SignalAttributes(NextStartIdx).Name
    Debug.Print "NextStart " & SignalAttributes(NextStartIdx).Name
#If False Then
        frmMain.cmdRecall.Enabled = True
#End If
    Else
'All classes have started
        frmMain.StatusBar1.Panels(1).Text = "All Classes Started"
'Remove the Pospenment signal
'        frmMain.cmdPostpone.BackColor = cbDefault
        
        Debug.Print "All Classes Started"
'Stop
    End If


End Function

'Must use after the MakeSignal has run
Public Function NextStartSignalIdx() As Long
Dim MyEvent As clsEvent

    For Each MyEvent In Myprofile
'If MyEvent.Signal = 0 Then Stop
        If MyEvent.Second <= ElapsedTime Then
            With SignalAttributes(MyEvent.Signal)
                If .Type = "Class" Then
                    If frmMain.Flags(.FlagIndex).Visible = True Then
                        NextStartSignalIdx = .FlagIndex
                        Exit Function
                    End If
                End If
            End With
        End If
    Next MyEvent
 
End Function
'Using Raised because On is a reserved word
'The code is triggered every time a Signal changes state
Public Function MakeSignals(Signal As Long, Raised As Boolean)
Dim LinkIndex As Long
Static LastEventTime As Long    'To keep any status messages for subsequent event at same time
Dim Message As String

'    frmMain.StatusBar1.Panels(2).Text = ""
    If Signal < LBound(SignalAttributes) Or Signal > UBound(SignalAttributes) Then
        MsgBox "Signal " & Signal & " not defined", vbExclamation, "MakeSignals"
        Exit Function
    End If
    
    If frmMain.Flags(Signal).Visible <> Raised Then
'Change the Flag to what has been asked for
        frmMain.Flags(Signal).Visible = Raised
'Don't start the timer until Linked have been  Raised/Lowered

'The Link Flag may have been changed eg 2 Short to 3 Short
'if recall required

'Check if a Link needs changing as well
        
'Choose the Raised or Lowered Link
        If Raised = True Then
            LinkIndex = 2  'Parent signal Raised
        Else
            LinkIndex = 1
        End If

        With SignalAttributes(Signal)
'Set this Link
            With .Link(LinkIndex)
'See if we have a Link (for this visibility)
                If .Flag > 0 Then
'This is re-entrant into this function
                    Call MakeSignals(.Flag, .Raise)
'Stop
                End If
            End With

'If SignalTimer is required for this signal
'And it is not already running, start the timer
'            If frmMain.SignalTimer(Signal).Enabled = False Then
'                If .TTL <> 0 Then
'                    frmMain.SignalTimer(Signal).Enabled = True
'                End If
'            Else
'If the timer is running, don't re-enable the timer
'            End If
        
        Select Case .Type
        Case Is = "Class"
            If Raised = True Then
'A class has a Warning
Debug.Print .Name & " " & "Up"
                Message = .Name & " Up"
            Else
'A class Start
                Message = .Name & " Down"
Debug.Print .Name & " " & "Dn"
            End If
            If LastEventTime <> ElapsedTime Then frmMain.StatusBar1.Panels(2).Text = ""
            If Message <> "" Then
                If frmMain.StatusBar1.Panels(2).Text = "" Then
                    frmMain.StatusBar1.Panels(2).Text = Message
                Else
                    frmMain.StatusBar1.Panels(2).Text = frmMain.StatusBar1.Panels(2).Text & ", " & Message
                End If
            End If
            LastEventTime = ElapsedTime
        Case Else
        End Select
        End With
    Else
'State of Signal has not changed
'Stop
    End If

End Function

Public Function LoadProfile()
Dim i As Long
Dim j As Long
Dim Secs As Long
Dim Ch As Long
Dim nextline As String
Dim Section As String   'Signal=nnn
Dim CleanLine As String
Dim CleanSection As String
Dim arry() As String    'Name=Values (Multiple Values Comma separated)
Dim arry1() As String   'Values in arry(1)
Dim Idx As Long   'Signal is the only section requiring an index (at the moment)
Dim ProfileFileName As String
Dim MySignalTimer As Timer
Dim MyFont As New StdFont
Dim MyPicture As New StdPicture
Dim MyFrame As Frame
Dim CommandFixed As Boolean     'Do not try and reposition this command
Dim SectionError As Boolean

Debug.Print "==============="
    SignalImageFilePath = Environ("AllUsersProfile") & "\Application Data\Arundale\RacingSignals\SignalImages\"

'Start a fresh Profile
    frmMain.RaceTimer.Enabled = False
    
'Clear existing profile
    Set Myprofile = Nothing     'This terminates all clsEvents as well
    Set Myprofile = New clsProfile
    ReDim SignalAttributes(1 To 1)  'this will clear the array
    Call frmMain.ResetSignalTimers
    Call frmMain.ResetCommands
    Call frmMain.ResetFlags
'Set up new profile
    frmMain.Caption = App.EXEName & " [" & App.Major & "." & App.Minor & "." _
    & App.Revision & "] " & frmMain.cboProfile.List(frmMain.cboProfile.ListIndex)
    
    ProfileFileName = Environ("AllUsersProfile") & "\Application Data\Arundale\RacingSignals\Sequences\" _
    & frmMain.cboProfile.List(frmMain.cboProfile.ListIndex) & ".ini"
    Ch = FreeFile
    Open ProfileFileName For Input As #Ch
    Do Until EOF(Ch)
        Line Input #Ch, nextline
'Strip Comments
        i = InStr(1, nextline, "'")
        If i > 0 Then
            CleanLine = Left$(nextline, i - 1)
        Else
            CleanLine = nextline
        End If
'Remove leading & trailing whitespace
        CleanLine = Replace(CleanLine, vbTab, "")
        CleanLine = Trim(CleanLine)
        If CleanLine <> "" Then
Debug.Print CleanLine
            i = InStr(1, CleanLine, "[")
            If (i > 0) Then
'This is Open or Close Section
                j = InStrRev(CleanLine, "]")
                If j < i Then
                    MsgBox "Parse error:" & vbCrLf & nextline, vbCritical, "LoadProfile"
                    CleanLine = ""  'skip this line
                Else
                    CleanSection = Mid$(CleanLine, i + 1, j - i - 1)
'If Cleaned up Section not blank
                    If Len(CleanSection) > 0 Then
                        arry = Split(CleanSection, "=")
                        If Left$(arry(0), 1) <> "/" Then
'Open section (Sets up Section & SectionIndex to be used by next Input Lines
                            If Section <> "" Then   'Already got a Section set
                                MsgBox "Can't open Section " & CleanSection & vbCrLf _
                                & "Section [" & Section & "] is still open", vbExclamation, "LoadProfile"
                            Else
                                Section = arry(0)
                                If Section <> "" Then
'Set up the new section [...]
                                    Select Case Section
                                    Case Is = "Profile", "Event"
                                    Case Is = "Signal"
                                        If IsNumeric(arry(1)) Then
                                            Idx = arry(1)
'Create the Signal Attributes array index
                                            If Idx > UBound(SignalAttributes) Then
                                                ReDim Preserve SignalAttributes(1 To Idx)
                                            Else
'Idx(1) is always created
                                                If Idx > 1 Then
MsgBox "Duplicated Signal detected", vbCritical, "LoadProfile"
                                                    Section = ""
                                                    GoTo Skip_Line
                                                End If
                                            End If
                                            SignalAttributes(Idx).FlagIndex = Idx
'Create a timer for each Signal (even if we dont use it)
                                            Load frmMain.SignalTimer(Idx)
'Create the Command(idx) if it doesn't exist
                                            If CommandExists(Idx) Then
                                                CommandFixed = True
                                            Else
                                                Load frmMain.Commands(Idx)
                                                frmMain.Commands(Idx).Visible = True
                                                frmMain.Commands(Idx).Enabled = True
                                                CommandFixed = False
                                            End If
'Create the Image Control if it doesnt exist
                                        Else
                                            MsgBox "Section " & Section & " has no Index", vbCritical, "LoadProfile"
                                            Section = ""
                                        End If
                                    Case Else
                                        MsgBox "Section " & Section & " not Defined", vbCritical, "LoadProfile"
                                        Section = ""
                                    End Select
                                End If
                            End If
                        Else
'Close section [/...]
                            If Mid$(arry(0), 2) <> Section Then
                                MsgBox "Section " & CleanSection _
                                & " not open", vbExclamation, "LoadProfile"
                            Else
                                Select Case Section
                                Case Is = "Profile"
                                Case Is = "Signal"
'We have to so thus at the end of the section because if CommandVisible has been changed
'we do not want to position it
                                    If CommandFixed = False Then
                                        Call frmMain.PositionCommand(Idx)
                                    End If
'Set command button caption to same as flag
                                    frmMain.Commands(Idx).Caption = SignalAttributes(Idx).Name
'Initially display the flag for debugging the Position
Call frmMain.RaiseRequest(Idx)
                                    Idx = 0       'End of this signal
                                    Section = ""
                                Case Is = "Event"
'With Event we Create the new Event when we have all the values (when the tag is closed)
'NewEvent(Second As Long, Message As Long, Signal As Long, State As Boolean)
                                    With IniNewEvent
                                        Myprofile.NewEvent .Second, .Message, .Signal, .Raised
                                    End With
                                    Set IniNewEvent = Nothing
                                End Select
                                Section = ""
                                Idx = 0
                            End If 'End Close Opened Section Section
                        End If  'Close Section
                    End If  'Clean Section not blank
                End If  'Valid Section parsed
                CleanSection = ""
        
            Else
'Not [Section] or [/Section]
'So it must be a Line within a section
                If Section <> "" Then
'Split the line arry Name=Value1,Value2
                    arry = Split(CleanLine, "=")
                    If UBound(arry) > 0 Then
                        arry1 = Split(arry(1), ",")
'Ensure we have an arry1(1) even if ""
                        ReDim Preserve arry1(2)
                    End If
    
                    Select Case Section
                    Case Is = "Profile"
                        Select Case arry(0)
                        Case Is = "Name"
 'Now use the file name thas is displayed in the Combo box
 '                           frmMain.Caption = App.EXEName & " [" & App.Major & "." & App.Minor & "." _
 '                           & App.Revision & "] " & arry(1)
                        Case Is = "Version"
 'This is the version of the .ini file
                        Case Is = "FlagFixedGroup"
                            FixedGroups(CLng(arry1(0))).GroupName = arry1(1)
                            If arry1(2) = "Sequential" Then
                                FixedGroups(CLng(arry1(0))).Queue = True
                            End If
                        Case Else
                            MsgBox "Invalid " & arry(0) & " in Profile Section " & Section
                        End Select
'Stop
                    Case Is = "Signal"
'Check again we've got an index
                        If Idx > 0 Then
                            Select Case arry(0)
                            Case Is = "Type"
                                SignalAttributes(Idx).Type = arry(1)
                            Case Is = "Name"
                                SignalAttributes(Idx).Name = arry(1)
                            Case Is = "TTL"
                                SignalAttributes(Idx).TTL = arry(1)
'Set TTD as same as TTL (On and Off for same time) - Is a Default
                                SignalAttributes(Idx).TTD = arry(1)
'Set timer interval immediately before it is enabled
'                                frmMain.SignalTimer(Idx).Interval = arry(1)
                            Case Is = "TTD"
                                SignalAttributes(Idx).TTL = arry(1)
                            Case Is = "Cycles"
                                Select Case SignalAttributes(Idx).Type
'Type must have been set in .ini before Cycles
                                Case Is = "Recall", "Postpone"
                                    SignalAttributes(Idx).CyclesRequired = arry(1) / Multiplier
                                Case Else
                                    SignalAttributes(Idx).CyclesRequired = arry(1)
                                End Select
                            Case Is = "UpLink"
                                SignalAttributes(Idx).Link(1).Flag = arry1(0)
                                SignalAttributes(Idx).Link(1).Raise = AtoBool(arry1(1))
                            Case Is = "DownLink"
                                SignalAttributes(Idx).Link(2).Flag = arry1(0)
                                SignalAttributes(Idx).Link(2).Raise = AtoBool(arry1(1))
'Flag attributes
                            Case Is = "Flag"
'Put in Image in the Signal attributes
                                Set SignalAttributes(Idx).Image = LoadPicture(SignalImageFilePath & arry1(0) & ".gif")
                            Case Is = "Position"
'If a specific Position is specific then use this
                                SignalAttributes(Idx).Position = arry(1)
                            Case Is = "Group"
                                SignalAttributes(Idx).Group = arry(1)
'Testing only
                            Case Is = "Raised"
                                frmMain.Flags(Idx).Visible = AtoBool(arry(1))
'CommandAttributes
                            Case Is = "CommandVisible"
                                frmMain.Commands(Idx).Visible = AtoBool(arry(1))
'CommandFrame not currently used
                            Case Is = "CommandFrame"
'                                    Set MyFrame = NametoFrame(arry(1))  'an Object
                                    Select Case arry(1)
                                        Case Is = "Postponement", "Horn"
                                        Set frmMain.Commands(Idx).Container = MyFrame
'                                    Set frmMain.Commands(Idx).Container = frmMain.fraPostponement
'                                    Set frmMain.Commands(Idx).Container = frmMain.fraHorn
                                        frmMain.Commands(Idx).Top = 0
                                        frmMain.Commands(Idx).Left = 0
                                        frmMain.Commands(Idx).Width = 1700
'Position in middle of frame at the bottom
                                        frmMain.Commands(Idx).Move _
                                        (MyFrame.Width _
                                        - frmMain.Commands(Idx).Width) / 2 _
                                        , MyFrame.Height _
                                       - frmMain.Commands(Idx).Height - 100
                                       MyFont.Name = "Verdana"
                                        MyFont.Size = 14
                                        MyFont.Bold = True
                                        Set frmMain.Commands(Idx).Font = MyFont
                                Case Else
MsgBox "Command Frame " & arry(1) & " Container doesnt exist"

                                End Select
'frmMain.Commands(Idx).Visible = True
'Stop
                            Case Else
MsgBox "Invalid " & arry(0) & " in Profile Section " & Section
                            End Select
                        Else
MsgBox "No index " & arry(0) & " in Profile Section " & Section
                        End If  'Got a flag for this index
                    Case Is = "Event"
                        If IniNewEvent Is Nothing Then
                            Set IniNewEvent = New clsEvent
'Enables is only used to enable command buttons
                            IniNewEvent.Enabled = True  'Enable by default
                        End If
                        With IniNewEvent
                            Select Case arry(0)
                            Case Is = "Time"
                                .Second = arry(1) / Multiplier
                            Case Is = "Signal"
                                .Signal = arry(1)
                            Case Is = "Raised"
                                .Raised = AtoBool(arry(1))
                            Case Is = "Enabled"     'Set to true by default
                                .Enabled = AtoBool(arry(1))
                            Case Is = "Message"
                                If .Message <> "" Then .Message = .Message & ", "
                                .Message = .Message & arry(1)
                            Case Else
MsgBox "Invalid " & arry(0) & " in Profile Section " & Section
                            End Select
                        End With
                    Case Else
MsgBox "Invalid Section in Initialisation File"
                    End Select
                Else
                    MsgBox "Line Outside section" & vbCrLf & nextline & vbCrLf, vbExclamation, "LoadProfile"
                End If
            End If
        End If
Skip_Line:
    Loop
    Close #Ch
                
'Check Command Button Signals have been defined
    Call CommandIdx("Horn Short")
    Call CommandIdx("Recall")
    Call CommandIdx("Finish")
        
    frmMain.cmdEvents.Enabled = True
'temp stop    frmMain.RaceTimer.Enabled = True

Debug.Print "LoadEvents " & i

End Function

Private Function FlagRaised(Idx As Long) As Boolean
Dim MyImage As Image
    For Each MyImage In frmMain.Flags
        If MyImage.Index = Idx Then
            FlagRaised = True
            Exit Function
        End If
    Next MyImage
End Function

Private Function CommandExists(Idx As Long) As Boolean
Dim MyCommand As CommandButton
    For Each MyCommand In frmMain.Commands
        If MyCommand.Index = Idx Then
            CommandExists = True
            Exit Function
        End If
    Next MyCommand
End Function

Private Function AtoBool(kb As String) As Boolean
    If kb = "True" Then AtoBool = True
End Function

Public Function SignalIdx(SignalType As String, Optional SignalName As String) As Long
Dim i As Long
Dim kb As String

    For i = 1 To UBound(SignalAttributes)
        If SignalAttributes(i).Type = SignalType Then
            If SignalName <> "" Then
                If SignalAttributes(i).Name = SignalName Then
                    SignalIdx = i
                    Exit Function
                End If
            Else
                SignalIdx = i
                Exit Function
            End If
        End If
    Next i
    kb = "Signal Type " & SignalType
    If SignalName <> "" Then
        kb = kb & ", Signal Name " & SignalName
    End If
    kb = kb & " not found"
MsgBox kb, vbExclamation, "SignalIdx"
End Function

Public Function CommandIdx(CommandName As String) As Long
Dim kb As String
Dim MyCommand As CommandButton
    
    For Each MyCommand In frmMain.Commands
        If MyCommand.Caption = CommandName Then
            CommandIdx = MyCommand.Index
            Exit Function
        End If
    Next MyCommand
MsgBox "Command Button " & CommandName & " not found", vbExclamation, "CommandIdx"
End Function

'Check to see if Timer(index) exists, there seems to be no other way to check
'other than trying to access the index
Public Function TimerExists(Idx As Long) As Boolean
    On Error GoTo NoTimer
    If frmMain.SignalTimer(Idx).Index Then
        TimerExists = True
    End If
NoTimer:
End Function

Private Function UnloadSignalTimers()
Dim oSignalTimer As Timer
Dim Idx As Integer
    For Each oSignalTimer In frmMain.SignalTimer
        Idx = oSignalTimer.Index
'        Set oSignalTimer = Nothing
        If Idx > 0 Then
'            Unload frmMain.SignalTimer(Idx)
            Unload oSignalTimer
        End If
    Next
End Function

Public Function HasIndex(ControlArray As Object, ByVal Index As Integer) As Boolean
    HasIndex = (VarType(ControlArray(Index)) <> vbObject)
End Function



