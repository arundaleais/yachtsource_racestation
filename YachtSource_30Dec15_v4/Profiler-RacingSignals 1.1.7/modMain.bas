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

'Public Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

Public cbOrange As Long

Public Type defLink
    Type As String  'UpLink or DownLink (Replaces using UpLink & DownLink)
    Flag As Long    'The Link Flag that is associated
    Raise As Boolean    'True = Raise Linked Flag, False=Lower Linked Flag
    Temp As Boolean     'True = DeleteLink when actioned
End Type

Public Type defFlag  'Properties where to display the Image (Blank if not displayed)
    Pos As Long         'Flag(idx) = 0 if not Up (Both Row & Col > 0)
    Row As Long         'Row on which to display the flag,  Set on UP, Clear after down
    Col As Long         'Col on which to display the flag
    FixedRow As Long 'Position is preset on load, so do not clear on event
    FixedCol As Long
    Queue As Boolean    'Queue any Flag Event (Sounds, Controllers), do not clear on event
    Changed As Boolean  'Used to generate Controller Event  'Clear after event handled
End Type

Private Type defGroupDefault  'Are applied to the above when Signal Is Loaded (by LoadProfile)
    Group As String     'Group to which these defaults are applied
    FixedRow As Long
    FixedCol As Long
    Queue As Boolean
End Type

Private Type defController
    Name As String
    IpAddress As String
    On As String
    Off As String
    Connection As String
    Sound As String
End Type

'Public Type defGroup    'Only defined for Fixed Groups (Sound, Lights)
'    GroupName As String
'    Queue As Boolean    'Process signals sequentially
'End Type

Private Type defSignalAttribute 'Initialled loaded from .ini file [Signal] section
'These are defined again as they are used once the timer is
'running - they are loaded from the UpDown
'These are the same for OFF and ON
    Type As String  'Class, Finish, Sound, Recall
    Name As String  'Name of the Signal  Class Flag 1
    Image As Picture    'GIF image
    Flag As defFlag   'Flag Attributes
    Group As String  'Flag is positioned below any UP Flag in this Group
    TTL As Long     'Time this flag is displayed in Millisecs
                    'It will be off for the same Period (if more than 1 cycle)
    CyclesRequired As Long  'No of On cycles by timer before creating OFF event
    OnCycles As Long    'Count of on cycles, completed after next off(when timer is enabled)
    TTD As Long         'Time Off
    ImageFilePath As String 'Flag Image
                            'timer must be unique
    Links() As defLink   'Used when signal is Raised(UP) & when Lowered(Down)
    Controller As Long  'This is the controller used when the flag is visible
                        '-1 = no linked controller
End Type

Private IniNewEvent As clsEvent    'This is used to keep the variables require to setup
'a New Event from the .ini file

Public IniFileName As String

Public Loading As Boolean   'Suppress Queueing Commands

Public Controllers() As defController
Public SignalAttributes() As defSignalAttribute
Public Myprofile As clsProfile
'Public FixedGroups() As defGroup    'Group predefined for this column
Public ElapsedTime As Long
Public Multiplier As Long
Public RecallIdx As Long    'Keep to remove necessity of looking up at end of time cycle
Public RowCount As Long
Public ColCount As Long
Public ColCountFree As Long 'ColCount less any Fixed Cols
Private SignalImageFilePath As String
Private DebugLoadProfile As Boolean

Sub Main()
'    Action.Load (Environ("AllUsersProfile") & "\Application Data\Arundale\RacingSignals\Sequences\" & "ScarboroughMultiple.csv")
    vbwInitializeProfiler ' Initialize VB Watch
vbwProfiler.vbwProcIn 55
vbwProfiler.vbwExecuteLine 883
    cbOrange = RGB(255, 102, 0)

vbwProfiler.vbwExecuteLine 884
    Multiplier = 20
vbwProfiler.vbwExecuteLine 885
    Load frmMain
vbwProfiler.vbwExecuteLine 886
    If frmMain.cboProfile.ListCount = 0 Then
vbwProfiler.vbwExecuteLine 887
        Unload frmMain  'exit program
    Else
vbwProfiler.vbwExecuteLine 888 'B
vbwProfiler.vbwExecuteLine 889
        frmMain.Show
    End If
vbwProfiler.vbwExecuteLine 890 'B


'    If IniFileName <> "" Then
'        Call LoadProfile(frmMain.cboProfile.List(frmMain.cboProfile.ListIndex))
'        frmMain.RaceTimer.Enabled = True
'    End If
vbwProfiler.vbwProcOut 55
vbwProfiler.vbwExecuteLine 891
End Sub
'False if it fails

Public Function TimerOutput(OutputTime As Date) As Boolean
'Debug.Print Format$(OutputTime, "hh:mm:ss")
vbwProfiler.vbwProcIn 56
vbwProfiler.vbwExecuteLine 892
    TimerOutput = True
vbwProfiler.vbwProcOut 56
vbwProfiler.vbwExecuteLine 893
End Function

'Converts a Nul string to "0"
Public Function NulToZero(TxtIn As String) As String
vbwProfiler.vbwProcIn 57
vbwProfiler.vbwExecuteLine 894
    If TxtIn = "" Then
vbwProfiler.vbwExecuteLine 895
        NulToZero = "0"
    Else
vbwProfiler.vbwExecuteLine 896 'B
vbwProfiler.vbwExecuteLine 897
        NulToZero = TxtIn
    End If
vbwProfiler.vbwExecuteLine 898 'B
vbwProfiler.vbwProcOut 57
vbwProfiler.vbwExecuteLine 899
End Function

Public Function DoTimerEvents(ElapsedTime As Long) As Boolean
vbwProfiler.vbwProcIn 58
Dim MyEvent As clsEvent
'Dim LastEvent As clsEvent
'Dim FirstEvent As clsEvent
Dim NextStartIdx As Long
'Stop
vbwProfiler.vbwExecuteLine 900
    If Myprofile.IsEventDue(ElapsedTime) = False Then
'If ElapsedTime = 0 Then Stop
vbwProfiler.vbwProcOut 58
vbwProfiler.vbwExecuteLine 901
        Exit Function
    End If
vbwProfiler.vbwExecuteLine 902 'B

vbwProfiler.vbwExecuteLine 903
    DoTimerEvents = True    'Start sequence has started
vbwProfiler.vbwExecuteLine 904
    For Each MyEvent In Myprofile
'If MyEvent.Signal = 0 Then Stop
vbwProfiler.vbwExecuteLine 905
        If MyEvent.Second = ElapsedTime Then
vbwProfiler.vbwExecuteLine 906
Debug.Print MyEvent.Second & " " & MyEvent.Signal & " " & MyEvent.Raised
'MakeSignals will generated any linked signals
'And use the SignalTimer if required
vbwProfiler.vbwExecuteLine 907
            Call MakeSignals(MyEvent.Signal, MyEvent.Raised)
        End If
vbwProfiler.vbwExecuteLine 908 'B
vbwProfiler.vbwExecuteLine 909
    Next MyEvent

'    Set FirstEvent = myProfile.FirstEvent
'    Set LastEvent = myProfile.LastEvent

'If Timer Events have started and not finished you cannot Postpone
vbwProfiler.vbwExecuteLine 910
    With frmMain
vbwProfiler.vbwExecuteLine 911
        Select Case ElapsedTime
'vbwLine 912:        Case Is >= Myprofile.LastEvent.Second
        Case Is >= IIf(vbwProfiler.vbwExecuteLine(912), VBWPROFILER_EMPTY, _
        Myprofile.LastEvent.Second)
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
'vbwLine 913:        Case Is >= Myprofile.FirstEvent.Second
        Case Is >= IIf(vbwProfiler.vbwExecuteLine(913), VBWPROFILER_EMPTY, _
        Myprofile.FirstEvent.Second)
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
vbwProfiler.vbwExecuteLine 914 'B
'Start Sequence not started
 '           .cmdPostpone.BackColor = vbGreen
 '           .cmdPostpone.Enabled = True
 '           .cmdPostpone.SetFocus
        End Select
vbwProfiler.vbwExecuteLine 915 'B
vbwProfiler.vbwExecuteLine 916
    End With
'    Set FirstEvent = Nothing
'    Set LastEvent = Nothing

#If False Then
    NextStartIdx = NextStartSignalIdx
#End If
vbwProfiler.vbwExecuteLine 917
    If NextStartIdx Then
vbwProfiler.vbwExecuteLine 918
        frmMain.StatusBar1.Panels(1).Text = "Next Start " & SignalAttributes(NextStartIdx).Name
vbwProfiler.vbwExecuteLine 919
    Debug.Print "NextStart " & SignalAttributes(NextStartIdx).Name
#If False Then
        frmMain.cmdRecall.Enabled = True
#End If
    Else
vbwProfiler.vbwExecuteLine 920 'B
'All classes have started
vbwProfiler.vbwExecuteLine 921
        frmMain.StatusBar1.Panels(1).Text = "All Classes Started"
'Remove the Pospenment signal
'        frmMain.cmdPostpone.BackColor = cbDefault

vbwProfiler.vbwExecuteLine 922
        Debug.Print "All Classes Started"
'Stop
    End If
vbwProfiler.vbwExecuteLine 923 'B


vbwProfiler.vbwProcOut 58
vbwProfiler.vbwExecuteLine 924
End Function

#If False Then
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
#End If
'Using Raised because On is a reserved word
'The code is triggered every time a Signal changes state
Public Function MakeSignals(Signal As Long, Raised As Boolean)
vbwProfiler.vbwProcIn 59
Dim LinkIndex As Long
Static LastEventTime As Long    'To keep any status messages for subsequent event at same time
Dim Message As String
vbwProfiler.vbwExecuteLine 925
Stop
#If False Then
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
#End If
vbwProfiler.vbwProcOut 59
vbwProfiler.vbwExecuteLine 926
End Function

Public Function LoadProfile()
vbwProfiler.vbwProcIn 60
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
Dim Cidx As Long    'Controller() index
Dim MySignalTimer As Timer
Dim MyFont As New StdFont
Dim MyPicture As New StdPicture
Dim MyFrame As Frame
Dim MyLink As defLink
Dim CommandFixed As Boolean     'Do not try and reposition this command
Dim SectionError As Boolean
Dim GroupDefaults() As defGroupDefault
Dim Lidx As Long
vbwProfiler.vbwExecuteLine 927
Debug.Print "==============="
vbwProfiler.vbwExecuteLine 928
    SignalImageFilePath = Environ("AllUsersProfile") & "\Application Data\Arundale\RacingSignals\SignalImages\"

'Start a fresh Profile
vbwProfiler.vbwExecuteLine 929
    frmMain.RaceTimer.Enabled = False
vbwProfiler.vbwExecuteLine 930
    Loading = True
'Clear existing profile

'Set up Controller(0) for the Horn, even if nothing connected
vbwProfiler.vbwExecuteLine 931
    ReDim Controllers(0)
vbwProfiler.vbwExecuteLine 932
    ReDim GroupDefaults(0)
vbwProfiler.vbwExecuteLine 933
    Set Myprofile = Nothing     'This terminates all clsEvents as well
vbwProfiler.vbwExecuteLine 934
    Set Myprofile = New clsProfile
vbwProfiler.vbwExecuteLine 935
    ReDim SignalAttributes(1 To 1)  'this will clear the array
'Set up Controller(0) for the Horn, even if nothing connected
'Min of 1 controller
vbwProfiler.vbwExecuteLine 936
    Call frmMain.ResetSignalTimers
vbwProfiler.vbwExecuteLine 937
    Call frmMain.ResetCommands
vbwProfiler.vbwExecuteLine 938
    Call frmMain.ResetFlags
'Set up new profile
vbwProfiler.vbwExecuteLine 939
    frmMain.Caption = App.EXEName & " [" & App.Major & "." & App.Minor & "." _
    & App.Revision & "] " & frmMain.cboProfile.List(frmMain.cboProfile.ListIndex)

vbwProfiler.vbwExecuteLine 940
    ProfileFileName = Environ("AllUsersProfile") & "\Application Data\Arundale\RacingSignals\Sequences\" _
    & frmMain.cboProfile.List(frmMain.cboProfile.ListIndex) & ".ini"
vbwProfiler.vbwExecuteLine 941
    Ch = FreeFile
vbwProfiler.vbwExecuteLine 942
    Open ProfileFileName For Input As #Ch
'vbwLine 943:    Do Until EOF(Ch)
    Do Until vbwProfiler.vbwExecuteLine(943) Or EOF(Ch)
vbwProfiler.vbwExecuteLine 944
        Line Input #Ch, nextline
'Strip Comments
vbwProfiler.vbwExecuteLine 945
        i = InStr(1, nextline, "'")
vbwProfiler.vbwExecuteLine 946
        If i > 0 Then
vbwProfiler.vbwExecuteLine 947
            CleanLine = Left$(nextline, i - 1)
        Else
vbwProfiler.vbwExecuteLine 948 'B
vbwProfiler.vbwExecuteLine 949
            CleanLine = nextline
        End If
vbwProfiler.vbwExecuteLine 950 'B
'Remove leading & trailing whitespace
vbwProfiler.vbwExecuteLine 951
        CleanLine = Replace(CleanLine, vbTab, "")
vbwProfiler.vbwExecuteLine 952
        CleanLine = Trim(CleanLine)
vbwProfiler.vbwExecuteLine 953
        If CleanLine <> "" Then
vbwProfiler.vbwExecuteLine 954
Debug.Print CleanLine
vbwProfiler.vbwExecuteLine 955
            i = InStr(1, CleanLine, "[")
vbwProfiler.vbwExecuteLine 956
            If (i > 0) Then
'This is Open or Close Section
vbwProfiler.vbwExecuteLine 957
                j = InStrRev(CleanLine, "]")
vbwProfiler.vbwExecuteLine 958
                If j < i Then
vbwProfiler.vbwExecuteLine 959
                    MsgBox "Parse error:" & vbCrLf & nextline, vbCritical, "LoadProfile"
vbwProfiler.vbwExecuteLine 960
                    CleanLine = ""  'skip this line
                Else
vbwProfiler.vbwExecuteLine 961 'B
vbwProfiler.vbwExecuteLine 962
                    CleanSection = Mid$(CleanLine, i + 1, j - i - 1)
'If Cleaned up Section not blank
vbwProfiler.vbwExecuteLine 963
                    If Len(CleanSection) > 0 Then
vbwProfiler.vbwExecuteLine 964
                        arry = Split(CleanSection, "=")
vbwProfiler.vbwExecuteLine 965
                        If Left$(arry(0), 1) <> "/" Then
'Open section (Sets up Section & SectionIndex to be used by next Input Lines
vbwProfiler.vbwExecuteLine 966
                            If Section <> "" Then   'Already got a Section set
vbwProfiler.vbwExecuteLine 967
                                MsgBox "Can't open Section " & CleanSection & vbCrLf _
                                & "Section [" & Section & "] is still open", vbExclamation, "LoadProfile"
                            Else
vbwProfiler.vbwExecuteLine 968 'B
vbwProfiler.vbwExecuteLine 969
                                Section = arry(0)
vbwProfiler.vbwExecuteLine 970
                                If Section <> "" Then
'Set up the new section [...]
vbwProfiler.vbwExecuteLine 971
                                    CommandFixed = False
vbwProfiler.vbwExecuteLine 972
                                    Select Case Section
'vbwLine 973:                                    Case Is = "Profile", "Event"
                                    Case Is = IIf(vbwProfiler.vbwExecuteLine(973), VBWPROFILER_EMPTY, _
        "Profile"), "Event"
'vbwLine 974:                                    Case Is = "Controller"
                                    Case Is = IIf(vbwProfiler.vbwExecuteLine(974), VBWPROFILER_EMPTY, _
        "Controller")
'Stop
vbwProfiler.vbwExecuteLine 975
                                        If IsNumeric(arry(1)) Then
vbwProfiler.vbwExecuteLine 976
                                            Cidx = arry(1)
vbwProfiler.vbwExecuteLine 977
                                            If Cidx > UBound(Controllers) Then
vbwProfiler.vbwExecuteLine 978
                                                ReDim Preserve Controllers(Cidx)
                                            End If
vbwProfiler.vbwExecuteLine 979 'B
                                        End If
vbwProfiler.vbwExecuteLine 980 'B
vbwProfiler.vbwExecuteLine 981
                                        Controllers(Cidx).Name = "Controller (" & Cidx & ")"
'vbwLine 982:                                    Case Is = "Signal"
                                    Case Is = IIf(vbwProfiler.vbwExecuteLine(982), VBWPROFILER_EMPTY, _
        "Signal")
vbwProfiler.vbwExecuteLine 983
                                        Lidx = 0
vbwProfiler.vbwExecuteLine 984
                                        If IsNumeric(arry(1)) Then
vbwProfiler.vbwExecuteLine 985
                                            Idx = arry(1)
'Create the Signal Attributes array index
vbwProfiler.vbwExecuteLine 986
                                            If Idx > UBound(SignalAttributes) Then
vbwProfiler.vbwExecuteLine 987
                                                i = UBound(SignalAttributes)
vbwProfiler.vbwExecuteLine 988
                                                ReDim Preserve SignalAttributes(1 To Idx)
                                            Else
vbwProfiler.vbwExecuteLine 989 'B
'Idx(1) is always created
vbwProfiler.vbwExecuteLine 990
                                                If Idx > 1 Then
'Signals must be created in ascending order
vbwProfiler.vbwExecuteLine 991
MsgBox "Duplicated Signal(" & Idx & ") detected", vbCritical, "LoadProfile"
vbwProfiler.vbwExecuteLine 992
                                                    Section = ""
vbwProfiler.vbwExecuteLine 993
                                                    GoTo Skip_Line
                                                End If
vbwProfiler.vbwExecuteLine 994 'B
                                            End If
vbwProfiler.vbwExecuteLine 995 'B
'Default is None
vbwProfiler.vbwExecuteLine 996
                                            SignalAttributes(Idx).Controller = -1
'Create a timer for each Signal (even if we dont use it)
vbwProfiler.vbwExecuteLine 997
                                            Load frmMain.SignalTimer(Idx)
'Create the Command(idx) if it doesn't exist
vbwProfiler.vbwExecuteLine 998
                                            If CommandExists(Idx) Then
vbwProfiler.vbwExecuteLine 999
                                                CommandFixed = True
                                            Else
vbwProfiler.vbwExecuteLine 1000 'B
vbwProfiler.vbwExecuteLine 1001
                                                Load frmMain.Commands(Idx)
vbwProfiler.vbwExecuteLine 1002
                                                frmMain.Commands(Idx).Visible = True
vbwProfiler.vbwExecuteLine 1003
                                                frmMain.Commands(Idx).Enabled = True
                                            End If
vbwProfiler.vbwExecuteLine 1004 'B
'Create the Image Control if it doesnt exist
                                        Else
vbwProfiler.vbwExecuteLine 1005 'B
vbwProfiler.vbwExecuteLine 1006
                                            MsgBox "Section " & Section & " has no Index", vbCritical, "LoadProfile"
vbwProfiler.vbwExecuteLine 1007
                                            Section = ""
                                        End If
vbwProfiler.vbwExecuteLine 1008 'B
                                    Case Else
vbwProfiler.vbwExecuteLine 1009 'B
vbwProfiler.vbwExecuteLine 1010
                                        MsgBox "Section " & Section & " not Defined", vbCritical, "LoadProfile"
vbwProfiler.vbwExecuteLine 1011
                                        Section = ""
                                    End Select
vbwProfiler.vbwExecuteLine 1012 'B
                                End If
vbwProfiler.vbwExecuteLine 1013 'B
                            End If
vbwProfiler.vbwExecuteLine 1014 'B
                        Else
vbwProfiler.vbwExecuteLine 1015 'B
'Close section [/...]
vbwProfiler.vbwExecuteLine 1016
                            If Mid$(arry(0), 2) <> Section Then
vbwProfiler.vbwExecuteLine 1017
                                MsgBox "Section " & CleanSection _
                                & " not open", vbExclamation, "LoadProfile"
                            Else
vbwProfiler.vbwExecuteLine 1018 'B
vbwProfiler.vbwExecuteLine 1019
                                Select Case Section
'vbwLine 1020:                                Case Is = "Profile"
                                Case Is = IIf(vbwProfiler.vbwExecuteLine(1020), VBWPROFILER_EMPTY, _
        "Profile")
'vbwLine 1021:                                Case Is = "Controller"
                                Case Is = IIf(vbwProfiler.vbwExecuteLine(1021), VBWPROFILER_EMPTY, _
        "Controller")
'Stop
vbwProfiler.vbwExecuteLine 1022
                                    Cidx = 0    'End of this Controller Default
'vbwLine 1023:                                Case Is = "Signal"
                                Case Is = IIf(vbwProfiler.vbwExecuteLine(1023), VBWPROFILER_EMPTY, _
        "Signal")
'We have to so this at the end of the section because if CommandVisible has been changed
'we do not want to position it
vbwProfiler.vbwExecuteLine 1024
                                    If CommandFixed = False Then
vbwProfiler.vbwExecuteLine 1025
                                        Call frmMain.PositionCommand(Idx)
                                    End If
vbwProfiler.vbwExecuteLine 1026 'B
'Set command button caption to same as flag
vbwProfiler.vbwExecuteLine 1027
                                    frmMain.Commands(Idx).Caption = SignalAttributes(Idx).Name
vbwProfiler.vbwExecuteLine 1028
                                    Idx = 0       'End of this signal
'vbwLine 1029:                                Case Is = "Event"
                                Case Is = IIf(vbwProfiler.vbwExecuteLine(1029), VBWPROFILER_EMPTY, _
        "Event")
'With Event we Create the new Event when we have all the values (when the tag is closed)
'NewEvent(Second As Long, Message As Long, Signal As Long, State As Boolean)
vbwProfiler.vbwExecuteLine 1030
                                    With IniNewEvent
vbwProfiler.vbwExecuteLine 1031
                                        Myprofile.NewEvent .Second, .Message, .Signal, .Raised
vbwProfiler.vbwExecuteLine 1032
                                    End With
vbwProfiler.vbwExecuteLine 1033
                                    Set IniNewEvent = Nothing
                                End Select
vbwProfiler.vbwExecuteLine 1034 'B
vbwProfiler.vbwExecuteLine 1035
                                Section = ""
vbwProfiler.vbwExecuteLine 1036
                                Idx = 0
                            End If 'End Close Opened Section Section
vbwProfiler.vbwExecuteLine 1037 'B
                        End If  'Close Section
vbwProfiler.vbwExecuteLine 1038 'B
                    End If  'Clean Section not blank
vbwProfiler.vbwExecuteLine 1039 'B
                End If  'Valid Section parsed
vbwProfiler.vbwExecuteLine 1040 'B
vbwProfiler.vbwExecuteLine 1041
                CleanSection = ""

            Else
vbwProfiler.vbwExecuteLine 1042 'B
'Not [Section] or [/Section]
'So it must be a Line within a section
vbwProfiler.vbwExecuteLine 1043
                If Section <> "" Then
'Split the line arry Name=Value1,Value2
vbwProfiler.vbwExecuteLine 1044
                    arry = Split(CleanLine, "=")
vbwProfiler.vbwExecuteLine 1045
                    ReDim arry1(0)
vbwProfiler.vbwExecuteLine 1046
                    If UBound(arry) > 0 Then
vbwProfiler.vbwExecuteLine 1047
                        arry1 = Split(arry(1), ",")
                    End If
vbwProfiler.vbwExecuteLine 1048 'B

vbwProfiler.vbwExecuteLine 1049
                    Select Case Section
'vbwLine 1050:                    Case Is = "Profile"
                    Case Is = IIf(vbwProfiler.vbwExecuteLine(1050), VBWPROFILER_EMPTY, _
        "Profile")
vbwProfiler.vbwExecuteLine 1051
                        Select Case arry(0)
'vbwLine 1052:                        Case Is = "Name"
                        Case Is = IIf(vbwProfiler.vbwExecuteLine(1052), VBWPROFILER_EMPTY, _
        "Name")
 'Now use the file name thas is displayed in the Combo box
 '                           frmMain.Caption = App.EXEName & " [" & App.Major & "." & App.Minor & "." _
 '                           & App.Revision & "] " & arry(1)
'vbwLine 1053:                        Case Is = "Version"
                        Case Is = IIf(vbwProfiler.vbwExecuteLine(1053), VBWPROFILER_EMPTY, _
        "Version")
 'This is the version of the .ini file
'vbwLine 1054:                        Case Is = "GroupDefault"
                        Case Is = IIf(vbwProfiler.vbwExecuteLine(1054), VBWPROFILER_EMPTY, _
        "GroupDefault")
vbwProfiler.vbwExecuteLine 1055
                            For i = 1 To UBound(GroupDefaults)
vbwProfiler.vbwExecuteLine 1056
                                If GroupDefaults(i).Group = arry1(0) Then
vbwProfiler.vbwExecuteLine 1057
                                    Exit For
                                End If
vbwProfiler.vbwExecuteLine 1058 'B
vbwProfiler.vbwExecuteLine 1059
                            Next i
vbwProfiler.vbwExecuteLine 1060
                            If i > UBound(GroupDefaults) Then
vbwProfiler.vbwExecuteLine 1061
                                ReDim Preserve GroupDefaults(i)
                            End If
vbwProfiler.vbwExecuteLine 1062 'B
vbwProfiler.vbwExecuteLine 1063
                            GroupDefaults(i).Group = arry1(0)
vbwProfiler.vbwExecuteLine 1064
                            For j = 1 To UBound(arry1)
vbwProfiler.vbwExecuteLine 1065
                                Select Case arry1(j)
'vbwLine 1066:                                Case Is = "LastCol"
                                Case Is = IIf(vbwProfiler.vbwExecuteLine(1066), VBWPROFILER_EMPTY, _
        "LastCol")
vbwProfiler.vbwExecuteLine 1067
                                    GroupDefaults(i).FixedCol = ColCount
vbwProfiler.vbwExecuteLine 1068
                                    ColCountFree = ColCount - 1
'vbwLine 1069:                                Case Is = "LastCol-1"
                                Case Is = IIf(vbwProfiler.vbwExecuteLine(1069), VBWPROFILER_EMPTY, _
        "LastCol-1")
vbwProfiler.vbwExecuteLine 1070
                                    GroupDefaults(i).FixedCol = ColCount - 1
vbwProfiler.vbwExecuteLine 1071
                                    ColCountFree = ColCount - 2
'vbwLine 1072:                                Case Is = "Row1"
                                Case Is = IIf(vbwProfiler.vbwExecuteLine(1072), VBWPROFILER_EMPTY, _
        "Row1")
vbwProfiler.vbwExecuteLine 1073
                                    GroupDefaults(i).FixedRow = 1
'vbwLine 1074:                                Case Is = "Row2"
                                Case Is = IIf(vbwProfiler.vbwExecuteLine(1074), VBWPROFILER_EMPTY, _
        "Row2")
vbwProfiler.vbwExecuteLine 1075
                                    GroupDefaults(i).FixedRow = 2
'vbwLine 1076:                                Case Is = "Row3"
                                Case Is = IIf(vbwProfiler.vbwExecuteLine(1076), VBWPROFILER_EMPTY, _
        "Row3")
vbwProfiler.vbwExecuteLine 1077
                                    GroupDefaults(i).FixedRow = 3
'vbwLine 1078:                                Case Is = "Row4"
                                Case Is = IIf(vbwProfiler.vbwExecuteLine(1078), VBWPROFILER_EMPTY, _
        "Row4")
vbwProfiler.vbwExecuteLine 1079
                                    GroupDefaults(i).FixedRow = 4
'vbwLine 1080:                                Case Is = "Queue"
                                Case Is = IIf(vbwProfiler.vbwExecuteLine(1080), VBWPROFILER_EMPTY, _
        "Queue")
vbwProfiler.vbwExecuteLine 1081
                                    GroupDefaults(i).Queue = True
                                Case Else
vbwProfiler.vbwExecuteLine 1082 'B
vbwProfiler.vbwExecuteLine 1083
                                    MsgBox "Invalid " & arry1(j) & " in Profile Section " & Section
                                End Select
vbwProfiler.vbwExecuteLine 1084 'B
vbwProfiler.vbwExecuteLine 1085
                            Next j
                        Case Else
vbwProfiler.vbwExecuteLine 1086 'B
vbwProfiler.vbwExecuteLine 1087
                            MsgBox "Invalid " & arry(0) & " in Profile Section " & Section
                        End Select
vbwProfiler.vbwExecuteLine 1088 'B
'vbwLine 1089:                    Case Is = "Controller"
                    Case Is = IIf(vbwProfiler.vbwExecuteLine(1089), VBWPROFILER_EMPTY, _
        "Controller")
vbwProfiler.vbwExecuteLine 1090
                        Select Case arry(0)
'vbwLine 1091:                        Case Is = "Name"
                        Case Is = IIf(vbwProfiler.vbwExecuteLine(1091), VBWPROFILER_EMPTY, _
        "Name")
vbwProfiler.vbwExecuteLine 1092
                            Controllers(Cidx).Name = arry(1)
'vbwLine 1093:                        Case Is = "IpAddress"
                        Case Is = IIf(vbwProfiler.vbwExecuteLine(1093), VBWPROFILER_EMPTY, _
        "IpAddress")
vbwProfiler.vbwExecuteLine 1094
                            Controllers(Cidx).IpAddress = arry(1)
'vbwLine 1095:                        Case Is = "On"
                        Case Is = IIf(vbwProfiler.vbwExecuteLine(1095), VBWPROFILER_EMPTY, _
        "On")
vbwProfiler.vbwExecuteLine 1096
                            Controllers(Cidx).On = arry(1)
'vbwLine 1097:                        Case Is = "Off"
                        Case Is = IIf(vbwProfiler.vbwExecuteLine(1097), VBWPROFILER_EMPTY, _
        "Off")
vbwProfiler.vbwExecuteLine 1098
                            Controllers(Cidx).Off = arry(1)
'vbwLine 1099:                        Case Is = "Connection"
                        Case Is = IIf(vbwProfiler.vbwExecuteLine(1099), VBWPROFILER_EMPTY, _
        "Connection")
vbwProfiler.vbwExecuteLine 1100
                            Controllers(Cidx).Connection = arry(1)
'vbwLine 1101:                        Case Is = "Sound"
                        Case Is = IIf(vbwProfiler.vbwExecuteLine(1101), VBWPROFILER_EMPTY, _
        "Sound")
'Check PC has sound available
vbwProfiler.vbwExecuteLine 1102
                            If HasSound = True Then
vbwProfiler.vbwExecuteLine 1103
                                If FileExists(Environ("AllUsersProfile") & "\Application Data\Arundale\RacingSignals\Sounds\" & arry(1)) Then
vbwProfiler.vbwExecuteLine 1104
                                    Controllers(Cidx).Sound = arry(1)
vbwProfiler.vbwExecuteLine 1105
                                    SoundFilePath = Environ("AllUsersProfile") & "\Application Data\Arundale\RacingSignals\Sounds\"
                                Else
vbwProfiler.vbwExecuteLine 1106 'B
vbwProfiler.vbwExecuteLine 1107
MsgBox "Sound File " & arry1(0) & " doesnt exist"
                                End If
vbwProfiler.vbwExecuteLine 1108 'B
                            End If
vbwProfiler.vbwExecuteLine 1109 'B
                        Case Else
vbwProfiler.vbwExecuteLine 1110 'B
vbwProfiler.vbwExecuteLine 1111
                            MsgBox "Invalid " & arry(0) & " in Controller Section " & Section
                        End Select
vbwProfiler.vbwExecuteLine 1112 'B
'vbwLine 1113:                    Case Is = "Signal"
                    Case Is = IIf(vbwProfiler.vbwExecuteLine(1113), VBWPROFILER_EMPTY, _
        "Signal")
'Check again we've got an index
vbwProfiler.vbwExecuteLine 1114
                        If Idx > 0 Then
vbwProfiler.vbwExecuteLine 1115
                            Select Case arry(0)
'vbwLine 1116:                            Case Is = "Type"
                            Case Is = IIf(vbwProfiler.vbwExecuteLine(1116), VBWPROFILER_EMPTY, _
        "Type")
vbwProfiler.vbwExecuteLine 1117
                                SignalAttributes(Idx).Type = arry(1)
'vbwLine 1118:                            Case Is = "Name"
                            Case Is = IIf(vbwProfiler.vbwExecuteLine(1118), VBWPROFILER_EMPTY, _
        "Name")
vbwProfiler.vbwExecuteLine 1119
                                SignalAttributes(Idx).Name = arry(1)
'vbwLine 1120:                            Case Is = "TTL"
                            Case Is = IIf(vbwProfiler.vbwExecuteLine(1120), VBWPROFILER_EMPTY, _
        "TTL")
vbwProfiler.vbwExecuteLine 1121
                                SignalAttributes(Idx).TTL = arry(1)
'Set TTD as same as TTL (On and Off for same time) - Is a Default
vbwProfiler.vbwExecuteLine 1122
                                SignalAttributes(Idx).TTD = arry(1)
'Set timer interval immediately before it is enabled
'                                frmMain.SignalTimer(Idx).Interval = arry(1)
'vbwLine 1123:                            Case Is = "TTD"
                            Case Is = IIf(vbwProfiler.vbwExecuteLine(1123), VBWPROFILER_EMPTY, _
        "TTD")
vbwProfiler.vbwExecuteLine 1124
                                SignalAttributes(Idx).TTD = arry(1)
'vbwLine 1125:                            Case Is = "Cycles"
                            Case Is = IIf(vbwProfiler.vbwExecuteLine(1125), VBWPROFILER_EMPTY, _
        "Cycles")
vbwProfiler.vbwExecuteLine 1126
                                Select Case SignalAttributes(Idx).Type
'Type must have been set in .ini before Cycles
'vbwLine 1127:                                Case Is = "Recall", "Postpone"
                                Case Is = IIf(vbwProfiler.vbwExecuteLine(1127), VBWPROFILER_EMPTY, _
        "Recall"), "Postpone"
vbwProfiler.vbwExecuteLine 1128
                                    SignalAttributes(Idx).CyclesRequired = arry(1) / Multiplier
                                Case Else
vbwProfiler.vbwExecuteLine 1129 'B
vbwProfiler.vbwExecuteLine 1130
                                    SignalAttributes(Idx).CyclesRequired = arry(1)
                                End Select
vbwProfiler.vbwExecuteLine 1131 'B
'vbwLine 1132:                            Case Is = "UpLink", "DownLink"
                            Case Is = IIf(vbwProfiler.vbwExecuteLine(1132), VBWPROFILER_EMPTY, _
        "UpLink"), "DownLink"
vbwProfiler.vbwExecuteLine 1133
                                MyLink.Type = arry(0)
vbwProfiler.vbwExecuteLine 1134
                                MyLink.Flag = arry1(0)
vbwProfiler.vbwExecuteLine 1135
                                If UBound(arry1) > 0 Then
vbwProfiler.vbwExecuteLine 1136
                                    MyLink.Raise = AtoBool(arry1(1))
vbwProfiler.vbwExecuteLine 1137
                                    MyLink.Temp = False
'Create the Links(Next Link Index)
vbwProfiler.vbwExecuteLine 1138
                                    Call CreateLink(Idx, MyLink)
                                Else
vbwProfiler.vbwExecuteLine 1139 'B
vbwProfiler.vbwExecuteLine 1140
MsgBox "Flag (" & Idx & "), " & arry(0) & " requires True or False"
                                End If
vbwProfiler.vbwExecuteLine 1141 'B
'vbwLine 1142:                            Case Is = "Controller"
                            Case Is = IIf(vbwProfiler.vbwExecuteLine(1142), VBWPROFILER_EMPTY, _
        "Controller")
vbwProfiler.vbwExecuteLine 1143
                                SignalAttributes(Idx).Controller = arry(1)
'Flag attributes
'vbwLine 1144:                            Case Is = "Flag"
                            Case Is = IIf(vbwProfiler.vbwExecuteLine(1144), VBWPROFILER_EMPTY, _
        "Flag")
'Put in Image in the Signal attributes
vbwProfiler.vbwExecuteLine 1145
                                If FileExists(SignalImageFilePath & arry1(0) & ".gif") Then
vbwProfiler.vbwExecuteLine 1146
                                    Set SignalAttributes(Idx).Image = LoadPicture(SignalImageFilePath & arry1(0) & ".gif")
                                Else
vbwProfiler.vbwExecuteLine 1147 'B
vbwProfiler.vbwExecuteLine 1148
MsgBox "Flag " & arry1(0) & " doesnt exist"
                                End If
vbwProfiler.vbwExecuteLine 1149 'B
'vbwLine 1150:                            Case Is = "Group"
                            Case Is = IIf(vbwProfiler.vbwExecuteLine(1150), VBWPROFILER_EMPTY, _
        "Group")
vbwProfiler.vbwExecuteLine 1151
                                SignalAttributes(Idx).Group = arry(1)
vbwProfiler.vbwExecuteLine 1152
                                For i = 1 To UBound(GroupDefaults)
vbwProfiler.vbwExecuteLine 1153
                                    If GroupDefaults(i).Group = arry(1) Then
vbwProfiler.vbwExecuteLine 1154
                                        Exit For
                                    End If
vbwProfiler.vbwExecuteLine 1155 'B
vbwProfiler.vbwExecuteLine 1156
                                Next i
vbwProfiler.vbwExecuteLine 1157
                                If i <= UBound(GroupDefaults) Then
vbwProfiler.vbwExecuteLine 1158
                                    SignalAttributes(Idx).Flag.FixedCol = GroupDefaults(i).FixedCol
vbwProfiler.vbwExecuteLine 1159
                                    SignalAttributes(Idx).Flag.FixedRow = GroupDefaults(i).FixedRow
vbwProfiler.vbwExecuteLine 1160
                                    SignalAttributes(Idx).Flag.Queue = GroupDefaults(i).Queue
                                End If
vbwProfiler.vbwExecuteLine 1161 'B
'vbwLine 1162:                            Case Is = "Row"
                            Case Is = IIf(vbwProfiler.vbwExecuteLine(1162), VBWPROFILER_EMPTY, _
        "Row")
vbwProfiler.vbwExecuteLine 1163
                                    SignalAttributes(Idx).Flag.FixedRow = arry(1)

'Raise on load Testing only, can only do when This Signal is closed, as we do not know the position
'vbwLine 1164:                            Case Is = "Raised"
                            Case Is = IIf(vbwProfiler.vbwExecuteLine(1164), VBWPROFILER_EMPTY, _
        "Raised")
 'Initially display the flag for debugging the Position
vbwProfiler.vbwExecuteLine 1165
                                    If AtoBool(arry(1)) = True Then
vbwProfiler.vbwExecuteLine 1166
                                        Call frmMain.RaiseRequest(Idx)
                                    End If
vbwProfiler.vbwExecuteLine 1167 'B
'CommandAttributes
'vbwLine 1168:                            Case Is = "CommandVisible"
                            Case Is = IIf(vbwProfiler.vbwExecuteLine(1168), VBWPROFILER_EMPTY, _
        "CommandVisible")
vbwProfiler.vbwExecuteLine 1169
                                frmMain.Commands(Idx).Visible = AtoBool(arry(1))
'CommandFrame not currently used
'vbwLine 1170:                            Case Is = "CommandFrame"
                            Case Is = IIf(vbwProfiler.vbwExecuteLine(1170), VBWPROFILER_EMPTY, _
        "CommandFrame")
'                                    Set MyFrame = NametoFrame(arry(1))  'an Object
vbwProfiler.vbwExecuteLine 1171
                                    Select Case arry(1)
'vbwLine 1172:                                        Case Is = "Postponement", "Horn"
                                        Case Is = IIf(vbwProfiler.vbwExecuteLine(1172), VBWPROFILER_EMPTY, _
        "Postponement"), "Horn"
vbwProfiler.vbwExecuteLine 1173
                                        Set frmMain.Commands(Idx).Container = MyFrame
'                                    Set frmMain.Commands(Idx).Container = frmMain.fraPostponement
'                                    Set frmMain.Commands(Idx).Container = frmMain.fraHorn
vbwProfiler.vbwExecuteLine 1174
                                        frmMain.Commands(Idx).Top = 0
vbwProfiler.vbwExecuteLine 1175
                                        frmMain.Commands(Idx).Left = 0
vbwProfiler.vbwExecuteLine 1176
                                        frmMain.Commands(Idx).Width = 1700
'Position in middle of frame at the bottom
vbwProfiler.vbwExecuteLine 1177
                                        frmMain.Commands(Idx).Move _
                                        (MyFrame.Width _
                                        - frmMain.Commands(Idx).Width) / 2 _
                                        , MyFrame.Height _
                                       - frmMain.Commands(Idx).Height - 100
vbwProfiler.vbwExecuteLine 1178
                                       MyFont.Name = "Verdana"
vbwProfiler.vbwExecuteLine 1179
                                        MyFont.Size = 14
vbwProfiler.vbwExecuteLine 1180
                                        MyFont.Bold = True
vbwProfiler.vbwExecuteLine 1181
                                        Set frmMain.Commands(Idx).Font = MyFont
                                Case Else
vbwProfiler.vbwExecuteLine 1182 'B
vbwProfiler.vbwExecuteLine 1183
MsgBox "Command Frame " & arry(1) & " Container doesnt exist"

                                End Select
vbwProfiler.vbwExecuteLine 1184 'B
'frmMain.Commands(Idx).Visible = True
'Stop
                            Case Else
vbwProfiler.vbwExecuteLine 1185 'B
vbwProfiler.vbwExecuteLine 1186
MsgBox "Invalid " & arry(0) & " in Profile Section " & Section
                            End Select
vbwProfiler.vbwExecuteLine 1187 'B
                        Else
vbwProfiler.vbwExecuteLine 1188 'B
vbwProfiler.vbwExecuteLine 1189
MsgBox "No index " & arry(0) & " in Profile Section " & Section
                        End If  'Got a flag for this index
vbwProfiler.vbwExecuteLine 1190 'B
'vbwLine 1191:                    Case Is = "Event"
                    Case Is = IIf(vbwProfiler.vbwExecuteLine(1191), VBWPROFILER_EMPTY, _
        "Event")
vbwProfiler.vbwExecuteLine 1192
                        If IniNewEvent Is Nothing Then
vbwProfiler.vbwExecuteLine 1193
                            Set IniNewEvent = New clsEvent
'Enables is only used to enable command buttons
vbwProfiler.vbwExecuteLine 1194
                            IniNewEvent.Enabled = True  'Enable by default
                        End If
vbwProfiler.vbwExecuteLine 1195 'B
vbwProfiler.vbwExecuteLine 1196
                        With IniNewEvent
vbwProfiler.vbwExecuteLine 1197
                            Select Case arry(0)
'vbwLine 1198:                            Case Is = "Time"
                            Case Is = IIf(vbwProfiler.vbwExecuteLine(1198), VBWPROFILER_EMPTY, _
        "Time")
vbwProfiler.vbwExecuteLine 1199
                                .Second = arry(1) / Multiplier
'vbwLine 1200:                            Case Is = "Signal"
                            Case Is = IIf(vbwProfiler.vbwExecuteLine(1200), VBWPROFILER_EMPTY, _
        "Signal")
vbwProfiler.vbwExecuteLine 1201
                                .Signal = arry(1)
'vbwLine 1202:                            Case Is = "Raised"
                            Case Is = IIf(vbwProfiler.vbwExecuteLine(1202), VBWPROFILER_EMPTY, _
        "Raised")
vbwProfiler.vbwExecuteLine 1203
                                .Raised = AtoBool(arry(1))
'vbwLine 1204:                            Case Is = "Enabled" 'Set to true by default
                            Case Is = IIf(vbwProfiler.vbwExecuteLine(1204), VBWPROFILER_EMPTY, _
        "Enabled" )'Set to true by default
vbwProfiler.vbwExecuteLine 1205
                                .Enabled = AtoBool(arry(1))
'vbwLine 1206:                            Case Is = "Message"
                            Case Is = IIf(vbwProfiler.vbwExecuteLine(1206), VBWPROFILER_EMPTY, _
        "Message")
vbwProfiler.vbwExecuteLine 1207
                                If .Message <> "" Then
vbwProfiler.vbwExecuteLine 1208
                                     .Message = .Message & ", "
                                End If
vbwProfiler.vbwExecuteLine 1209 'B
vbwProfiler.vbwExecuteLine 1210
                                .Message = .Message & arry(1)
                            Case Else
vbwProfiler.vbwExecuteLine 1211 'B
vbwProfiler.vbwExecuteLine 1212
MsgBox "Invalid " & arry(0) & " in Profile Section " & Section
                            End Select
vbwProfiler.vbwExecuteLine 1213 'B
vbwProfiler.vbwExecuteLine 1214
                        End With
                    Case Else
vbwProfiler.vbwExecuteLine 1215 'B
vbwProfiler.vbwExecuteLine 1216
MsgBox "Invalid Section in Initialisation File"
                    End Select
vbwProfiler.vbwExecuteLine 1217 'B
                Else
vbwProfiler.vbwExecuteLine 1218 'B
vbwProfiler.vbwExecuteLine 1219
                    MsgBox "Line Outside section" & vbCrLf & nextline & vbCrLf, vbExclamation, "LoadProfile"
                End If
vbwProfiler.vbwExecuteLine 1220 'B
            End If
vbwProfiler.vbwExecuteLine 1221 'B
        End If
vbwProfiler.vbwExecuteLine 1222 'B
Skip_Line:
vbwProfiler.vbwExecuteLine 1223
    Loop
vbwProfiler.vbwExecuteLine 1224
    Close #Ch

'Check Command Button Signals have been defined
vbwProfiler.vbwExecuteLine 1225
    Call CommandIdx("Postpone")
vbwProfiler.vbwExecuteLine 1226
    Call CommandIdx("Horn Short")
vbwProfiler.vbwExecuteLine 1227
    Call CommandIdx("Recall")
vbwProfiler.vbwExecuteLine 1228
    Call CommandIdx("General Recall")
vbwProfiler.vbwExecuteLine 1229
    Call CommandIdx("Finish")

'    For i = 0 To UBound(frmMain.CmdQ)
'        If frmMain.CmdQ(i) = 0 And Commands(frmMain.CmdQ(i)).BackColor = vbCyan Then
'Stop
'        End If
'    Next i

vbwProfiler.vbwExecuteLine 1230
    frmMain.cmdEvents.Enabled = True
vbwProfiler.vbwExecuteLine 1231
    Loading = False
vbwProfiler.vbwExecuteLine 1232
    frmMain.RaceTimer.Enabled = True
'temp stop    frmMain.RaceTimer.Enabled = True

vbwProfiler.vbwProcOut 60
vbwProfiler.vbwExecuteLine 1233
End Function

Private Function FlagRaised(Idx As Long) As Boolean
vbwProfiler.vbwProcIn 61
Dim MyImage As Image
vbwProfiler.vbwExecuteLine 1234
    For Each MyImage In frmMain.Flags
vbwProfiler.vbwExecuteLine 1235
        If MyImage.Index = Idx Then
vbwProfiler.vbwExecuteLine 1236
            FlagRaised = True
vbwProfiler.vbwProcOut 61
vbwProfiler.vbwExecuteLine 1237
            Exit Function
        End If
vbwProfiler.vbwExecuteLine 1238 'B
vbwProfiler.vbwExecuteLine 1239
    Next MyImage
vbwProfiler.vbwProcOut 61
vbwProfiler.vbwExecuteLine 1240
End Function

Private Function CommandExists(Idx As Long) As Boolean
vbwProfiler.vbwProcIn 62
Dim MyCommand As CommandButton
vbwProfiler.vbwExecuteLine 1241
    For Each MyCommand In frmMain.Commands
vbwProfiler.vbwExecuteLine 1242
        If MyCommand.Index = Idx Then
vbwProfiler.vbwExecuteLine 1243
            CommandExists = True
vbwProfiler.vbwProcOut 62
vbwProfiler.vbwExecuteLine 1244
            Exit Function
        End If
vbwProfiler.vbwExecuteLine 1245 'B
vbwProfiler.vbwExecuteLine 1246
    Next MyCommand
vbwProfiler.vbwProcOut 62
vbwProfiler.vbwExecuteLine 1247
End Function

Private Function AtoBool(kb As String) As Boolean
vbwProfiler.vbwProcIn 63
vbwProfiler.vbwExecuteLine 1248
    If kb = "True" Then
vbwProfiler.vbwExecuteLine 1249
         AtoBool = True
    End If
vbwProfiler.vbwExecuteLine 1250 'B
vbwProfiler.vbwProcOut 63
vbwProfiler.vbwExecuteLine 1251
End Function

Public Function SignalIdx(SignalType As String, Optional SignalName As String) As Long
vbwProfiler.vbwProcIn 64
Dim i As Long
Dim kb As String

vbwProfiler.vbwExecuteLine 1252
    For i = 1 To UBound(SignalAttributes)
vbwProfiler.vbwExecuteLine 1253
        If SignalAttributes(i).Type = SignalType Then
vbwProfiler.vbwExecuteLine 1254
            If SignalName <> "" Then
vbwProfiler.vbwExecuteLine 1255
                If SignalAttributes(i).Name = SignalName Then
vbwProfiler.vbwExecuteLine 1256
                    SignalIdx = i
vbwProfiler.vbwProcOut 64
vbwProfiler.vbwExecuteLine 1257
                    Exit Function
                End If
vbwProfiler.vbwExecuteLine 1258 'B
            Else
vbwProfiler.vbwExecuteLine 1259 'B
vbwProfiler.vbwExecuteLine 1260
                SignalIdx = i
vbwProfiler.vbwProcOut 64
vbwProfiler.vbwExecuteLine 1261
                Exit Function
            End If
vbwProfiler.vbwExecuteLine 1262 'B
        End If
vbwProfiler.vbwExecuteLine 1263 'B
vbwProfiler.vbwExecuteLine 1264
    Next i
vbwProfiler.vbwExecuteLine 1265
    kb = "Signal Type " & SignalType
vbwProfiler.vbwExecuteLine 1266
    If SignalName <> "" Then
vbwProfiler.vbwExecuteLine 1267
        kb = kb & ", Signal Name " & SignalName
    End If
vbwProfiler.vbwExecuteLine 1268 'B
vbwProfiler.vbwExecuteLine 1269
    kb = kb & " not found"
vbwProfiler.vbwExecuteLine 1270
MsgBox kb, vbExclamation, "SignalIdx"
vbwProfiler.vbwProcOut 64
vbwProfiler.vbwExecuteLine 1271
End Function

Public Function CommandIdx(CommandName As String) As Long
vbwProfiler.vbwProcIn 65
Dim kb As String
Dim MyCommand As CommandButton

vbwProfiler.vbwExecuteLine 1272
    For Each MyCommand In frmMain.Commands
vbwProfiler.vbwExecuteLine 1273
        If MyCommand.Caption = CommandName Then
vbwProfiler.vbwExecuteLine 1274
            CommandIdx = MyCommand.Index
vbwProfiler.vbwProcOut 65
vbwProfiler.vbwExecuteLine 1275
            Exit Function
        End If
vbwProfiler.vbwExecuteLine 1276 'B
vbwProfiler.vbwExecuteLine 1277
    Next MyCommand
vbwProfiler.vbwExecuteLine 1278
MsgBox "Command Button " & CommandName & " not found", vbExclamation, "CommandIdx"
vbwProfiler.vbwProcOut 65
vbwProfiler.vbwExecuteLine 1279
End Function

'Check to see if Timer(index) exists, there seems to be no other way to check
'other than trying to access the index
Public Function TimerExists(Idx As Long) As Boolean
vbwProfiler.vbwProcIn 66
vbwProfiler.vbwExecuteLine 1280
    On Error GoTo NoTimer
vbwProfiler.vbwExecuteLine 1281
    If frmMain.SignalTimer(Idx).Index Then
vbwProfiler.vbwExecuteLine 1282
        TimerExists = True
    End If
vbwProfiler.vbwExecuteLine 1283 'B
NoTimer:
vbwProfiler.vbwProcOut 66
vbwProfiler.vbwExecuteLine 1284
End Function

Private Function UnloadSignalTimers()
vbwProfiler.vbwProcIn 67
Dim oSignalTimer As Timer
Dim Idx As Integer
vbwProfiler.vbwExecuteLine 1285
    For Each oSignalTimer In frmMain.SignalTimer
vbwProfiler.vbwExecuteLine 1286
        Idx = oSignalTimer.Index
'        Set oSignalTimer = Nothing
vbwProfiler.vbwExecuteLine 1287
        If Idx > 0 Then
'            Unload frmMain.SignalTimer(Idx)
vbwProfiler.vbwExecuteLine 1288
            Unload oSignalTimer
        End If
vbwProfiler.vbwExecuteLine 1289 'B
vbwProfiler.vbwExecuteLine 1290
    Next
vbwProfiler.vbwProcOut 67
vbwProfiler.vbwExecuteLine 1291
End Function

Public Function HasIndex(ControlArray As Object, ByVal Index As Integer) As Boolean
vbwProfiler.vbwProcIn 68
vbwProfiler.vbwExecuteLine 1292
    HasIndex = (VarType(ControlArray(Index)) <> vbObject)
vbwProfiler.vbwProcOut 68
vbwProfiler.vbwExecuteLine 1293
End Function

Public Function NameFromFullPath(FullPath As String, Optional Delimiter As String, Optional RemoveRollover As Boolean) As String
'Input: Name/Full Path of a file
'Returns: Name of file
vbwProfiler.vbwProcIn 69

    Dim sPath As String
    Dim sList() As String
    Dim sAns As String
    Dim iArrayLen As Integer
    Dim i As Integer
    Dim j As Integer
    Dim kb As String

'MsgBox FullPath
vbwProfiler.vbwExecuteLine 1294
    If Delimiter = "" Then
vbwProfiler.vbwExecuteLine 1295
         Delimiter = "\"
    End If
vbwProfiler.vbwExecuteLine 1296 'B
vbwProfiler.vbwExecuteLine 1297
    If Len(FullPath) = 0 Then
vbwProfiler.vbwProcOut 69
vbwProfiler.vbwExecuteLine 1298
         Exit Function
    End If
vbwProfiler.vbwExecuteLine 1299 'B
vbwProfiler.vbwExecuteLine 1300
    sList = Split(FullPath, Delimiter)
vbwProfiler.vbwExecuteLine 1301
    iArrayLen = UBound(sList)
vbwProfiler.vbwExecuteLine 1302
    sAns = IIf(iArrayLen = 0, "", sList(iArrayLen))
'only filename
'MsgBox FullPath
vbwProfiler.vbwExecuteLine 1303
    If sAns = "" And iArrayLen = 0 Then
vbwProfiler.vbwExecuteLine 1304
         sAns = FullPath
    End If
vbwProfiler.vbwExecuteLine 1305 'B
vbwProfiler.vbwExecuteLine 1306
    If RemoveRollover And sAns <> "" Then
vbwProfiler.vbwExecuteLine 1307
        j = InStr(sAns, ".") 'get the first dot
vbwProfiler.vbwExecuteLine 1308
        If j = 0 Then 'no dot so all the string
vbwProfiler.vbwExecuteLine 1309
             j = Len(sAns)
        End If
vbwProfiler.vbwExecuteLine 1310 'B
vbwProfiler.vbwExecuteLine 1311
        i = InStrRev(Left$(sAns, j), "_")
vbwProfiler.vbwExecuteLine 1312
        If j = i + 9 Then 'msu be _yyyymmdd.
vbwProfiler.vbwExecuteLine 1313
            If IsNumeric(Mid$(sAns, i + 1, 8)) Then
vbwProfiler.vbwExecuteLine 1314
                sAns = Replace(sAns, Mid$(sAns, i, 9), "")
            End If
vbwProfiler.vbwExecuteLine 1315 'B
        End If
vbwProfiler.vbwExecuteLine 1316 'B
    End If
vbwProfiler.vbwExecuteLine 1317 'B

vbwProfiler.vbwExecuteLine 1318
    NameFromFullPath = sAns

vbwProfiler.vbwProcOut 69
vbwProfiler.vbwExecuteLine 1319
End Function


' Return True if a file exists
Public Function FileExists(FileName As String) As Boolean
vbwProfiler.vbwProcIn 70
vbwProfiler.vbwExecuteLine 1320
    FileExists = False
'MsgBox FileName & ":" & GetAttr(FileName)
vbwProfiler.vbwExecuteLine 1321
    On Error GoTo ErrorHandler
vbwProfiler.vbwExecuteLine 1322
    If NameFromFullPath(FileName) <> "" Then  'directory
'does file exists
vbwProfiler.vbwExecuteLine 1323
        If (GetAttr(FileName) And vbNormal) = vbNormal Then
vbwProfiler.vbwExecuteLine 1324
             FileExists = True
        End If
vbwProfiler.vbwExecuteLine 1325 'B
    End If
vbwProfiler.vbwExecuteLine 1326 'B
'MsgBox Filename & vbCrLf & FileExists
ErrorHandler:
    ' if an error occurs, this function returns False
vbwProfiler.vbwProcOut 70
vbwProfiler.vbwExecuteLine 1327
End Function

Public Function CreateLink(ByRef Idx As Long, Link As defLink)
vbwProfiler.vbwProcIn 71

vbwProfiler.vbwExecuteLine 1328
    With SignalAttributes(Idx)
vbwProfiler.vbwExecuteLine 1329
        If IsArrayInitialised(.Links) = False Then
vbwProfiler.vbwExecuteLine 1330
            ReDim .Links(0)
        Else
vbwProfiler.vbwExecuteLine 1331 'B
vbwProfiler.vbwExecuteLine 1332
            ReDim Preserve .Links(UBound(.Links) + 1)
        End If
vbwProfiler.vbwExecuteLine 1333 'B
vbwProfiler.vbwExecuteLine 1334
        .Links(UBound(.Links)) = Link
vbwProfiler.vbwExecuteLine 1335
    End With
vbwProfiler.vbwProcOut 71
vbwProfiler.vbwExecuteLine 1336
End Function

Public Function IsArrayInitialised(ByRef arr() As defLink) As Boolean
vbwProfiler.vbwProcIn 72
Dim Temp As Long
'Return True if array is initalized
vbwProfiler.vbwExecuteLine 1337
    On Error GoTo errHandler 'Raise error if directory doesnot exist
vbwProfiler.vbwExecuteLine 1338
    Temp = UBound(arr)
  'Reach this point only if arr is initalized i.e. no error occured
vbwProfiler.vbwExecuteLine 1339
    If Temp > -1 Then 'UBound is greater then -1
vbwProfiler.vbwExecuteLine 1340
         IsArrayInitialised = True
    End If
vbwProfiler.vbwExecuteLine 1341 'B
vbwProfiler.vbwProcOut 72
vbwProfiler.vbwExecuteLine 1342
Exit Function
errHandler:
  'if an error occurs, this function returns False. i.e. array not initialized
vbwProfiler.vbwProcOut 72
vbwProfiler.vbwExecuteLine 1343
End Function



