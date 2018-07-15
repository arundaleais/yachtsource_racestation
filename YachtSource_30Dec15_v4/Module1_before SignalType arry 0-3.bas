Attribute VB_Name = "modMain"
Option Explicit

Private Type defSignalType
    LinkedFlag As Long
    TTL As Long 'Time this flag is displayed in Millisecs
                'It will be off for the same Period (if more than 1 cycle)
    CyclesRequired As Long  'No of On cycles
End Type
Private Type defSignalAttribute
    index As Long
'These are defined againg as they are used once the timer is
'running - they are loaded from the SignalType
    TTL As Long 'Time this flag is displayed in Millisecs
                'It will be off for the same Period (if more than 1 cycle)
    CyclesRequired As Long  'No of On cycles
    CyclesCompleted As Long 'timer must be unique
    SignalType(0 To 1) As defSignalType '0=Going on, 1=going off
End Type

Public SignalAttributes() As defSignalAttribute

Public myProfile As clsProfile
Public ElapsedTime As Long

Sub main()
'    Action.Load (Environ("AllUsersProfile") & "\Application Data\Arundale\RacingSignals\Sequences\" & "ScarboroughMultiple.csv")
    Load frmMain
    frmMain.Show
    Call LoadElapsedSeconds
    frmMain.RaceTimer.Enabled = True
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
Dim LastEvent As clsEvent
Dim FirstEvent As clsEvent

    If myProfile.IsEventDue(ElapsedTime) = False Then
'If ElapsedTime = 0 Then Stop
        Exit Function
    End If
            
    DoTimerEvents = True    'Start sequence has started
    For Each MyEvent In myProfile
'If MyEvent.Signal = 0 Then Stop
        If MyEvent.Second = ElapsedTime Then
Debug.Print MyEvent.Second & " " & MyEvent.Signal & " " & MyEvent.State
'MakeSignals will generated any linked signals
'And use the SignalTimer if required
            Call MakeSignals(MyEvent.Signal, MyEvent.State)
        End If
    Next MyEvent
    
    Set FirstEvent = myProfile.FirstEvent
    Set LastEvent = myProfile.LastEvent

'If Timer Events have started and not finished you cannot Postpone
    Select Case ElapsedTime
    Case Is >= LastEvent.Second
            frmMain.cmdPostpone.Enabled = True
    Case Is >= FirstEvent.Second
            frmMain.cmdPostpone.Enabled = False
    Case Else
            frmMain.cmdPostpone.Enabled = True
    End Select
    Set FirstEvent = Nothing
    Set LastEvent = Nothing
End Function

'Using Visible because On is a reserved word
Public Function MakeSignals(Signal As Long, Visible As Boolean)
    If frmMain.Flags(Signal).Visible <> Visible Then
'State of Signal requires changing
        If Visible = True Then
'Turn ON Signal
            frmMain.Flags(Signal).Visible = True
'Load the SignalType to be used by the timer
            SignalAttributes(Signal).TTL = _
            SignalAttributes(Signal).SignalType(0).TTL
            SignalAttributes(Signal).CyclesRequired = _
            SignalAttributes(Signal).SignalType(0).CyclesRequired
            SignalAttributes(Signal).CyclesCompleted = 0
'Check if any linked Signals
            If SignalAttributes(Signal).SignalType(0).LinkedFlag <> 0 Then
'This is re-entrant into this function
'Stop
                Call MakeSignals(SignalAttributes(Signal).SignalType(0).LinkedFlag, Visible)
            End If
'Check if any linked ON Signals
            If SignalAttributes(Signal).SignalType(0).LinkedFlag <> 0 Then
'This is re-entrant into this function
'Stop
'                Call MakeSignals(SignalAttributes(Signal).SignalType(0).LinkedFlag, Visible)
'??? if turning ON, does linked (eg horn) signal require turning off or on ???
                Call MakeSignals(SignalAttributes(Signal).SignalType(1).LinkedFlag, True)
            End If
'end of ON
        Else
'Turn OFF signal
            frmMain.Flags(Signal).Visible = False
            SignalAttributes(Signal).TTL = _
            SignalAttributes(Signal).SignalType(1).TTL
            SignalAttributes(Signal).CyclesRequired = _
            SignalAttributes(Signal).SignalType(1).CyclesRequired

'Check if any linked OFF Signals
            If SignalAttributes(Signal).SignalType(1).LinkedFlag <> 0 Then
'This is re-entrant into this function
'Stop
'                Call MakeSignals(SignalAttributes(Signal).SignalType(1).LinkedFlag, Visible)
'??? if turning off, does linked (eg horn) signal require turning off or on ???
                Call MakeSignals(SignalAttributes(Signal).SignalType(1).LinkedFlag, True)
            End If
'end of OFF
        End If

'If SignalTimer is required for this signal
'And it is not already running, start the timer
        If frmMain.SignalTimer(Signal).Enabled = False Then
            If SignalAttributes(Signal).TTL <> 0 Then
                frmMain.SignalTimer(Signal).Enabled = True
            End If
        Else
'If the timer is running, don't re-enable the timer
        End If

    Else
'State of Signal has not changed
    End If

End Function

Public Function LoadElapsedSeconds()
Dim i As Long
Dim Secs As Long
Dim Ch As Long
Dim nextline As String
Dim arry() As String
Dim ActionsFileName
Dim MyImage As Image
Dim j As Long
Dim arType() As String  'SignalType array

'Extract the Signal attributes from the Tag field on the Signal Image Control
'these are any defaults that I have set up within the image Tag
'These could be overridden by the .ini file
                                        'Flags(0) is not used
    ReDim SignalAttributes(1 To frmMain.Flags.Count - 1)
    For Each MyImage In frmMain.Flags
        If MyImage.index > 0 Then   'Flags(0) is not used
'Create a timer for each Signal (even if we dont use it)
            Load frmMain.SignalTimer(MyImage.index)

'It helps debugging to keep the index so you can see it in the Watch window
            SignalAttributes(MyImage.index).index = MyImage.index
            If MyImage.Tag <> "" Then
                arType = Split(MyImage.Tag, "/")
                For j = 0 To UBound(arType)
'Tag format is LinkedFlag,TTL,CyclesRequired
                    arry = Split(arType(j), ",")
'Could be just first Item ie "5"
                    For i = 0 To UBound(arry)
                        arry(i) = NulToZero(arry(i))
                        Select Case i
                        Case Is = 0
                            SignalAttributes(MyImage.index).SignalType(j).LinkedFlag = arry(0)
                        Case Is = 1
                            SignalAttributes(MyImage.index).SignalType(j).TTL = arry(1)
                        Case Is = 2
                            SignalAttributes(MyImage.index).SignalType(j).CyclesRequired = arry(2)
                        End Select
                    Next i
                Next j
            End If
        End If
    Next MyImage


'Start a fresh profile
    Set myProfile = Nothing
    Set myProfile = New clsProfile
    myProfile.Name = "Profile1"
    ActionsFileName = Environ("AllUsersProfile") & "\Application Data\Arundale\RacingSignals\Sequences\" & "ScarboroughMultiple.csv"
    Ch = FreeFile
    Open ActionsFileName For Input As #Ch
    Do Until EOF(Ch)
        Line Input #Ch, nextline
        arry() = Split(nextline, ",")
        If UBound(arry) = 4 Then
'create the required action
'0 is the class (Not required as we have the class flag)
            myProfile.NewEvent CLng(arry(0)), 0, CLng(arry(1)), CLng(arry(2))
            i = i + 1
        End If
    Loop
    Close #Ch
    frmMain.cmdPostpone.Enabled = True
'This is how to load a Signal(Flag or Sound)
'Load frmMain.Flags(5)  'Create the Control Array
'frmMain.Flags(1).Picture = LoadPicture(Environ("AllUsersProfile") & "\Application Data\Arundale\RacingSignals\SignalImages\" & "TrafficLightWhite.gif")
    

Debug.Print "LoadEvents " & i

End Function


