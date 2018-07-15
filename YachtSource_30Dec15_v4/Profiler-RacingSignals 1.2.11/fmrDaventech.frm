VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Begin VB.Form frmDaventech 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer ReconnectTimer 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   120
      Top             =   840
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   240
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmDaventech"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DataSnd As String   'Keep a copy of the CommandString in case we need to retry
Dim Controller As Long
Dim Command As Long
Dim OnTime As Long
Private Closing As Boolean  'Stops reconnect timer
Private ReplyWait As Boolean    'Must wait for a reply after data is sent

Private Sub Command1_Click()
vbwProfiler.vbwProcIn 126
Dim State As Long
Dim arry(2) As String
Dim ControlString As String
'Debug.Print "======="
vbwProfiler.vbwExecuteLine 2364
    If Command = 0 Then
vbwProfiler.vbwExecuteLine 2365
        Call OpenAndSend("")    'reset
vbwProfiler.vbwExecuteLine 2366
        Command = 32
vbwProfiler.vbwProcOut 126
vbwProfiler.vbwExecuteLine 2367
        Exit Sub
    End If
vbwProfiler.vbwExecuteLine 2368 'B
vbwProfiler.vbwExecuteLine 2369
    If Controller = 0 Then
vbwProfiler.vbwExecuteLine 2370
         Controller = 1
    End If
vbwProfiler.vbwExecuteLine 2371 'B
vbwProfiler.vbwExecuteLine 2372
    If Controller = 9 Then
vbwProfiler.vbwExecuteLine 2373
        Controller = 1
vbwProfiler.vbwExecuteLine 2374
        If Command = 32 Then
vbwProfiler.vbwExecuteLine 2375
            Command = 33
        Else
vbwProfiler.vbwExecuteLine 2376 'B
vbwProfiler.vbwExecuteLine 2377
            Command = 32
        End If
vbwProfiler.vbwExecuteLine 2378 'B
    End If
vbwProfiler.vbwExecuteLine 2379 'B
vbwProfiler.vbwExecuteLine 2380
    arry(0) = Command
vbwProfiler.vbwExecuteLine 2381
    arry(1) = Controller
vbwProfiler.vbwExecuteLine 2382
    arry(2) = OnTime
vbwProfiler.vbwExecuteLine 2383
    ControlString = Join(arry, ",")

vbwProfiler.vbwExecuteLine 2384
    State = OpenAndSend(ControlString)
vbwProfiler.vbwExecuteLine 2385
    Controller = Controller + 1
vbwProfiler.vbwProcOut 126
vbwProfiler.vbwExecuteLine 2386
End Sub

'Returns the state of the socket
Public Function OpenAndSend(ControlString As String) As Long
vbwProfiler.vbwProcIn 127
vbwProfiler.vbwExecuteLine 2387
    If ControlString = "" Then
    'Reset all off
vbwProfiler.vbwExecuteLine 2388
        DataSnd = "35,0,0"
    Else
vbwProfiler.vbwExecuteLine 2389 'B
vbwProfiler.vbwExecuteLine 2390
        DataSnd = ControlString
    End If
vbwProfiler.vbwExecuteLine 2391 'B
vbwProfiler.vbwExecuteLine 2392
    With Winsock1
vbwProfiler.vbwExecuteLine 2393
        Select Case .State
'vbwLine 2394:        Case Is = sckClosed
        Case Is = IIf(vbwProfiler.vbwExecuteLine(2394), VBWPROFILER_EMPTY, _
        sckClosed)
vbwProfiler.vbwExecuteLine 2395
            Call CreateWinsock
'vbwLine 2396:        Case Is = sckConnected
        Case Is = IIf(vbwProfiler.vbwExecuteLine(2396), VBWPROFILER_EMPTY, _
        sckConnected)
vbwProfiler.vbwExecuteLine 2397
            Call WinsockOutput
        End Select
vbwProfiler.vbwExecuteLine 2398 'B
vbwProfiler.vbwExecuteLine 2399
    OpenAndSend = .State
vbwProfiler.vbwExecuteLine 2400
    End With
vbwProfiler.vbwProcOut 127
vbwProfiler.vbwExecuteLine 2401
End Function

Private Sub CreateWinsock()
vbwProfiler.vbwProcIn 128
Dim lWaitUntil As Long
vbwProfiler.vbwExecuteLine 2402
    With Winsock1
vbwProfiler.vbwExecuteLine 2403
        If .State = 0 Then
vbwProfiler.vbwExecuteLine 2404
            .Protocol = sckTCPProtocol
vbwProfiler.vbwExecuteLine 2405
            If .RemoteHost = "eth008" Then
vbwProfiler.vbwExecuteLine 2406
                .RemoteHost = "192.168.0.200"
            Else
vbwProfiler.vbwExecuteLine 2407 'B
vbwProfiler.vbwExecuteLine 2408
                .RemoteHost = "eth008"
            End If
vbwProfiler.vbwExecuteLine 2409 'B
'Debug.Print .RemoteHost
vbwProfiler.vbwExecuteLine 2410
            .RemotePort = 17494
vbwProfiler.vbwExecuteLine 2411
            lWaitUntil = TimeToQuit(5)
vbwProfiler.vbwExecuteLine 2412
            .Connect
'vbwLine 2413:            Do Until Winsock1.State = sckConnected Or Timer > lWaitUntil
            Do Until vbwProfiler.vbwExecuteLine(2413) Or Winsock1.State = sckConnected Or Timer > lWaitUntil
vbwProfiler.vbwExecuteLine 2414
                 DoEvents
vbwProfiler.vbwExecuteLine 2415
            Loop
vbwProfiler.vbwExecuteLine 2416
            If Winsock1.State = sckConnected Then
'Debug.Print "Connection Successful"
                Else
vbwProfiler.vbwExecuteLine 2417 'B
'Debug.Print "Connection TimedOut"
'Call DisplayWinsock
            End If
vbwProfiler.vbwExecuteLine 2418 'B
        End If
vbwProfiler.vbwExecuteLine 2419 'B
vbwProfiler.vbwExecuteLine 2420
    End With
'Terminate ReconnectTimer when unloading frmDaventech
vbwProfiler.vbwExecuteLine 2421
    ReconnectTimer.Enabled = Not Closing
vbwProfiler.vbwProcOut 128
vbwProfiler.vbwExecuteLine 2422
End Sub

Public Sub CloseWinsock()
vbwProfiler.vbwProcIn 129
Dim lWaitUntil As Long

vbwProfiler.vbwExecuteLine 2423
    On Error GoTo Winsock_err
vbwProfiler.vbwExecuteLine 2424
    With Winsock1
vbwProfiler.vbwExecuteLine 2425
        lWaitUntil = TimeToQuit(5)
vbwProfiler.vbwExecuteLine 2426
        .Close
'vbwLine 2427:        Do Until Winsock1.State = sckClosed Or Timer > lWaitUntil
        Do Until vbwProfiler.vbwExecuteLine(2427) Or Winsock1.State = sckClosed Or Timer > lWaitUntil
vbwProfiler.vbwExecuteLine 2428
            DoEvents
vbwProfiler.vbwExecuteLine 2429
        Loop
vbwProfiler.vbwExecuteLine 2430
        If Winsock1.State = sckClosed Then
'Debug.Print "Connection Closed"
        Else
vbwProfiler.vbwExecuteLine 2431 'B
'Debug.Print "Close TimedOut"
        End If
vbwProfiler.vbwExecuteLine 2432 'B
vbwProfiler.vbwExecuteLine 2433
    End With
vbwProfiler.vbwProcOut 129
vbwProfiler.vbwExecuteLine 2434
Exit Sub
Winsock_err:
vbwProfiler.vbwExecuteLine 2435
    On Error GoTo 0
vbwProfiler.vbwExecuteLine 2436
    MsgBox "CloseWinsock Error " & Err.Number & " " & Err.Description & vbCrLf, , "Close Winsock"
vbwProfiler.vbwProcOut 129
vbwProfiler.vbwExecuteLine 2437
End Sub

Private Sub Form_Load()
vbwProfiler.vbwProcIn 130
vbwProfiler.vbwExecuteLine 2438
    Call CreateWinsock
vbwProfiler.vbwProcOut 130
vbwProfiler.vbwExecuteLine 2439
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
vbwProfiler.vbwProcIn 131

vbwProfiler.vbwExecuteLine 2440
Debug.Print "frmDaventech.unload " & Winsock1.State
vbwProfiler.vbwExecuteLine 2441
    Closing = True  'stop reconnecttimer restarting
vbwProfiler.vbwExecuteLine 2442
    ReconnectTimer.Enabled = False
vbwProfiler.vbwExecuteLine 2443
    If Winsock1.State = sckConnected Then
vbwProfiler.vbwExecuteLine 2444
        DataSnd = "35,0,0"
vbwProfiler.vbwExecuteLine 2445
        Call WinsockOutput
vbwProfiler.vbwExecuteLine 2446
        Call CloseWinsock
    End If
vbwProfiler.vbwExecuteLine 2447 'B
vbwProfiler.vbwProcOut 131
vbwProfiler.vbwExecuteLine 2448
End Sub

Private Sub ReconnectTimer_Timer()
'Debug.Print "Reconnect " & Winsock1.State
vbwProfiler.vbwProcIn 132

vbwProfiler.vbwExecuteLine 2449
    With Winsock1
vbwProfiler.vbwExecuteLine 2450
        Select Case .State
'vbwLine 2451:        Case Is = sckClosed, sckClosing ', sckError
        Case Is = IIf(vbwProfiler.vbwExecuteLine(2451), VBWPROFILER_EMPTY, _
        sckClosed), sckClosing ', sckError
vbwProfiler.vbwExecuteLine 2452
            Call CreateWinsock
'vbwLine 2453:        Case Is = sckConnecting, sckError 'Client
        Case Is = IIf(vbwProfiler.vbwExecuteLine(2453), VBWPROFILER_EMPTY, _
        sckConnecting), sckError 'Client
vbwProfiler.vbwExecuteLine 2454
            Call CloseWinsock
vbwProfiler.vbwExecuteLine 2455
            Call CreateWinsock
'vbwLine 2456:        Case Is = sckConnectionPending 'Server Only
        Case Is = IIf(vbwProfiler.vbwExecuteLine(2456), VBWPROFILER_EMPTY, _
        sckConnectionPending )'Server Only
'            Call CloseHandler(Idx)
'            Call OpenHandler(Idx)
'vbwLine 2457:        Case Is = sckConnected
        Case Is = IIf(vbwProfiler.vbwExecuteLine(2457), VBWPROFILER_EMPTY, _
        sckConnected)
'            ReconnectTimer.Enabled = False
        End Select
vbwProfiler.vbwExecuteLine 2458 'B
vbwProfiler.vbwExecuteLine 2459
        If .State = sckConnected Then
vbwProfiler.vbwExecuteLine 2460
            frmMain.StatusBar1.Panels(2).Picture = LoadPicture(SignalImageFilePath & "connected.gif")
'            Call SendControllers
        Else
vbwProfiler.vbwExecuteLine 2461 'B
vbwProfiler.vbwExecuteLine 2462
            frmMain.StatusBar1.Panels(2).Picture = LoadPicture(SignalImageFilePath & "notconnected.gif")
        End If
vbwProfiler.vbwExecuteLine 2463 'B
vbwProfiler.vbwExecuteLine 2464
    End With

vbwProfiler.vbwProcOut 132
vbwProfiler.vbwExecuteLine 2465
End Sub

Private Sub Winsock1_Connect()
'Debug.Print "Connect " & Winsock1.State
vbwProfiler.vbwProcIn 133
vbwProfiler.vbwExecuteLine 2466
    If Winsock1.State = sckConnected Then
'        Call WinsockOutput
    End If
vbwProfiler.vbwExecuteLine 2467 'B
vbwProfiler.vbwProcOut 133
vbwProfiler.vbwExecuteLine 2468
End Sub

Sub WinsockOutput()
vbwProfiler.vbwProcIn 134
Dim lWaitUntil As Long
Dim b() As Byte
Dim arry() As String
Dim i As Long

'the Winsock Control element may not have had time to be created
'by the time it starts
'the Port & Socket may have been closed by the user while there'
'were unsent Sentences in the buffer
'Debug.Print DataSnd
vbwProfiler.vbwExecuteLine 2469
    With Winsock1
vbwProfiler.vbwExecuteLine 2470
        If .State = sckConnected Then
vbwProfiler.vbwExecuteLine 2471
            If DataSnd <> "" Then
vbwProfiler.vbwExecuteLine 2472
                arry = Split(DataSnd, ",")
vbwProfiler.vbwExecuteLine 2473
                ReDim b(UBound(arry))

vbwProfiler.vbwExecuteLine 2474
                For i = 0 To UBound(arry)
vbwProfiler.vbwExecuteLine 2475
                    If IsNumeric(arry(i)) Then
vbwProfiler.vbwExecuteLine 2476
                        b(i) = arry(i)
                    End If
vbwProfiler.vbwExecuteLine 2477 'B
vbwProfiler.vbwExecuteLine 2478
                Next i
'Debug.Print "DataSnd=" & DataSnd
vbwProfiler.vbwExecuteLine 2479
                ReplyWait = True
vbwProfiler.vbwExecuteLine 2480
                lWaitUntil = TimeToQuit(5)
vbwProfiler.vbwExecuteLine 2481
                .SendData b
'vbwLine 2482:                Do Until ReplyWait = False Or Timer > lWaitUntil
                Do Until vbwProfiler.vbwExecuteLine(2482) Or ReplyWait = False Or Timer > lWaitUntil
'Wait until a reply is received
vbwProfiler.vbwExecuteLine 2483
                     DoEvents
vbwProfiler.vbwExecuteLine 2484
                Loop


            End If
vbwProfiler.vbwExecuteLine 2485 'B
        End If
vbwProfiler.vbwExecuteLine 2486 'B
vbwProfiler.vbwExecuteLine 2487
    End With
vbwProfiler.vbwProcOut 134
vbwProfiler.vbwExecuteLine 2488
Exit Sub
SendData_err:
vbwProfiler.vbwExecuteLine 2489
    MsgBox "Send Data Error " & Str(Err.Number) & " " & Err.Description & vbCrLf
vbwProfiler.vbwProcOut 134
vbwProfiler.vbwExecuteLine 2490
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
vbwProfiler.vbwProcIn 135
Dim DataRcv As String
Dim kb As String
Dim i As Long

vbwProfiler.vbwExecuteLine 2491
    On Error GoTo DataArrival_err
vbwProfiler.vbwExecuteLine 2492
    With Winsock1
vbwProfiler.vbwExecuteLine 2493
        .GetData DataRcv, vbString
vbwProfiler.vbwExecuteLine 2494
    End With
vbwProfiler.vbwExecuteLine 2495
    For i = 1 To bytesTotal
vbwProfiler.vbwExecuteLine 2496
        kb = kb & Asc(Mid$(DataRcv, i, 1)) & " "
vbwProfiler.vbwExecuteLine 2497
    Next i
'Debug.Print "Reply=" & kb
vbwProfiler.vbwExecuteLine 2498
    ReplyWait = False
vbwProfiler.vbwProcOut 135
vbwProfiler.vbwExecuteLine 2499
Exit Sub

DataArrival_err:
vbwProfiler.vbwExecuteLine 2500
    Select Case Err.Number
'vbwLine 2501:    Case Is = sckBadState
    Case Is = IIf(vbwProfiler.vbwExecuteLine(2501), VBWPROFILER_EMPTY, _
        sckBadState)
'vbwLine 2502:    Case Is = sckMsgTooBig
    Case Is = IIf(vbwProfiler.vbwExecuteLine(2502), VBWPROFILER_EMPTY, _
        sckMsgTooBig)
'vbwLine 2503:    Case Is = sckConnectionReset
    Case Is = IIf(vbwProfiler.vbwExecuteLine(2503), VBWPROFILER_EMPTY, _
        sckConnectionReset)
    Case Else
vbwProfiler.vbwExecuteLine 2504 'B
vbwProfiler.vbwExecuteLine 2505
        MsgBox "UDP/TCP DataArrival Error " & Str(Err.Number) & " " & Err.Description
    End Select
vbwProfiler.vbwExecuteLine 2506 'B
vbwProfiler.vbwProcOut 135
vbwProfiler.vbwExecuteLine 2507
End Sub

'If we cant connect - just close the connection (it will be retried by the reconnect timer)
Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'frmMain.StatusBar1.Panels(1).Text = Description
vbwProfiler.vbwProcIn 136
Dim kb As String

vbwProfiler.vbwExecuteLine 2508
    kb = "Error:" & Number & vbCrLf & "Message:" & Description
'Call DisplayWinsock(kb)
vbwProfiler.vbwExecuteLine 2509
    Select Case Number
'vbwLine 2510:    Case Is = 11001 'No Host Found
    Case Is = IIf(vbwProfiler.vbwExecuteLine(2510), VBWPROFILER_EMPTY, _
        11001 )'No Host Found
vbwProfiler.vbwExecuteLine 2511
        Call CloseWinsock
'        Call CreateWinsock
'vbwLine 2512:    Case Is = 10065 'No route to host
    Case Is = IIf(vbwProfiler.vbwExecuteLine(2512), VBWPROFILER_EMPTY, _
        10065 )'No route to host
vbwProfiler.vbwExecuteLine 2513
        Call CloseWinsock
'        Call CreateWinsock
'Stop
    Case Else
vbwProfiler.vbwExecuteLine 2514 'B
vbwProfiler.vbwExecuteLine 2515
        Call CloseWinsock
'Stop
    End Select
vbwProfiler.vbwExecuteLine 2516 'B
vbwProfiler.vbwProcOut 136
vbwProfiler.vbwExecuteLine 2517
End Sub

Public Function TimeToQuit(TimeToWait As Long) As Long
'http://www.freevbcode.com/ShowCode.asp?ID=1977
'PURPOSE:  Returns a TimeOut value, in metric
'Seconds from Midnight (i.e., return value of Timer function)
'taking into account that Midnight may occur within the elapsed time
vbwProfiler.vbwProcIn 137

'PARAMETER: TimeToWait: Number of Seconds to Wait

'RETURN VALUE:  When to TimeOut, in Seconds From Midnight

'EXAMPLE: Implements a 30 second timeout before giving
'up on a winsock Connection

'Dim lWaitUntil as Long
'lWaitUntil = TimeToQuit(30)
'Winsock1.Connect
'Do Until Winsock1.State = sckConnected or Timer > _
'   lWaitUntil
'     DoEvents
'Loop
'If Winsock1.State = sckConnected Then
'     MsgBox "Connection Successful"
'Else
'    MsgBox "Connection TimedOut
'End If
'*************************************************************
Dim lStart As Long
Dim lTimeToQuit As Long
Dim lTimeToWait As Long

vbwProfiler.vbwExecuteLine 2518
lStart = Timer
vbwProfiler.vbwExecuteLine 2519
lTimeToWait = TimeToWait

vbwProfiler.vbwExecuteLine 2520
If lStart + TimeToWait < 86400 Then
vbwProfiler.vbwExecuteLine 2521
        lTimeToQuit = lStart + lTimeToWait
    Else
vbwProfiler.vbwExecuteLine 2522 'B
vbwProfiler.vbwExecuteLine 2523
        lTimeToQuit = (lStart - 86400) + TimeToWait
    End If
vbwProfiler.vbwExecuteLine 2524 'B

vbwProfiler.vbwExecuteLine 2525
TimeToQuit = lTimeToQuit

vbwProfiler.vbwProcOut 137
vbwProfiler.vbwExecuteLine 2526
End Function

Public Sub DisplayWinsock(Optional ErrorMessage As String)
vbwProfiler.vbwProcIn 138
Dim kb As String
vbwProfiler.vbwExecuteLine 2527
    With Winsock1
vbwProfiler.vbwExecuteLine 2528
            kb = kb & vbTab & "Local IP=" & .LocalIP & vbCrLf
vbwProfiler.vbwExecuteLine 2529
            kb = kb & vbTab & "Local Port=" & .LocalPort & vbCrLf
vbwProfiler.vbwExecuteLine 2530
            kb = kb & vbTab & "Protocol=" & aProtocol(.Protocol) & vbCrLf
vbwProfiler.vbwExecuteLine 2531
            kb = kb & vbTab & "Remote Host=" & .RemoteHost & vbCrLf
vbwProfiler.vbwExecuteLine 2532
            kb = kb & vbTab & "Remote Host IP=" & .RemoteHostIP & vbCrLf
vbwProfiler.vbwExecuteLine 2533
            kb = kb & vbTab & "Remote Port=" & .RemotePort & vbCrLf
vbwProfiler.vbwExecuteLine 2534
            kb = kb & vbTab & "State=" & aState(.State) & vbCrLf
vbwProfiler.vbwExecuteLine 2535
            If ErrorMessage <> "" Then
vbwProfiler.vbwExecuteLine 2536
                kb = kb & ErrorMessage & vbCrLf
            End If
vbwProfiler.vbwExecuteLine 2537 'B
vbwProfiler.vbwExecuteLine 2538
    End With
vbwProfiler.vbwExecuteLine 2539
    MsgBox kb, , "TCP/IP Sockets"
vbwProfiler.vbwProcOut 138
vbwProfiler.vbwExecuteLine 2540
End Sub

Public Function aProtocol(Protocol As Integer)
vbwProfiler.vbwProcIn 139
vbwProfiler.vbwExecuteLine 2541
    Select Case Protocol
'vbwLine 2542:    Case Is = sckTCPProtocol
    Case Is = IIf(vbwProfiler.vbwExecuteLine(2542), VBWPROFILER_EMPTY, _
        sckTCPProtocol)
vbwProfiler.vbwExecuteLine 2543
        aProtocol = "TCP"
'vbwLine 2544:    Case Is = sckUDPProtocol
    Case Is = IIf(vbwProfiler.vbwExecuteLine(2544), VBWPROFILER_EMPTY, _
        sckUDPProtocol)
vbwProfiler.vbwExecuteLine 2545
        aProtocol = "UDP"
    End Select
vbwProfiler.vbwExecuteLine 2546 'B
vbwProfiler.vbwProcOut 139
vbwProfiler.vbwExecuteLine 2547
End Function

Public Function aState(State As Integer) As String
vbwProfiler.vbwProcIn 140
vbwProfiler.vbwExecuteLine 2548
    Select Case State
'vbwLine 2549:    Case Is = -1
    Case Is = IIf(vbwProfiler.vbwExecuteLine(2549), VBWPROFILER_EMPTY, _
        -1)
vbwProfiler.vbwExecuteLine 2550
        aState = "Nothing"
'vbwLine 2551:    Case Is = 0
    Case Is = IIf(vbwProfiler.vbwExecuteLine(2551), VBWPROFILER_EMPTY, _
        0)
vbwProfiler.vbwExecuteLine 2552
        aState = "Closed"
'vbwLine 2553:    Case Is = 1
    Case Is = IIf(vbwProfiler.vbwExecuteLine(2553), VBWPROFILER_EMPTY, _
        1)
vbwProfiler.vbwExecuteLine 2554
        aState = "Open"
'vbwLine 2555:    Case Is = 2
    Case Is = IIf(vbwProfiler.vbwExecuteLine(2555), VBWPROFILER_EMPTY, _
        2)
vbwProfiler.vbwExecuteLine 2556
        aState = "Listening"
'vbwLine 2557:    Case Is = 3
    Case Is = IIf(vbwProfiler.vbwExecuteLine(2557), VBWPROFILER_EMPTY, _
        3)
vbwProfiler.vbwExecuteLine 2558
        aState = "Connection pending"
'vbwLine 2559:    Case Is = 4
    Case Is = IIf(vbwProfiler.vbwExecuteLine(2559), VBWPROFILER_EMPTY, _
        4)
vbwProfiler.vbwExecuteLine 2560
        aState = "Resolving host"
'vbwLine 2561:    Case Is = 5
    Case Is = IIf(vbwProfiler.vbwExecuteLine(2561), VBWPROFILER_EMPTY, _
        5)
vbwProfiler.vbwExecuteLine 2562
        aState = "Host resolved"
'vbwLine 2563:    Case Is = 6
    Case Is = IIf(vbwProfiler.vbwExecuteLine(2563), VBWPROFILER_EMPTY, _
        6)
vbwProfiler.vbwExecuteLine 2564
        aState = "Connecting"
'vbwLine 2565:    Case Is = 7
    Case Is = IIf(vbwProfiler.vbwExecuteLine(2565), VBWPROFILER_EMPTY, _
        7)
vbwProfiler.vbwExecuteLine 2566
        aState = "Connected"
'vbwLine 2567:    Case Is = 8
    Case Is = IIf(vbwProfiler.vbwExecuteLine(2567), VBWPROFILER_EMPTY, _
        8)
vbwProfiler.vbwExecuteLine 2568
        aState = "Peer is closing connection"
'vbwLine 2569:    Case Is = 9
    Case Is = IIf(vbwProfiler.vbwExecuteLine(2569), VBWPROFILER_EMPTY, _
        9)
vbwProfiler.vbwExecuteLine 2570
        aState = "Error"
'vbwLine 2571:    Case Is = 11
    Case Is = IIf(vbwProfiler.vbwExecuteLine(2571), VBWPROFILER_EMPTY, _
        11)
vbwProfiler.vbwExecuteLine 2572
        aState = "Opening"
'vbwLine 2573:    Case Is = 18
    Case Is = IIf(vbwProfiler.vbwExecuteLine(2573), VBWPROFILER_EMPTY, _
        18)
vbwProfiler.vbwExecuteLine 2574
        aState = "Closing"
'vbwLine 2575:    Case Is = 21
    Case Is = IIf(vbwProfiler.vbwExecuteLine(2575), VBWPROFILER_EMPTY, _
        21)
vbwProfiler.vbwExecuteLine 2576
        aState = "Data loss"
'vbwLine 2577:    Case Is = 22
    Case Is = IIf(vbwProfiler.vbwExecuteLine(2577), VBWPROFILER_EMPTY, _
        22)
vbwProfiler.vbwExecuteLine 2578
        aState = "Data in buffer"
    Case Else
vbwProfiler.vbwExecuteLine 2579 'B
vbwProfiler.vbwExecuteLine 2580
        aState = "Invalid"
    End Select
vbwProfiler.vbwExecuteLine 2581 'B
vbwProfiler.vbwProcOut 140
vbwProfiler.vbwExecuteLine 2582
End Function



