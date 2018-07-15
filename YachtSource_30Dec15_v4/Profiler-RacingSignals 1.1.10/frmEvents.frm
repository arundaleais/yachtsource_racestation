VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFlxGd.ocx"
Begin VB.Form frmEvents 
   Caption         =   "Events"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   6075
   StartUpPosition =   3  'Windows Default
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshEvents 
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   2566
      _Version        =   393216
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frmEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
vbwProfiler.vbwProcIn 83
vbwProfiler.vbwExecuteLine 1520
    With mshEvents
vbwProfiler.vbwExecuteLine 1521
        .Top = ScaleTop
vbwProfiler.vbwExecuteLine 1522
        .Left = ScaleLeft
vbwProfiler.vbwExecuteLine 1523
        .Width = ScaleWidth
vbwProfiler.vbwExecuteLine 1524
        .Height = ScaleHeight
vbwProfiler.vbwExecuteLine 1525
        .FormatString = "^Event|<Time|Idx|<Signal|<Action"
vbwProfiler.vbwExecuteLine 1526
        .ColWidth(1) = 800  'Time
vbwProfiler.vbwExecuteLine 1527
        .ColWidth(2) = 0  'Idx
vbwProfiler.vbwExecuteLine 1528
        .ColWidth(3) = 1400 'Signal
vbwProfiler.vbwExecuteLine 1529
        .ColWidth(4) = 2420 'Action
'        For i = 1 To 20
'            .Rows = i + 1
'            .TextMatrix(i, 0) = i
'        Next i
'        .TextMatrix(1, 1) = "13:22:45"
vbwProfiler.vbwExecuteLine 1530
    End With
vbwProfiler.vbwExecuteLine 1531
    Call ListEvents

vbwProfiler.vbwProcOut 83
vbwProfiler.vbwExecuteLine 1532
End Sub

Private Function ListEvents()
vbwProfiler.vbwProcIn 84
Dim Row As Long
Dim Eidx As Long
Dim Sidx As Long
Dim Bidx As Long
Dim Idx As Long     'Keeps the last Signal or button (to combine Action on 1 line)
Dim kb As String

vbwProfiler.vbwExecuteLine 1533
    With mshEvents
vbwProfiler.vbwExecuteLine 1534
        For Eidx = 0 To UBound(Evts)
vbwProfiler.vbwExecuteLine 1535
            Row = Row + 1
'created with first row ""
vbwProfiler.vbwExecuteLine 1536
            If .TextMatrix(1, 0) <> "" Then
vbwProfiler.vbwExecuteLine 1537
                .AddItem Row, Row
            End If
vbwProfiler.vbwExecuteLine 1538 'B
vbwProfiler.vbwExecuteLine 1539
            .TextMatrix(Row, 0) = Row
vbwProfiler.vbwExecuteLine 1540
            .TextMatrix(Row, 1) = Format$(Evts(Eidx).ElapsedTime, "00:00")
vbwProfiler.vbwExecuteLine 1541
            If IsSignalsInitialised(Evts(Eidx).Signals) Then
vbwProfiler.vbwExecuteLine 1542
                For Sidx = 0 To UBound(Evts(Eidx).Signals)
vbwProfiler.vbwExecuteLine 1543
                    If Evts(Eidx).Signals(Sidx).Signal <> Idx Then
'Add another Row after First Eidx for this time
vbwProfiler.vbwExecuteLine 1544
                        If Sidx > 0 Then
vbwProfiler.vbwExecuteLine 1545
                            Row = Row + 1
vbwProfiler.vbwExecuteLine 1546
                            .AddItem Row, Row
                        End If
vbwProfiler.vbwExecuteLine 1547 'B
                    End If
vbwProfiler.vbwExecuteLine 1548 'B
vbwProfiler.vbwExecuteLine 1549
                    Idx = Evts(Eidx).Signals(Sidx).Signal
vbwProfiler.vbwExecuteLine 1550
                    .TextMatrix(Row, 2) = Idx
vbwProfiler.vbwExecuteLine 1551
                    .TextMatrix(Row, 3) = SignalAttributes(Idx).Name
vbwProfiler.vbwExecuteLine 1552
                    kb = Evts(Eidx).Signals(Sidx).Raise
vbwProfiler.vbwExecuteLine 1553
                    If kb <> "" Then
vbwProfiler.vbwExecuteLine 1554
                        If kb = "True" Then
vbwProfiler.vbwExecuteLine 1555
                            Call AddMessage(Row, "Up")
                        Else
vbwProfiler.vbwExecuteLine 1556 'B
vbwProfiler.vbwExecuteLine 1557
                            Call AddMessage(Row, "Down")
                        End If
vbwProfiler.vbwExecuteLine 1558 'B
                    End If
vbwProfiler.vbwExecuteLine 1559 'B
vbwProfiler.vbwExecuteLine 1560
                    kb = Evts(Eidx).Signals(Sidx).Silent
vbwProfiler.vbwExecuteLine 1561
                    If kb <> "" Then
vbwProfiler.vbwExecuteLine 1562
                        If kb = "True" Then
vbwProfiler.vbwExecuteLine 1563
                            Call AddMessage(Row, "Silent")
                        Else
vbwProfiler.vbwExecuteLine 1564 'B
vbwProfiler.vbwExecuteLine 1565
                            Call AddMessage(Row, "Sound")
                        End If
vbwProfiler.vbwExecuteLine 1566 'B
                    End If
vbwProfiler.vbwExecuteLine 1567 'B
vbwProfiler.vbwExecuteLine 1568
                Next Sidx
            End If
vbwProfiler.vbwExecuteLine 1569 'B
vbwProfiler.vbwExecuteLine 1570
            If IsButtonsInitialised(Evts(Eidx).Buttons) Then
vbwProfiler.vbwExecuteLine 1571
                For Bidx = 0 To UBound(Evts(Eidx).Buttons)
'Add another Row after First Eidx for this time
vbwProfiler.vbwExecuteLine 1572
                    If Evts(Eidx).Buttons(Bidx).Button <> Idx Then
'Add another Row after First Eidx for this time
vbwProfiler.vbwExecuteLine 1573
                        If Sidx > 0 Or Bidx > 0 Then
vbwProfiler.vbwExecuteLine 1574
                            Row = Row + 1
vbwProfiler.vbwExecuteLine 1575
                            .AddItem Row, Row
                        End If
vbwProfiler.vbwExecuteLine 1576 'B
                    End If
vbwProfiler.vbwExecuteLine 1577 'B
vbwProfiler.vbwExecuteLine 1578
                    Idx = Evts(Eidx).Buttons(Bidx).Button
vbwProfiler.vbwExecuteLine 1579
                    .TextMatrix(Row, 2) = Idx
vbwProfiler.vbwExecuteLine 1580
                    .TextMatrix(Row, 3) = SignalAttributes(Idx).Name
vbwProfiler.vbwExecuteLine 1581
                    kb = Evts(Eidx).Buttons(Bidx).Enabled
vbwProfiler.vbwExecuteLine 1582
                    If kb <> "" Then
vbwProfiler.vbwExecuteLine 1583
                        If kb = "True" Then
vbwProfiler.vbwExecuteLine 1584
                            Call AddMessage(Row, "Enabled")
                        Else
vbwProfiler.vbwExecuteLine 1585 'B
vbwProfiler.vbwExecuteLine 1586
                            Call AddMessage(Row, "Disabled")
                        End If
vbwProfiler.vbwExecuteLine 1587 'B
                    End If
vbwProfiler.vbwExecuteLine 1588 'B
vbwProfiler.vbwExecuteLine 1589
                Next Bidx
            End If
vbwProfiler.vbwExecuteLine 1590 'B
vbwProfiler.vbwExecuteLine 1591
            If Evts(Eidx).Focus > 0 Then
vbwProfiler.vbwExecuteLine 1592
                Call AddMessage(Row, "Focus-" & SignalAttributes(Evts(Eidx).Focus).Name)
            End If
vbwProfiler.vbwExecuteLine 1593 'B
vbwProfiler.vbwExecuteLine 1594
        Next Eidx
vbwProfiler.vbwExecuteLine 1595
Me.Visible = True
vbwProfiler.vbwExecuteLine 1596
    .Col = 0
vbwProfiler.vbwExecuteLine 1597
    .Row = 0
vbwProfiler.vbwExecuteLine 1598
    .FocusRect = flexFocusNone ' (The selected cell changes)
vbwProfiler.vbwExecuteLine 1599
    End With
vbwProfiler.vbwExecuteLine 1600
Me.Visible = False
vbwProfiler.vbwProcOut 84
vbwProfiler.vbwExecuteLine 1601
End Function

Private Function AddMessage(Row, kb)
vbwProfiler.vbwProcIn 85
vbwProfiler.vbwExecuteLine 1602
    With mshEvents
vbwProfiler.vbwExecuteLine 1603
        If .TextMatrix(Row, 4) = "" Then
vbwProfiler.vbwExecuteLine 1604
            .TextMatrix(Row, 4) = kb
        Else
vbwProfiler.vbwExecuteLine 1605 'B
vbwProfiler.vbwExecuteLine 1606
            .TextMatrix(Row, 4) = .TextMatrix(Row, 4) & "," & kb
        End If
vbwProfiler.vbwExecuteLine 1607 'B
vbwProfiler.vbwExecuteLine 1608
    End With
vbwProfiler.vbwProcOut 85
vbwProfiler.vbwExecuteLine 1609
End Function

Private Sub mshEvents_SelChange()
vbwProfiler.vbwProcIn 86
vbwProfiler.vbwExecuteLine 1610
    Call ActionEvent
vbwProfiler.vbwProcOut 86
vbwProfiler.vbwExecuteLine 1611
End Sub

Private Function ActionEvent()
vbwProfiler.vbwProcIn 87
Dim Idx As Long
Dim Index As Long
Dim ElapsedTime As Long


vbwProfiler.vbwExecuteLine 1612
    With mshEvents
vbwProfiler.vbwExecuteLine 1613
        If .Row < 1 Then 'Miss Header Row
vbwProfiler.vbwExecuteLine 1614
             .Row = 1
        End If
vbwProfiler.vbwExecuteLine 1615 'B
vbwProfiler.vbwExecuteLine 1616
            Index = .TextMatrix(.Row, 0)
'                If .Col = 0 Then
'                    Call frmMain.DoEvent(Index)
'                End If
vbwProfiler.vbwExecuteLine 1617
                If .Col = 1 Then
vbwProfiler.vbwExecuteLine 1618
                   If .TextMatrix(.Row, .Col) <> "" Then
vbwProfiler.vbwExecuteLine 1619
                        ElapsedTime = Replace(.TextMatrix(.Row, .Col), ":", "")
vbwProfiler.vbwExecuteLine 1620
                        Call frmMain.DoTimerEvents(ElapsedTime)
                   End If
vbwProfiler.vbwExecuteLine 1621 'B
                End If
vbwProfiler.vbwExecuteLine 1622 'B
vbwProfiler.vbwExecuteLine 1623
        If .Row = .Rows - 1 Then
vbwProfiler.vbwExecuteLine 1624
            .Row = 0
vbwProfiler.vbwExecuteLine 1625
            .TopRow = 1
        End If
vbwProfiler.vbwExecuteLine 1626 'B
vbwProfiler.vbwExecuteLine 1627
    End With

vbwProfiler.vbwProcOut 87
vbwProfiler.vbwExecuteLine 1628
End Function




