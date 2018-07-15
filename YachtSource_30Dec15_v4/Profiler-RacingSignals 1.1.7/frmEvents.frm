VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFlxGd.ocx"
Begin VB.Form frmEvents 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   6315
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
vbwProfiler.vbwProcIn 101
vbwProfiler.vbwExecuteLine 1453
    With mshEvents
vbwProfiler.vbwExecuteLine 1454
        .Top = ScaleTop
vbwProfiler.vbwExecuteLine 1455
        .Left = ScaleLeft
vbwProfiler.vbwExecuteLine 1456
        .Width = ScaleWidth
vbwProfiler.vbwExecuteLine 1457
        .Height = ScaleHeight
vbwProfiler.vbwExecuteLine 1458
        .FormatString = "^Event|<Time|<Signal|<Action"
vbwProfiler.vbwExecuteLine 1459
        .ColWidth(1) = 800  'Time
vbwProfiler.vbwExecuteLine 1460
        .ColWidth(2) = 1400 'Signal
vbwProfiler.vbwExecuteLine 1461
        .ColWidth(3) = 2420 'Action
'        For i = 1 To 20
'            .Rows = i + 1
'            .TextMatrix(i, 0) = i
'        Next i
'        .TextMatrix(1, 1) = "13:22:45"
vbwProfiler.vbwExecuteLine 1462
    End With
vbwProfiler.vbwExecuteLine 1463
    Call ListEvents

vbwProfiler.vbwProcOut 101
vbwProfiler.vbwExecuteLine 1464
End Sub

Private Function ListEvents()
vbwProfiler.vbwProcIn 102
Dim MyEvent As clsEvent
Dim Row As Long
vbwProfiler.vbwExecuteLine 1465
    Me.Show
vbwProfiler.vbwExecuteLine 1466
    With mshEvents
vbwProfiler.vbwExecuteLine 1467
        For Each MyEvent In Myprofile
vbwProfiler.vbwExecuteLine 1468
            Row = MyEvent.Index
'created with first row ""
vbwProfiler.vbwExecuteLine 1469
            If .TextMatrix(1, 0) <> "" Then
vbwProfiler.vbwExecuteLine 1470
                .AddItem Row, Row
            End If
vbwProfiler.vbwExecuteLine 1471 'B
vbwProfiler.vbwExecuteLine 1472
            .TextMatrix(Row, 0) = Row
vbwProfiler.vbwExecuteLine 1473
            .TextMatrix(Row, 1) = Format$(MyEvent.Second, "00:00")
vbwProfiler.vbwExecuteLine 1474
            .TextMatrix(Row, 2) = SignalAttributes(MyEvent.Signal).Name
vbwProfiler.vbwExecuteLine 1475
            If MyEvent.Raised = True Then
vbwProfiler.vbwExecuteLine 1476
                If SignalAttributes(MyEvent.Signal).TTL = 0 Then
vbwProfiler.vbwExecuteLine 1477
                    .TextMatrix(Row, 3) = "Up"
                Else
vbwProfiler.vbwExecuteLine 1478 'B
vbwProfiler.vbwExecuteLine 1479
                    .TextMatrix(Row, 3) = "On"
                End If
vbwProfiler.vbwExecuteLine 1480 'B
            Else
vbwProfiler.vbwExecuteLine 1481 'B
vbwProfiler.vbwExecuteLine 1482
                If SignalAttributes(MyEvent.Signal).TTL = 0 Then
vbwProfiler.vbwExecuteLine 1483
                    .TextMatrix(Row, 3) = "Down"
                Else
vbwProfiler.vbwExecuteLine 1484 'B
vbwProfiler.vbwExecuteLine 1485
                    .TextMatrix(Row, 3) = "Off"
                End If
vbwProfiler.vbwExecuteLine 1486 'B
            End If
vbwProfiler.vbwExecuteLine 1487 'B
vbwProfiler.vbwExecuteLine 1488
        Next MyEvent
vbwProfiler.vbwExecuteLine 1489
    End With
vbwProfiler.vbwProcOut 102
vbwProfiler.vbwExecuteLine 1490
End Function


