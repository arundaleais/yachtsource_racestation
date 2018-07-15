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
vbwProfiler.vbwProcIn 112
vbwProfiler.vbwExecuteLine 1688
    With mshEvents
vbwProfiler.vbwExecuteLine 1689
        .Top = ScaleTop
vbwProfiler.vbwExecuteLine 1690
        .Left = ScaleLeft
vbwProfiler.vbwExecuteLine 1691
        .Width = ScaleWidth
vbwProfiler.vbwExecuteLine 1692
        .Height = ScaleHeight
vbwProfiler.vbwExecuteLine 1693
        .FormatString = "^Event|<Time|Idx|<Signal|<Action"
vbwProfiler.vbwExecuteLine 1694
        .ColWidth(1) = 800  'Time
vbwProfiler.vbwExecuteLine 1695
        .ColWidth(2) = 0  'Idx
vbwProfiler.vbwExecuteLine 1696
        .ColWidth(3) = 1400 'Signal
vbwProfiler.vbwExecuteLine 1697
        .ColWidth(4) = 2420 'Action
'        For i = 1 To 20
'            .Rows = i + 1
'            .TextMatrix(i, 0) = i
'        Next i
'        .TextMatrix(1, 1) = "13:22:45"
vbwProfiler.vbwExecuteLine 1698
    End With
vbwProfiler.vbwExecuteLine 1699
    Call ListEvents

vbwProfiler.vbwProcOut 112
vbwProfiler.vbwExecuteLine 1700
End Sub

Private Function ListEvents()
vbwProfiler.vbwProcIn 113
Dim Row As Long
Dim Eidx As Long
Dim Sidx As Long
Dim Bidx As Long
Dim Idx As Long     'Keeps the last Signal or button (to combine Action on 1 line)
Dim kb As String

vbwProfiler.vbwExecuteLine 1701
    With mshEvents
vbwProfiler.vbwExecuteLine 1702
        For Eidx = 0 To UBound(Evts)
vbwProfiler.vbwExecuteLine 1703
            Row = Row + 1
'created with first row ""
vbwProfiler.vbwExecuteLine 1704
            If .TextMatrix(1, 0) <> "" Then
vbwProfiler.vbwExecuteLine 1705
                .AddItem Row, Row
            End If
vbwProfiler.vbwExecuteLine 1706 'B
vbwProfiler.vbwExecuteLine 1707
            .TextMatrix(Row, 0) = Row
vbwProfiler.vbwExecuteLine 1708
            .TextMatrix(Row, 1) = Format$(Evts(Eidx).ElapsedTime, "00:00")
vbwProfiler.vbwExecuteLine 1709
            If IsSignalsInitialised(Evts(Eidx).Signals) Then
vbwProfiler.vbwExecuteLine 1710
                For Sidx = 0 To UBound(Evts(Eidx).Signals)
vbwProfiler.vbwExecuteLine 1711
                    If Evts(Eidx).Signals(Sidx).Signal <> Idx Then
'Add another Row after First Eidx for this time
vbwProfiler.vbwExecuteLine 1712
                        If Sidx > 0 Then
vbwProfiler.vbwExecuteLine 1713
                            Row = Row + 1
vbwProfiler.vbwExecuteLine 1714
                            .AddItem Row, Row
                        End If
vbwProfiler.vbwExecuteLine 1715 'B
                    End If
vbwProfiler.vbwExecuteLine 1716 'B
vbwProfiler.vbwExecuteLine 1717
                    Idx = Evts(Eidx).Signals(Sidx).Signal
vbwProfiler.vbwExecuteLine 1718
                    .TextMatrix(Row, 2) = Idx
vbwProfiler.vbwExecuteLine 1719
                    .TextMatrix(Row, 3) = SignalAttributes(Idx).Name
vbwProfiler.vbwExecuteLine 1720
                    kb = Evts(Eidx).Signals(Sidx).Raise
vbwProfiler.vbwExecuteLine 1721
                    If kb <> "" Then
vbwProfiler.vbwExecuteLine 1722
                        If kb = "True" Then
vbwProfiler.vbwExecuteLine 1723
                            Call AddMessage(Row, "Up")
                        Else
vbwProfiler.vbwExecuteLine 1724 'B
vbwProfiler.vbwExecuteLine 1725
                            Call AddMessage(Row, "Down")
                        End If
vbwProfiler.vbwExecuteLine 1726 'B
                    End If
vbwProfiler.vbwExecuteLine 1727 'B
vbwProfiler.vbwExecuteLine 1728
                    kb = Evts(Eidx).Signals(Sidx).Silent
vbwProfiler.vbwExecuteLine 1729
                    If kb <> "" Then
vbwProfiler.vbwExecuteLine 1730
                        If kb = "True" Then
vbwProfiler.vbwExecuteLine 1731
                            Call AddMessage(Row, "Silent")
                        Else
vbwProfiler.vbwExecuteLine 1732 'B
vbwProfiler.vbwExecuteLine 1733
                            Call AddMessage(Row, "Sound")
                        End If
vbwProfiler.vbwExecuteLine 1734 'B
                    End If
vbwProfiler.vbwExecuteLine 1735 'B
vbwProfiler.vbwExecuteLine 1736
                Next Sidx
            End If
vbwProfiler.vbwExecuteLine 1737 'B
vbwProfiler.vbwExecuteLine 1738
            If IsButtonsInitialised(Evts(Eidx).Buttons) Then
vbwProfiler.vbwExecuteLine 1739
                For Bidx = 0 To UBound(Evts(Eidx).Buttons)
'Add another Row after First Eidx for this time
vbwProfiler.vbwExecuteLine 1740
                    If Evts(Eidx).Buttons(Bidx).Button <> Idx Then
'Add another Row after First Eidx for this time
vbwProfiler.vbwExecuteLine 1741
                        If Sidx > 0 Or Bidx > 0 Then
vbwProfiler.vbwExecuteLine 1742
                            Row = Row + 1
vbwProfiler.vbwExecuteLine 1743
                            .AddItem Row, Row
                        End If
vbwProfiler.vbwExecuteLine 1744 'B
                    End If
vbwProfiler.vbwExecuteLine 1745 'B
vbwProfiler.vbwExecuteLine 1746
                    Idx = Evts(Eidx).Buttons(Bidx).Button
vbwProfiler.vbwExecuteLine 1747
                    .TextMatrix(Row, 2) = Idx
vbwProfiler.vbwExecuteLine 1748
                    .TextMatrix(Row, 3) = SignalAttributes(Idx).Name
vbwProfiler.vbwExecuteLine 1749
                    kb = Evts(Eidx).Buttons(Bidx).Enabled
vbwProfiler.vbwExecuteLine 1750
                    If kb <> "" Then
vbwProfiler.vbwExecuteLine 1751
                        If kb = "True" Then
vbwProfiler.vbwExecuteLine 1752
                            Call AddMessage(Row, "Enabled")
                        Else
vbwProfiler.vbwExecuteLine 1753 'B
vbwProfiler.vbwExecuteLine 1754
                            Call AddMessage(Row, "Disabled")
                        End If
vbwProfiler.vbwExecuteLine 1755 'B
                    End If
vbwProfiler.vbwExecuteLine 1756 'B
vbwProfiler.vbwExecuteLine 1757
                Next Bidx
            End If
vbwProfiler.vbwExecuteLine 1758 'B
vbwProfiler.vbwExecuteLine 1759
            If Evts(Eidx).Focus > 0 Then
vbwProfiler.vbwExecuteLine 1760
                Call AddMessage(Row, "Focus-" & SignalAttributes(Evts(Eidx).Focus).Name)
            End If
vbwProfiler.vbwExecuteLine 1761 'B
vbwProfiler.vbwExecuteLine 1762
        Next Eidx
vbwProfiler.vbwExecuteLine 1763
Me.Visible = True
vbwProfiler.vbwExecuteLine 1764
    .Col = 0
vbwProfiler.vbwExecuteLine 1765
    .Row = 0
vbwProfiler.vbwExecuteLine 1766
    .FocusRect = flexFocusNone ' (The selected cell changes)
vbwProfiler.vbwExecuteLine 1767
    End With
vbwProfiler.vbwExecuteLine 1768
Me.Visible = False
vbwProfiler.vbwProcOut 113
vbwProfiler.vbwExecuteLine 1769
End Function

Private Function AddMessage(Row, kb)
vbwProfiler.vbwProcIn 114
vbwProfiler.vbwExecuteLine 1770
    With mshEvents
vbwProfiler.vbwExecuteLine 1771
        If .TextMatrix(Row, 4) = "" Then
vbwProfiler.vbwExecuteLine 1772
            .TextMatrix(Row, 4) = kb
        Else
vbwProfiler.vbwExecuteLine 1773 'B
vbwProfiler.vbwExecuteLine 1774
            .TextMatrix(Row, 4) = .TextMatrix(Row, 4) & "," & kb
        End If
vbwProfiler.vbwExecuteLine 1775 'B
vbwProfiler.vbwExecuteLine 1776
    End With
vbwProfiler.vbwProcOut 114
vbwProfiler.vbwExecuteLine 1777
End Function

Private Sub mshEvents_SelChange()
vbwProfiler.vbwProcIn 115
vbwProfiler.vbwExecuteLine 1778
    Call ActionEvent
vbwProfiler.vbwProcOut 115
vbwProfiler.vbwExecuteLine 1779
End Sub

Private Function ActionEvent()
vbwProfiler.vbwProcIn 116
Dim Idx As Long
Dim MyEvent As clsEvent
Dim Index As Long
Dim ElapsedTime As Long


vbwProfiler.vbwExecuteLine 1780
    With mshEvents
vbwProfiler.vbwExecuteLine 1781
        If .Row < 1 Then 'Miss Header Row
vbwProfiler.vbwExecuteLine 1782
             .Row = 1
        End If
vbwProfiler.vbwExecuteLine 1783 'B
vbwProfiler.vbwExecuteLine 1784
            Index = .TextMatrix(.Row, 0)
'                If .Col = 0 Then
'                    Call frmMain.DoEvent(Index)
'                End If
vbwProfiler.vbwExecuteLine 1785
                If .Col = 1 Then
vbwProfiler.vbwExecuteLine 1786
                   If .TextMatrix(.Row, .Col) <> "" Then
vbwProfiler.vbwExecuteLine 1787
                        ElapsedTime = Replace(.TextMatrix(.Row, .Col), ":", "")
vbwProfiler.vbwExecuteLine 1788
                        Call frmMain.DoTimerEvents(ElapsedTime)
                   End If
vbwProfiler.vbwExecuteLine 1789 'B
                End If
vbwProfiler.vbwExecuteLine 1790 'B
vbwProfiler.vbwExecuteLine 1791
        If .Row = .Rows - 1 Then
vbwProfiler.vbwExecuteLine 1792
            .Row = 0
vbwProfiler.vbwExecuteLine 1793
            .TopRow = 1
        End If
vbwProfiler.vbwExecuteLine 1794 'B
vbwProfiler.vbwExecuteLine 1795
    End With

vbwProfiler.vbwProcOut 116
vbwProfiler.vbwExecuteLine 1796
End Function


Private Function ListEvents_old()
vbwProfiler.vbwProcIn 117
Dim MyEvent As clsEvent
Dim Row As Long
Dim LastSec As Long
'    Me.Show
vbwProfiler.vbwExecuteLine 1797
    With mshEvents
vbwProfiler.vbwExecuteLine 1798
        For Each MyEvent In Myprofile
vbwProfiler.vbwExecuteLine 1799
            Row = MyEvent.Index
'created with first row ""
vbwProfiler.vbwExecuteLine 1800
            If .TextMatrix(1, 0) <> "" Then
vbwProfiler.vbwExecuteLine 1801
                .AddItem Row, Row
            End If
vbwProfiler.vbwExecuteLine 1802 'B
vbwProfiler.vbwExecuteLine 1803
            .TextMatrix(Row, 0) = Row
vbwProfiler.vbwExecuteLine 1804
            If MyEvent.Second <> LastSec Then
vbwProfiler.vbwExecuteLine 1805
                .TextMatrix(Row, 1) = Format$(MyEvent.Second, "00:00")
            End If
vbwProfiler.vbwExecuteLine 1806 'B
vbwProfiler.vbwExecuteLine 1807
            .TextMatrix(Row, 2) = MyEvent.Signal
vbwProfiler.vbwExecuteLine 1808
            .TextMatrix(Row, 3) = SignalAttributes(MyEvent.Signal).Name
vbwProfiler.vbwExecuteLine 1809
            If MyEvent.Raised = True Then
vbwProfiler.vbwExecuteLine 1810
                If SignalAttributes(MyEvent.Signal).TTL = 0 Then
vbwProfiler.vbwExecuteLine 1811
                    .TextMatrix(Row, 4) = "Up"
                Else
vbwProfiler.vbwExecuteLine 1812 'B
vbwProfiler.vbwExecuteLine 1813
                    .TextMatrix(Row, 4) = "On"
                End If
vbwProfiler.vbwExecuteLine 1814 'B
            Else
vbwProfiler.vbwExecuteLine 1815 'B
vbwProfiler.vbwExecuteLine 1816
                If SignalAttributes(MyEvent.Signal).TTL = 0 Then
vbwProfiler.vbwExecuteLine 1817
                    .TextMatrix(Row, 4) = "Down"
                Else
vbwProfiler.vbwExecuteLine 1818 'B
vbwProfiler.vbwExecuteLine 1819
                    .TextMatrix(Row, 4) = "Off"
                End If
vbwProfiler.vbwExecuteLine 1820 'B
            End If
vbwProfiler.vbwExecuteLine 1821 'B
vbwProfiler.vbwExecuteLine 1822
            LastSec = MyEvent.Second
vbwProfiler.vbwExecuteLine 1823
        Next MyEvent
vbwProfiler.vbwExecuteLine 1824
Me.Visible = True
vbwProfiler.vbwExecuteLine 1825
    .Col = 0
vbwProfiler.vbwExecuteLine 1826
    .Row = 0
vbwProfiler.vbwExecuteLine 1827
    .FocusRect = flexFocusNone ' (The selected cell changes)
vbwProfiler.vbwExecuteLine 1828
    End With
vbwProfiler.vbwExecuteLine 1829
Me.Visible = False
vbwProfiler.vbwProcOut 117
vbwProfiler.vbwExecuteLine 1830
End Function

Private Function ActionEvent_old()
vbwProfiler.vbwProcIn 118
Dim Idx As Long
Dim MyEvent As clsEvent
Dim Index As Long
Dim ElapsedTime As Long


vbwProfiler.vbwExecuteLine 1831
    With mshEvents
vbwProfiler.vbwExecuteLine 1832
        If .Row < 1 Then 'Miss Header Row
vbwProfiler.vbwExecuteLine 1833
             .Row = 1
        End If
vbwProfiler.vbwExecuteLine 1834 'B
vbwProfiler.vbwExecuteLine 1835
            Index = .TextMatrix(.Row, 0)
vbwProfiler.vbwExecuteLine 1836
            For Each MyEvent In Myprofile
vbwProfiler.vbwExecuteLine 1837
                If .Col = 0 Then
vbwProfiler.vbwExecuteLine 1838
                    If .Row = MyEvent.Index Then
vbwProfiler.vbwExecuteLine 1839
                        Call frmMain.DoEvent(MyEvent)
                    End If
vbwProfiler.vbwExecuteLine 1840 'B
                End If
vbwProfiler.vbwExecuteLine 1841 'B
vbwProfiler.vbwExecuteLine 1842
                If .Col = 1 Then
vbwProfiler.vbwExecuteLine 1843
                   If .TextMatrix(.Row, .Col) <> "" Then
vbwProfiler.vbwExecuteLine 1844
                        ElapsedTime = Replace(.TextMatrix(.Row, .Col), ":", "")
vbwProfiler.vbwExecuteLine 1845
                        Call frmMain.DoTimerEvents(ElapsedTime)
                   End If
vbwProfiler.vbwExecuteLine 1846 'B
                End If
vbwProfiler.vbwExecuteLine 1847 'B
vbwProfiler.vbwExecuteLine 1848
            Next MyEvent
vbwProfiler.vbwExecuteLine 1849
        If .Row = .Rows - 1 Then
vbwProfiler.vbwExecuteLine 1850
            .Row = 0
vbwProfiler.vbwExecuteLine 1851
            .TopRow = 1
        End If
vbwProfiler.vbwExecuteLine 1852 'B
vbwProfiler.vbwExecuteLine 1853
    End With

vbwProfiler.vbwProcOut 118
vbwProfiler.vbwExecuteLine 1854
End Function



