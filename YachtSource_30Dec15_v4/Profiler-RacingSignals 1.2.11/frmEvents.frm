VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFlxGd.ocx"
Begin VB.Form frmEvents 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Events"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   6075
   StartUpPosition =   3  'Windows Default
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshEvents 
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      _ExtentX        =   10186
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
vbwProfiler.vbwProcIn 110
vbwProfiler.vbwExecuteLine 2145
    Height = Screen.Height - 1000
vbwProfiler.vbwExecuteLine 2146
    Width = 6800
vbwProfiler.vbwExecuteLine 2147
    Left = Screen.Width - Width
vbwProfiler.vbwExecuteLine 2148
    Top = 100

'    Top = Screen.Top
vbwProfiler.vbwExecuteLine 2149
    With mshEvents
vbwProfiler.vbwExecuteLine 2150
        .Top = ScaleTop
vbwProfiler.vbwExecuteLine 2151
        .Left = ScaleLeft
vbwProfiler.vbwExecuteLine 2152
        .Width = ScaleWidth
vbwProfiler.vbwExecuteLine 2153
        .Height = ScaleHeight
vbwProfiler.vbwExecuteLine 2154
        .FormatString = "^Event|<Time|Idx|<Signal|<Action|<Class"
vbwProfiler.vbwExecuteLine 2155
        .ColWidth(1) = 800  'Time
vbwProfiler.vbwExecuteLine 2156
        .ColWidth(2) = 0  'Idx
vbwProfiler.vbwExecuteLine 2157
        .ColWidth(3) = 2000 'Signal
vbwProfiler.vbwExecuteLine 2158
        .ColWidth(4) = 2000 'Action
vbwProfiler.vbwExecuteLine 2159
        .ColWidth(5) = 1000
'        For i = 1 To 20
'            .Rows = i + 1
'            .TextMatrix(i, 0) = i
'        Next i
'        .TextMatrix(1, 1) = "13:22:45"
vbwProfiler.vbwExecuteLine 2160
    End With

vbwProfiler.vbwProcOut 110
vbwProfiler.vbwExecuteLine 2161
End Sub

Public Function ListEvents()
vbwProfiler.vbwProcIn 111
Dim Row As Long
Dim Eidx As Long
Dim Sidx As Long
Dim Bidx As Long
Dim Idx As Long     'Keeps the last Signal or button (to combine Action on 1 line)
Dim kb As String
Dim i As Long
Dim Class As String

'    .Visible = True
'    frmEvents.SetFocus
vbwProfiler.vbwExecuteLine 2162
    With mshEvents
'vbwLine 2163:        Do While .Rows > 2
        Do While vbwProfiler.vbwExecuteLine(2163) Or .Rows > 2
vbwProfiler.vbwExecuteLine 2164
            .RemoveItem .Rows - 1
vbwProfiler.vbwExecuteLine 2165
        Loop
vbwProfiler.vbwExecuteLine 2166
        For i = 0 To .Cols - 1
vbwProfiler.vbwExecuteLine 2167
            .TextMatrix(1, i) = ""  'blank remaining rows
vbwProfiler.vbwExecuteLine 2168
            .BackColor = vbWhite
vbwProfiler.vbwExecuteLine 2169
        Next i
vbwProfiler.vbwExecuteLine 2170
        Row = 0
vbwProfiler.vbwExecuteLine 2171
        For Eidx = 0 To UBound(Evts)
vbwProfiler.vbwExecuteLine 2172
            Row = Row + 1

'created with first row ""
vbwProfiler.vbwExecuteLine 2173
            If .TextMatrix(1, 0) <> "" Then
vbwProfiler.vbwExecuteLine 2174
                .AddItem Row, Row
            End If
vbwProfiler.vbwExecuteLine 2175 'B
vbwProfiler.vbwExecuteLine 2176
            .TextMatrix(Row, 0) = Row
vbwProfiler.vbwExecuteLine 2177
            .TextMatrix(Row, 1) = aMins(Evts(Eidx).ElapsedTime + Classes(Evts(Eidx).Class).Offset)
vbwProfiler.vbwExecuteLine 2178
            .TextMatrix(Row, 1) = Evts(Eidx).ElapsedTime + Classes(Evts(Eidx).Class).Offset
'This will be overwritten if there is a Sidx or Bidx
vbwProfiler.vbwExecuteLine 2179
            .TextMatrix(Row, 3) = Evts(Eidx).Message
vbwProfiler.vbwExecuteLine 2180
            If Evts(Eidx).Focus >= 0 Then
vbwProfiler.vbwExecuteLine 2181
                Call AddMessage(Row, "Focus-" & frmMain.Commands(Evts(Eidx).Focus).Caption)
            End If
vbwProfiler.vbwExecuteLine 2182 'B
vbwProfiler.vbwExecuteLine 2183
            If Evts(Eidx).Signal > 0 Then
vbwProfiler.vbwExecuteLine 2184
                .TextMatrix(Row, 5) = SignalAttributes(Evts(Eidx).Signal).Name
            End If
vbwProfiler.vbwExecuteLine 2185 'B
vbwProfiler.vbwExecuteLine 2186
            If IsSignalsInitialised(Evts(Eidx).Signals) Then
vbwProfiler.vbwExecuteLine 2187
                For Sidx = 0 To UBound(Evts(Eidx).Signals)
'                   If Evts(Eidx).Signals(Sidx).Signal <> Idx Then
'Add another Row after First Eidx for this time
'                        If Sidx > 0 Then
vbwProfiler.vbwExecuteLine 2188
                            Row = Row + 1
vbwProfiler.vbwExecuteLine 2189
                            .AddItem Row, Row
'                        End If
'                    End If
vbwProfiler.vbwExecuteLine 2190
                    Idx = Evts(Eidx).Signals(Sidx).Signal
vbwProfiler.vbwExecuteLine 2191
                    .TextMatrix(Row, 2) = Idx
vbwProfiler.vbwExecuteLine 2192
                    .TextMatrix(Row, 3) = SignalAttributes(Idx).Name
vbwProfiler.vbwExecuteLine 2193
                    kb = Evts(Eidx).Signals(Sidx).Raise
vbwProfiler.vbwExecuteLine 2194
                    If kb <> "" Then
vbwProfiler.vbwExecuteLine 2195
                        If kb = "True" Then
vbwProfiler.vbwExecuteLine 2196
                            Call AddMessage(Row, "Up")
                        Else
vbwProfiler.vbwExecuteLine 2197 'B
vbwProfiler.vbwExecuteLine 2198
                            Call AddMessage(Row, "Down")
                        End If
vbwProfiler.vbwExecuteLine 2199 'B
                    End If
vbwProfiler.vbwExecuteLine 2200 'B
vbwProfiler.vbwExecuteLine 2201
                    kb = Evts(Eidx).Signals(Sidx).Silent
vbwProfiler.vbwExecuteLine 2202
                    If kb <> "" Then
vbwProfiler.vbwExecuteLine 2203
                        If kb = "True" Then
vbwProfiler.vbwExecuteLine 2204
                            Call AddMessage(Row, "Silent")
                        Else
vbwProfiler.vbwExecuteLine 2205 'B
vbwProfiler.vbwExecuteLine 2206
                            Call AddMessage(Row, "Sound")
                        End If
vbwProfiler.vbwExecuteLine 2207 'B
                    End If
vbwProfiler.vbwExecuteLine 2208 'B
'                    If Evts(Eidx).Signals(Sidx).signal > 0 Then
'                        .TextMatrix(Row, 5) = SignalAttributes(Evts(Eidx).Signals(Sidx).signal).Name
'                    End If
vbwProfiler.vbwExecuteLine 2209
                Next Sidx
            End If
vbwProfiler.vbwExecuteLine 2210 'B
vbwProfiler.vbwExecuteLine 2211
            If IsButtonsInitialised(Evts(Eidx).Buttons) Then
vbwProfiler.vbwExecuteLine 2212
                For Bidx = 0 To UBound(Evts(Eidx).Buttons)
'Add another Row after First Eidx for this time
'                    If Evts(Eidx).Buttons(Bidx).Button <> Idx Then
'Add another Row after First Eidx for this time
'                        If Sidx > 0 Or Bidx > 0 Then
vbwProfiler.vbwExecuteLine 2213
                            Row = Row + 1
vbwProfiler.vbwExecuteLine 2214
                            .AddItem Row, Row
'                        End If
'                    End If
vbwProfiler.vbwExecuteLine 2215
                    Idx = Evts(Eidx).Buttons(Bidx).Button
vbwProfiler.vbwExecuteLine 2216
                    .TextMatrix(Row, 2) = Idx
vbwProfiler.vbwExecuteLine 2217
                    .TextMatrix(Row, 3) = frmMain.Commands(Idx).Caption
vbwProfiler.vbwExecuteLine 2218
                    kb = Evts(Eidx).Buttons(Bidx).Enabled
vbwProfiler.vbwExecuteLine 2219
                    If kb <> "" Then
vbwProfiler.vbwExecuteLine 2220
                        If kb = "True" Then
vbwProfiler.vbwExecuteLine 2221
                            Call AddMessage(Row, "Enabled")
                        Else
vbwProfiler.vbwExecuteLine 2222 'B
vbwProfiler.vbwExecuteLine 2223
                            Call AddMessage(Row, "Disabled")
                        End If
vbwProfiler.vbwExecuteLine 2224 'B
                    End If
vbwProfiler.vbwExecuteLine 2225 'B
'                    .TextMatrix(Row, 5) = SignalAttributes(Evts(Eidx).Buttons(Bidx).signal).Name
vbwProfiler.vbwExecuteLine 2226
                Next Bidx
            End If
vbwProfiler.vbwExecuteLine 2227 'B
vbwProfiler.vbwExecuteLine 2228
            Idx = 0
'Me.Visible = True
vbwProfiler.vbwExecuteLine 2229
        Next Eidx
vbwProfiler.vbwExecuteLine 2230
    Me.Visible = True
vbwProfiler.vbwExecuteLine 2231
    .Col = 0
vbwProfiler.vbwExecuteLine 2232
    .Row = 0
vbwProfiler.vbwExecuteLine 2233
    .FocusRect = flexFocusNone ' (The selected cell changes)
vbwProfiler.vbwExecuteLine 2234
    End With
'Me.Visible = False
vbwProfiler.vbwExecuteLine 2235
frmEvents.WindowState = vbNormal  'Scale will be 0 in VBE (window is minimized)
vbwProfiler.vbwExecuteLine 2236
frmEvents.Refresh
vbwProfiler.vbwExecuteLine 2237
frmEvents.Visible = True
vbwProfiler.vbwProcOut 111
vbwProfiler.vbwExecuteLine 2238
End Function

Private Function AddMessage(Row, kb)
vbwProfiler.vbwProcIn 112
vbwProfiler.vbwExecuteLine 2239
    With mshEvents
vbwProfiler.vbwExecuteLine 2240
        If .TextMatrix(Row, 4) = "" Then
vbwProfiler.vbwExecuteLine 2241
            .TextMatrix(Row, 4) = kb
        Else
vbwProfiler.vbwExecuteLine 2242 'B
vbwProfiler.vbwExecuteLine 2243
            .TextMatrix(Row, 4) = .TextMatrix(Row, 4) & "," & kb
        End If
vbwProfiler.vbwExecuteLine 2244 'B
vbwProfiler.vbwExecuteLine 2245
    End With
vbwProfiler.vbwProcOut 112
vbwProfiler.vbwExecuteLine 2246
End Function

Private Sub Form_Resize()
vbwProfiler.vbwProcIn 113
vbwProfiler.vbwExecuteLine 2247
    mshEvents.Move 0, 0, ScaleWidth, ScaleHeight
vbwProfiler.vbwProcOut 113
vbwProfiler.vbwExecuteLine 2248
End Sub


Private Sub mshEvents_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
vbwProfiler.vbwProcIn 114
vbwProfiler.vbwExecuteLine 2249
    Call ActionEvent
vbwProfiler.vbwProcOut 114
vbwProfiler.vbwExecuteLine 2250
End Sub

Private Sub mshEvents_SelChange()
'    Call ActionEvent
vbwProfiler.vbwProcIn 115
vbwProfiler.vbwProcOut 115
vbwProfiler.vbwExecuteLine 2251
End Sub

Private Function ActionEvent()
vbwProfiler.vbwProcIn 116
Dim Idx As Long
Dim Index As Long
Dim ElapsedTime As Long
Dim arry() As String
Dim Minus As Boolean
Dim Sign As Integer
vbwProfiler.vbwExecuteLine 2252
    With mshEvents
vbwProfiler.vbwExecuteLine 2253
        If .Row < 1 Then 'Miss Header Row
vbwProfiler.vbwExecuteLine 2254
             .Row = 1
        End If
vbwProfiler.vbwExecuteLine 2255 'B
vbwProfiler.vbwExecuteLine 2256
            If .Row = 1 Then
vbwProfiler.vbwExecuteLine 2257
                Call frmMain.DefaultsStartTimeSet
vbwProfiler.vbwExecuteLine 2258
frmMain.RaceTimer.Enabled = False
            End If
vbwProfiler.vbwExecuteLine 2259 'B
'            Index = .TextMatrix(.Row, 0)
'                If .Col = 0 Then
'                    Call frmMain.DoEvent(Index)
'                End If
vbwProfiler.vbwExecuteLine 2260
            If .Col = 1 Then
vbwProfiler.vbwExecuteLine 2261
                If .TextMatrix(.Row, .Col) <> "" Then
vbwProfiler.vbwExecuteLine 2262
                    arry = Split(.TextMatrix(.Row, .Col), ":")
vbwProfiler.vbwExecuteLine 2263
                    Sign = Sgn(Replace(.TextMatrix(.Row, .Col), ":", ""))
vbwProfiler.vbwExecuteLine 2264
                    ElapsedTime = CLng(arry(0)) * 60 + CLng(arry(1)) * Sign
vbwProfiler.vbwExecuteLine 2265
                    Call frmMain.DoTimerEvents(ElapsedTime)
vbwProfiler.vbwExecuteLine 2266
If .TextMatrix(.Row, 3) = "Finish Enabled" Then
vbwProfiler.vbwExecuteLine 2267
     frmMain.RaceTimer.Enabled = True
End If
vbwProfiler.vbwExecuteLine 2268 'B
                End If
vbwProfiler.vbwExecuteLine 2269 'B
            End If
vbwProfiler.vbwExecuteLine 2270 'B
vbwProfiler.vbwExecuteLine 2271
        If .Row = .Rows - 1 Then
vbwProfiler.vbwExecuteLine 2272
            .Row = 0
vbwProfiler.vbwExecuteLine 2273
            .TopRow = 1
        End If
vbwProfiler.vbwExecuteLine 2274 'B
vbwProfiler.vbwExecuteLine 2275
    End With

vbwProfiler.vbwProcOut 116
vbwProfiler.vbwExecuteLine 2276
End Function




