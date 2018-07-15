VERSION 5.00
Begin VB.Form frmDpyBox 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Message"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5880
   Icon            =   "frmDpyBox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   5880
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer RefreshTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   600
      Top             =   0
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      Height          =   1215
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   4575
   End
   Begin VB.Timer HideMeTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "frmDpyBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const MIN_TEXT_HEIGHT = 780#
Const MIN_TEXT_WIDTH = 4680#
Const TEXTBOX_PADDING = 50# 'otherwise text runs into form border

 'http://vbnet.mvps.org/index.html?code/textapi/txscroll.htm
Private Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
  (ByVal hwnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long
Private Declare Function PutFocus Lib "user32" _
   Alias "SetFocus" _
  (ByVal hwnd As Long) As Long
Private Declare Function BringWindowToTop Lib "user32" (ByVal _
    hwnd As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" _
      (ByVal hwnd As Long) As Long

 Private Const EM_LINESCROLL = &HB6
Const EM_GETLINECOUNT = 186

'Dim MaxFrmHeight As Single
'Dim MaxFrmWidth As Single
Dim FrmBorderWidth As Single    'External=Internal size
Dim FrmBorderHeight As Single
Dim MaxTextHeight As Single
Dim MaxTextWidth As Single
'Dim FrmWidth As Single  'Used to calculate the Frm Size, before setting at the end
'Dim FrmHeight As Single
Dim MsgBuffer As String

Public Sub DpyBox(Message As String, Optional DisplaySecs As Long, Optional strCaption As String)
vbwProfiler.vbwProcIn 67
    Dim SavedWnd As Long

   'save the window handle of the control that currently has focus
vbwProfiler.vbwExecuteLine 1227
    On Error Resume Next    'may be no window
vbwProfiler.vbwExecuteLine 1228
    SavedWnd = Screen.ActiveControl.hwnd
vbwProfiler.vbwExecuteLine 1229
    On Error GoTo 0
'Buffer the message while the display timer is enabled
'Cant put into Text1 as it would force a refresh
vbwProfiler.vbwExecuteLine 1230
    If strCaption <> "" Then
vbwProfiler.vbwExecuteLine 1231
        Caption = strCaption
    Else
vbwProfiler.vbwExecuteLine 1232 'B
vbwProfiler.vbwExecuteLine 1233
        Caption = "Message"
    End If
vbwProfiler.vbwExecuteLine 1234 'B
vbwProfiler.vbwExecuteLine 1235
    If DisplaySecs = 0 Then
vbwProfiler.vbwExecuteLine 1236
         DisplaySecs = 5
    End If
vbwProfiler.vbwExecuteLine 1237 'B
vbwProfiler.vbwExecuteLine 1238
    MsgBuffer = MsgBuffer + Message
vbwProfiler.vbwExecuteLine 1239
    HideMeTimer.Interval = DisplaySecs * 1000
vbwProfiler.vbwExecuteLine 1240
    HideMeTimer.Enabled = True   'restart timer
vbwProfiler.vbwExecuteLine 1241
    If RefreshTimer.Enabled = True Then
vbwProfiler.vbwProcOut 67
vbwProfiler.vbwExecuteLine 1242
         Exit Sub
    End If
vbwProfiler.vbwExecuteLine 1243 'B
vbwProfiler.vbwExecuteLine 1244
    Text1.SelStart = Len(Text1.Text)    'causes less flicker than above
vbwProfiler.vbwExecuteLine 1245
    Text1.SelText = MsgBuffer
vbwProfiler.vbwExecuteLine 1246
    MsgBuffer = ""
vbwProfiler.vbwExecuteLine 1247
    Call RefreshDisplay
'To slow it down to debug RefreshTimer.Interval = 4000
vbwProfiler.vbwExecuteLine 1248
    RefreshTimer.Enabled = True 'comment out to debug
'    Call SetForegroundWindow(Me.hWnd)
vbwProfiler.vbwExecuteLine 1249
    On Error Resume Next    'may not be a window
vbwProfiler.vbwExecuteLine 1250
    Call PutFocus(SavedWnd)
vbwProfiler.vbwExecuteLine 1251
    On Error GoTo 0
vbwProfiler.vbwProcOut 67
vbwProfiler.vbwExecuteLine 1252
End Sub

Private Function RefreshDisplay()
vbwProfiler.vbwProcIn 68
Dim LineCount As Long
Dim ret As Long

vbwProfiler.vbwExecuteLine 1253
    On Error GoTo Error_RefreshDisplay
'    Text1 = Text1 & Message

'CANT SHOW if a MODAL form is currently displayed
vbwProfiler.vbwExecuteLine 1254
    Show

'MsgBox TextWidth(Text1)
vbwProfiler.vbwExecuteLine 1255
    Select Case TextWidth(Text1)
'vbwLine 1256:    Case Is < MIN_TEXT_WIDTH
    Case Is < IIf(vbwProfiler.vbwExecuteLine(1256), VBWPROFILER_EMPTY, _
        MIN_TEXT_WIDTH)
vbwProfiler.vbwExecuteLine 1257
        Text1.Width = MIN_TEXT_WIDTH
'vbwLine 1258:    Case Is > MaxTextWidth
    Case Is > IIf(vbwProfiler.vbwExecuteLine(1258), VBWPROFILER_EMPTY, _
        MaxTextWidth)
vbwProfiler.vbwExecuteLine 1259
        Text1.Width = MaxTextWidth
    Case Else
vbwProfiler.vbwExecuteLine 1260 'B
vbwProfiler.vbwExecuteLine 1261
        Text1.Width = TextWidth(Text1)
    End Select
vbwProfiler.vbwExecuteLine 1262 'B

'Make the textbox the maximum size so we can calculate the no of lines
vbwProfiler.vbwExecuteLine 1263
    Width = Text1.Width + FrmBorderWidth + TEXTBOX_PADDING * 2
vbwProfiler.vbwExecuteLine 1264
    Height = Text1.Height + FrmBorderHeight + TEXTBOX_PADDING * 2
vbwProfiler.vbwExecuteLine 1265
    Top = Screen.Height - Height
vbwProfiler.vbwExecuteLine 1266
    Left = Screen.Width - Width
'Remove all lines over 30
vbwProfiler.vbwExecuteLine 1267
    LineCount = GetLineCount(Text1)
'vbwLine 1268:    Do Until LineCount < 30
    Do Until vbwProfiler.vbwExecuteLine(1268) Or LineCount < 30
vbwProfiler.vbwExecuteLine 1269
        Text1 = Mid$(Text1, InStr(1, Text1, vbCrLf) + 2)
'        Height = Text1.Height + FrmBorderHeight + 100    '50 each side
vbwProfiler.vbwExecuteLine 1270
        LineCount = GetLineCount(Text1)
vbwProfiler.vbwExecuteLine 1271
    Loop
'We now have all the text we wish to display in the textbox
vbwProfiler.vbwExecuteLine 1272
    Select Case TextHeight(Text1)
'vbwLine 1273:    Case Is < MIN_TEXT_HEIGHT
    Case Is < IIf(vbwProfiler.vbwExecuteLine(1273), VBWPROFILER_EMPTY, _
        MIN_TEXT_HEIGHT)
vbwProfiler.vbwExecuteLine 1274
        Text1.Height = MIN_TEXT_HEIGHT
'vbwLine 1275:    Case Is > MaxTextHeight
    Case Is > IIf(vbwProfiler.vbwExecuteLine(1275), VBWPROFILER_EMPTY, _
        MaxTextHeight)
vbwProfiler.vbwExecuteLine 1276
        Text1.Height = MaxTextHeight
    Case Else
vbwProfiler.vbwExecuteLine 1277 'B
vbwProfiler.vbwExecuteLine 1278
        Text1.Height = TextHeight(Text1)
    End Select
vbwProfiler.vbwExecuteLine 1279 'B
vbwProfiler.vbwExecuteLine 1280
    Height = Text1.Height + FrmBorderHeight + TEXTBOX_PADDING * 2
vbwProfiler.vbwExecuteLine 1281
    Top = Screen.Height - Height
vbwProfiler.vbwExecuteLine 1282
    ret = SendMessage(Text1.hwnd, EM_LINESCROLL, 0, 100)
vbwProfiler.vbwProcOut 68
vbwProfiler.vbwExecuteLine 1283
    Exit Function
Error_RefreshDisplay:
vbwProfiler.vbwExecuteLine 1284
    Select Case Err.Number
'vbwLine 1285:    Case Is = 401 'cant show nonmodal when modal displayed
    Case Is = IIf(vbwProfiler.vbwExecuteLine(1285), VBWPROFILER_EMPTY, _
        401 )'cant show nonmodal when modal displayed
                    'retry until modal form is closed
'vbwLine 1286:    Case Is = 6 'overflow (text1.text too big)
    Case Is = IIf(vbwProfiler.vbwExecuteLine(1286), VBWPROFILER_EMPTY, _
        6 )'overflow (text1.text too big)
vbwProfiler.vbwExecuteLine 1287
        MsgBuffer = ""
vbwProfiler.vbwExecuteLine 1288
        Text1 = ""
vbwProfiler.vbwExecuteLine 1289
        Text1 = "RefreshDisplay Error " & Str(Err.Number) & " " & Err.Description & vbCrLf
    Case Else
vbwProfiler.vbwExecuteLine 1290 'B
vbwProfiler.vbwExecuteLine 1291
        Text1 = Text1 & "RefreshDisplay Error " & Str(Err.Number) & " " & Err.Description & vbCrLf
    End Select
vbwProfiler.vbwExecuteLine 1292 'B
vbwProfiler.vbwProcOut 68
vbwProfiler.vbwExecuteLine 1293
End Function

#If False Then
' Make the TextBox fit its contents.
'http://www.vb-helper.com/howto_size_textbox.html
Private Sub FitTextBoxContents(ByVal txt As Textbox)
    Font = Text1.Font
    txt.Width = TextWidth(txt.Text) + 120
    txt.Height = TextHeight(txt.Text) + 120
End Sub

'http://vbnet.mvps.org/index.html?code/textapi/txscroll.htm
Function ScrollText(Textbox As Control, vLines As Integer) As Long
Dim Success As Long
Dim SavedWnd As Long
Dim moveLines As Long
   'save the window handle of the control that currently has focus
    SavedWnd = Screen.ActiveControl.hwnd
    moveLines = vLines
   'Set the focus to the passed control (text control)
    Textbox.SetFocus
   'Scroll the lines.
    Success = SendMessage(Textbox.hwnd, EM_LINESCROLL, 0, ByVal moveLines)
   'Restore the focus to the original control
    Call PutFocus(SavedWnd)
   'Return the number of lines actually scrolled (INCORRECT)
    ScrollText = Success
End Function
#End If

Function GetLineCount(Textbox As Control) As Long
vbwProfiler.vbwProcIn 69
Dim lCount As Long
'The EM_GETLINECOUNT message retrieves the total number of text
'lines, not just the number of lines that are currently visible.
'If the Wordwrap feature is enabled, the number of lines can
'change when the dimensions of the editing window change.

vbwProfiler.vbwExecuteLine 1294
        lCount = SendMessage(Textbox.hwnd, EM_GETLINECOUNT, 0, 0)
vbwProfiler.vbwExecuteLine 1295
    GetLineCount = lCount
vbwProfiler.vbwProcOut 69
vbwProfiler.vbwExecuteLine 1296
End Function
    
Private Sub Form_Load()
'You must NOT set scale explicitly as we are calculationg the height
vbwProfiler.vbwProcIn 70

'The icon needs setting up from a file. If you try loading it
'at design time - it just keeps the file location. This will
'be different when a user tries.
'    Me.Icon = LoadPicture(NmeaRouterIcon)

vbwProfiler.vbwExecuteLine 1297
    FrmBorderWidth = Width - ScaleWidth
vbwProfiler.vbwExecuteLine 1298
    FrmBorderHeight = Height - ScaleHeight
vbwProfiler.vbwExecuteLine 1299
    Text1.Top = TEXTBOX_PADDING
vbwProfiler.vbwExecuteLine 1300
    Text1.Left = TEXTBOX_PADDING
vbwProfiler.vbwExecuteLine 1301
    MaxTextHeight = Screen.Height / 2 - FrmBorderHeight - TEXTBOX_PADDING * 2
vbwProfiler.vbwExecuteLine 1302
    MaxTextWidth = Screen.Width / 2 - FrmBorderWidth - TEXTBOX_PADDING * 2
vbwProfiler.vbwExecuteLine 1303
    If MaxTextWidth < MIN_TEXT_WIDTH Then
vbwProfiler.vbwExecuteLine 1304
         MaxTextWidth = MIN_TEXT_WIDTH
    End If
vbwProfiler.vbwExecuteLine 1305 'B
vbwProfiler.vbwExecuteLine 1306
    Me.BackColor = vbWhite
vbwProfiler.vbwExecuteLine 1307
    Text1.BorderStyle = vbBSNone    'must be set at design time
vbwProfiler.vbwProcOut 70
vbwProfiler.vbwExecuteLine 1308
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
vbwProfiler.vbwProcIn 71
vbwProfiler.vbwExecuteLine 1309
    If UnloadMode = vbFormControlMenu Then
vbwProfiler.vbwExecuteLine 1310
        Call HideMe
    End If
vbwProfiler.vbwExecuteLine 1311 'B
vbwProfiler.vbwProcOut 71
vbwProfiler.vbwExecuteLine 1312
End Sub

Private Sub RefreshTimer_Timer()
vbwProfiler.vbwProcIn 72
Dim SavedWnd As Long

   'save the window handle of the control that currently has focus
vbwProfiler.vbwExecuteLine 1313
    SavedWnd = Screen.ActiveControl.hwnd
vbwProfiler.vbwExecuteLine 1314
    RefreshTimer.Enabled = False
vbwProfiler.vbwExecuteLine 1315
    HideMeTimer.Enabled = True   'restart
vbwProfiler.vbwExecuteLine 1316
    Call RefreshDisplay
vbwProfiler.vbwExecuteLine 1317
    Call PutFocus(SavedWnd)
vbwProfiler.vbwProcOut 72
vbwProfiler.vbwExecuteLine 1318
End Sub

Private Sub HideMeTimer_Timer()
vbwProfiler.vbwProcIn 73
vbwProfiler.vbwExecuteLine 1319
    HideMeTimer.Enabled = False
'Dont hide if were just waiting for the next update
vbwProfiler.vbwExecuteLine 1320
    If RefreshTimer.Enabled = False Then
vbwProfiler.vbwExecuteLine 1321
        Call HideMe
    End If
vbwProfiler.vbwExecuteLine 1322 'B
vbwProfiler.vbwProcOut 73
vbwProfiler.vbwExecuteLine 1323
End Sub

Private Sub HideMe()
vbwProfiler.vbwProcIn 74
vbwProfiler.vbwExecuteLine 1324
    On Error GoTo Modal_Error
vbwProfiler.vbwExecuteLine 1325
    Me.Hide
vbwProfiler.vbwExecuteLine 1326
    MsgBuffer = ""
vbwProfiler.vbwExecuteLine 1327
    Text1 = ""
vbwProfiler.vbwProcOut 74
vbwProfiler.vbwExecuteLine 1328
    Exit Sub

Modal_Error:
vbwProfiler.vbwExecuteLine 1329
    HideMeTimer.Enabled = True  'retry
vbwProfiler.vbwProcOut 74
vbwProfiler.vbwExecuteLine 1330
End Sub




