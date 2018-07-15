VERSION 5.00
Begin VB.Form frmDpyBox 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Message"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4830
   Icon            =   "frmDpyBox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   4830
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
vbwProfiler.vbwProcIn 54
    Dim SavedWnd As Long

   'save the window handle of the control that currently has focus
vbwProfiler.vbwExecuteLine 953
    On Error Resume Next    'may be no window
vbwProfiler.vbwExecuteLine 954
    SavedWnd = Screen.ActiveControl.hwnd
vbwProfiler.vbwExecuteLine 955
    On Error GoTo 0
'Buffer the message while the display timer is enabled
'Cant put into Text1 as it would force a refresh
vbwProfiler.vbwExecuteLine 956
    If strCaption <> "" Then
vbwProfiler.vbwExecuteLine 957
        Caption = strCaption
    Else
vbwProfiler.vbwExecuteLine 958 'B
vbwProfiler.vbwExecuteLine 959
        Caption = "Message"
    End If
vbwProfiler.vbwExecuteLine 960 'B
vbwProfiler.vbwExecuteLine 961
    If DisplaySecs = 0 Then
vbwProfiler.vbwExecuteLine 962
         DisplaySecs = 5
    End If
vbwProfiler.vbwExecuteLine 963 'B
vbwProfiler.vbwExecuteLine 964
    MsgBuffer = MsgBuffer + Message
vbwProfiler.vbwExecuteLine 965
    HideMeTimer.Interval = DisplaySecs * 1000
vbwProfiler.vbwExecuteLine 966
    HideMeTimer.Enabled = True   'restart timer
vbwProfiler.vbwExecuteLine 967
    If RefreshTimer.Enabled = True Then
vbwProfiler.vbwProcOut 54
vbwProfiler.vbwExecuteLine 968
         Exit Sub
    End If
vbwProfiler.vbwExecuteLine 969 'B
vbwProfiler.vbwExecuteLine 970
    Text1.SelStart = Len(Text1.Text)    'causes less flicker than above
vbwProfiler.vbwExecuteLine 971
    Text1.SelText = MsgBuffer
vbwProfiler.vbwExecuteLine 972
    MsgBuffer = ""
vbwProfiler.vbwExecuteLine 973
    Call RefreshDisplay
'To slow it down to debug RefreshTimer.Interval = 4000
vbwProfiler.vbwExecuteLine 974
    RefreshTimer.Enabled = True 'comment out to debug
'    Call SetForegroundWindow(Me.hWnd)
vbwProfiler.vbwExecuteLine 975
    On Error Resume Next    'may not be a window
vbwProfiler.vbwExecuteLine 976
    Call PutFocus(SavedWnd)
vbwProfiler.vbwExecuteLine 977
    On Error GoTo 0
vbwProfiler.vbwProcOut 54
vbwProfiler.vbwExecuteLine 978
End Sub

Private Function RefreshDisplay()
vbwProfiler.vbwProcIn 55
Dim LineCount As Long
Dim ret As Long

vbwProfiler.vbwExecuteLine 979
    On Error GoTo Error_RefreshDisplay
'    Text1 = Text1 & Message

'CANT SHOW if a MODAL form is currently displayed
vbwProfiler.vbwExecuteLine 980
    Show

'MsgBox TextWidth(Text1)
vbwProfiler.vbwExecuteLine 981
    Select Case TextWidth(Text1)
'vbwLine 982:    Case Is < MIN_TEXT_WIDTH
    Case Is < IIf(vbwProfiler.vbwExecuteLine(982), VBWPROFILER_EMPTY, _
        MIN_TEXT_WIDTH)
vbwProfiler.vbwExecuteLine 983
        Text1.Width = MIN_TEXT_WIDTH
'vbwLine 984:    Case Is > MaxTextWidth
    Case Is > IIf(vbwProfiler.vbwExecuteLine(984), VBWPROFILER_EMPTY, _
        MaxTextWidth)
vbwProfiler.vbwExecuteLine 985
        Text1.Width = MaxTextWidth
    Case Else
vbwProfiler.vbwExecuteLine 986 'B
vbwProfiler.vbwExecuteLine 987
        Text1.Width = TextWidth(Text1)
    End Select
vbwProfiler.vbwExecuteLine 988 'B

'Make the textbox the maximum size so we can calculate the no of lines
vbwProfiler.vbwExecuteLine 989
    Width = Text1.Width + FrmBorderWidth + TEXTBOX_PADDING * 2
vbwProfiler.vbwExecuteLine 990
    Height = Text1.Height + FrmBorderHeight + TEXTBOX_PADDING * 2
vbwProfiler.vbwExecuteLine 991
    Top = Screen.Height - Height
vbwProfiler.vbwExecuteLine 992
    Left = Screen.Width - Width
'Remove all lines over 30
vbwProfiler.vbwExecuteLine 993
    LineCount = GetLineCount(Text1)
'vbwLine 994:    Do Until LineCount < 30
    Do Until vbwProfiler.vbwExecuteLine(994) Or LineCount < 30
vbwProfiler.vbwExecuteLine 995
        Text1 = Mid$(Text1, InStr(1, Text1, vbCrLf) + 2)
'        Height = Text1.Height + FrmBorderHeight + 100    '50 each side
vbwProfiler.vbwExecuteLine 996
        LineCount = GetLineCount(Text1)
vbwProfiler.vbwExecuteLine 997
    Loop
'We now have all the text we wish to display in the textbox
vbwProfiler.vbwExecuteLine 998
    Select Case TextHeight(Text1)
'vbwLine 999:    Case Is < MIN_TEXT_HEIGHT
    Case Is < IIf(vbwProfiler.vbwExecuteLine(999), VBWPROFILER_EMPTY, _
        MIN_TEXT_HEIGHT)
vbwProfiler.vbwExecuteLine 1000
        Text1.Height = MIN_TEXT_HEIGHT
'vbwLine 1001:    Case Is > MaxTextHeight
    Case Is > IIf(vbwProfiler.vbwExecuteLine(1001), VBWPROFILER_EMPTY, _
        MaxTextHeight)
vbwProfiler.vbwExecuteLine 1002
        Text1.Height = MaxTextHeight
    Case Else
vbwProfiler.vbwExecuteLine 1003 'B
vbwProfiler.vbwExecuteLine 1004
        Text1.Height = TextHeight(Text1)
    End Select
vbwProfiler.vbwExecuteLine 1005 'B
vbwProfiler.vbwExecuteLine 1006
    Height = Text1.Height + FrmBorderHeight + TEXTBOX_PADDING * 2
vbwProfiler.vbwExecuteLine 1007
    Top = Screen.Height - Height
vbwProfiler.vbwExecuteLine 1008
    ret = SendMessage(Text1.hwnd, EM_LINESCROLL, 0, 100)
vbwProfiler.vbwProcOut 55
vbwProfiler.vbwExecuteLine 1009
    Exit Function
Error_RefreshDisplay:
vbwProfiler.vbwExecuteLine 1010
    Select Case Err.Number
'vbwLine 1011:    Case Is = 401 'cant show nonmodal when modal displayed
    Case Is = IIf(vbwProfiler.vbwExecuteLine(1011), VBWPROFILER_EMPTY, _
        401 )'cant show nonmodal when modal displayed
                    'retry until modal form is closed
'vbwLine 1012:    Case Is = 6 'overflow (text1.text too big)
    Case Is = IIf(vbwProfiler.vbwExecuteLine(1012), VBWPROFILER_EMPTY, _
        6 )'overflow (text1.text too big)
vbwProfiler.vbwExecuteLine 1013
        MsgBuffer = ""
vbwProfiler.vbwExecuteLine 1014
        Text1 = ""
vbwProfiler.vbwExecuteLine 1015
        Text1 = "RefreshDisplay Error " & Str(Err.Number) & " " & Err.Description & vbCrLf
    Case Else
vbwProfiler.vbwExecuteLine 1016 'B
vbwProfiler.vbwExecuteLine 1017
        Text1 = Text1 & "RefreshDisplay Error " & Str(Err.Number) & " " & Err.Description & vbCrLf
    End Select
vbwProfiler.vbwExecuteLine 1018 'B
vbwProfiler.vbwProcOut 55
vbwProfiler.vbwExecuteLine 1019
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
vbwProfiler.vbwProcIn 56
Dim lCount As Long
'The EM_GETLINECOUNT message retrieves the total number of text
'lines, not just the number of lines that are currently visible.
'If the Wordwrap feature is enabled, the number of lines can
'change when the dimensions of the editing window change.

vbwProfiler.vbwExecuteLine 1020
        lCount = SendMessage(Textbox.hwnd, EM_GETLINECOUNT, 0, 0)
vbwProfiler.vbwExecuteLine 1021
    GetLineCount = lCount
vbwProfiler.vbwProcOut 56
vbwProfiler.vbwExecuteLine 1022
End Function
    
Private Sub Form_Load()
'You must NOT set scale explicitly as we are calculationg the height
vbwProfiler.vbwProcIn 57

'The icon needs setting up from a file. If you try loading it
'at design time - it just keeps the file location. This will
'be different when a user tries.
'    Me.Icon = LoadPicture(NmeaRouterIcon)

vbwProfiler.vbwExecuteLine 1023
    FrmBorderWidth = Width - ScaleWidth
vbwProfiler.vbwExecuteLine 1024
    FrmBorderHeight = Height - ScaleHeight
vbwProfiler.vbwExecuteLine 1025
    Text1.Top = TEXTBOX_PADDING
vbwProfiler.vbwExecuteLine 1026
    Text1.Left = TEXTBOX_PADDING
vbwProfiler.vbwExecuteLine 1027
    MaxTextHeight = Screen.Height / 2 - FrmBorderHeight - TEXTBOX_PADDING * 2
vbwProfiler.vbwExecuteLine 1028
    MaxTextWidth = Screen.Width / 2 - FrmBorderWidth - TEXTBOX_PADDING * 2
vbwProfiler.vbwExecuteLine 1029
    If MaxTextWidth < MIN_TEXT_WIDTH Then
vbwProfiler.vbwExecuteLine 1030
         MaxTextWidth = MIN_TEXT_WIDTH
    End If
vbwProfiler.vbwExecuteLine 1031 'B
vbwProfiler.vbwExecuteLine 1032
    Me.BackColor = vbWhite
vbwProfiler.vbwExecuteLine 1033
    Text1.BorderStyle = vbBSNone    'must be set at design time
vbwProfiler.vbwProcOut 57
vbwProfiler.vbwExecuteLine 1034
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
vbwProfiler.vbwProcIn 58
vbwProfiler.vbwExecuteLine 1035
    If UnloadMode = vbFormControlMenu Then
vbwProfiler.vbwExecuteLine 1036
        Call HideMe
    End If
vbwProfiler.vbwExecuteLine 1037 'B
vbwProfiler.vbwProcOut 58
vbwProfiler.vbwExecuteLine 1038
End Sub

Private Sub RefreshTimer_Timer()
vbwProfiler.vbwProcIn 59
Dim SavedWnd As Long

   'save the window handle of the control that currently has focus
vbwProfiler.vbwExecuteLine 1039
    SavedWnd = Screen.ActiveControl.hwnd
vbwProfiler.vbwExecuteLine 1040
    RefreshTimer.Enabled = False
vbwProfiler.vbwExecuteLine 1041
    HideMeTimer.Enabled = True   'restart
vbwProfiler.vbwExecuteLine 1042
    Call RefreshDisplay
vbwProfiler.vbwExecuteLine 1043
    Call PutFocus(SavedWnd)
vbwProfiler.vbwProcOut 59
vbwProfiler.vbwExecuteLine 1044
End Sub

Private Sub HideMeTimer_Timer()
vbwProfiler.vbwProcIn 60
vbwProfiler.vbwExecuteLine 1045
    HideMeTimer.Enabled = False
'Dont hide if were just waiting for the next update
vbwProfiler.vbwExecuteLine 1046
    If RefreshTimer.Enabled = False Then
vbwProfiler.vbwExecuteLine 1047
        Call HideMe
    End If
vbwProfiler.vbwExecuteLine 1048 'B
vbwProfiler.vbwProcOut 60
vbwProfiler.vbwExecuteLine 1049
End Sub

Private Sub HideMe()
vbwProfiler.vbwProcIn 61
vbwProfiler.vbwExecuteLine 1050
    On Error GoTo Modal_Error
vbwProfiler.vbwExecuteLine 1051
    Me.Hide
vbwProfiler.vbwExecuteLine 1052
    MsgBuffer = ""
vbwProfiler.vbwExecuteLine 1053
    Text1 = ""
vbwProfiler.vbwProcOut 61
vbwProfiler.vbwExecuteLine 1054
    Exit Sub

Modal_Error:
vbwProfiler.vbwExecuteLine 1055
    HideMeTimer.Enabled = True  'retry
vbwProfiler.vbwProcOut 61
vbwProfiler.vbwExecuteLine 1056
End Sub




