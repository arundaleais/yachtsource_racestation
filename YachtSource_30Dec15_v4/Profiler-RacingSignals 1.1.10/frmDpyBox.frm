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
vbwProfiler.vbwProcIn 53
    Dim SavedWnd As Long

   'save the window handle of the control that currently has focus
vbwProfiler.vbwExecuteLine 915
    On Error Resume Next    'may be no window
vbwProfiler.vbwExecuteLine 916
    SavedWnd = Screen.ActiveControl.hwnd
vbwProfiler.vbwExecuteLine 917
    On Error GoTo 0
'Buffer the message while the display timer is enabled
'Cant put into Text1 as it would force a refresh
vbwProfiler.vbwExecuteLine 918
    If strCaption <> "" Then
vbwProfiler.vbwExecuteLine 919
        Caption = strCaption
    Else
vbwProfiler.vbwExecuteLine 920 'B
vbwProfiler.vbwExecuteLine 921
        Caption = "Message"
    End If
vbwProfiler.vbwExecuteLine 922 'B
vbwProfiler.vbwExecuteLine 923
    If DisplaySecs = 0 Then
vbwProfiler.vbwExecuteLine 924
         DisplaySecs = 5
    End If
vbwProfiler.vbwExecuteLine 925 'B
vbwProfiler.vbwExecuteLine 926
    MsgBuffer = MsgBuffer + Message
vbwProfiler.vbwExecuteLine 927
    HideMeTimer.Interval = DisplaySecs * 1000
vbwProfiler.vbwExecuteLine 928
    HideMeTimer.Enabled = True   'restart timer
vbwProfiler.vbwExecuteLine 929
    If RefreshTimer.Enabled = True Then
vbwProfiler.vbwProcOut 53
vbwProfiler.vbwExecuteLine 930
         Exit Sub
    End If
vbwProfiler.vbwExecuteLine 931 'B
vbwProfiler.vbwExecuteLine 932
    Text1.SelStart = Len(Text1.Text)    'causes less flicker than above
vbwProfiler.vbwExecuteLine 933
    Text1.SelText = MsgBuffer
vbwProfiler.vbwExecuteLine 934
    MsgBuffer = ""
vbwProfiler.vbwExecuteLine 935
    Call RefreshDisplay
'To slow it down to debug RefreshTimer.Interval = 4000
vbwProfiler.vbwExecuteLine 936
    RefreshTimer.Enabled = True 'comment out to debug
'    Call SetForegroundWindow(Me.hWnd)
vbwProfiler.vbwExecuteLine 937
    On Error Resume Next    'may not be a window
vbwProfiler.vbwExecuteLine 938
    Call PutFocus(SavedWnd)
vbwProfiler.vbwExecuteLine 939
    On Error GoTo 0
vbwProfiler.vbwProcOut 53
vbwProfiler.vbwExecuteLine 940
End Sub

Private Function RefreshDisplay()
vbwProfiler.vbwProcIn 54
Dim LineCount As Long
Dim ret As Long

vbwProfiler.vbwExecuteLine 941
    On Error GoTo Error_RefreshDisplay
'    Text1 = Text1 & Message

'CANT SHOW if a MODAL form is currently displayed
vbwProfiler.vbwExecuteLine 942
    Show

'MsgBox TextWidth(Text1)
vbwProfiler.vbwExecuteLine 943
    Select Case TextWidth(Text1)
'vbwLine 944:    Case Is < MIN_TEXT_WIDTH
    Case Is < IIf(vbwProfiler.vbwExecuteLine(944), VBWPROFILER_EMPTY, _
        MIN_TEXT_WIDTH)
vbwProfiler.vbwExecuteLine 945
        Text1.Width = MIN_TEXT_WIDTH
'vbwLine 946:    Case Is > MaxTextWidth
    Case Is > IIf(vbwProfiler.vbwExecuteLine(946), VBWPROFILER_EMPTY, _
        MaxTextWidth)
vbwProfiler.vbwExecuteLine 947
        Text1.Width = MaxTextWidth
    Case Else
vbwProfiler.vbwExecuteLine 948 'B
vbwProfiler.vbwExecuteLine 949
        Text1.Width = TextWidth(Text1)
    End Select
vbwProfiler.vbwExecuteLine 950 'B

'Make the textbox the maximum size so we can calculate the no of lines
vbwProfiler.vbwExecuteLine 951
    Width = Text1.Width + FrmBorderWidth + TEXTBOX_PADDING * 2
vbwProfiler.vbwExecuteLine 952
    Height = Text1.Height + FrmBorderHeight + TEXTBOX_PADDING * 2
vbwProfiler.vbwExecuteLine 953
    Top = Screen.Height - Height
vbwProfiler.vbwExecuteLine 954
    Left = Screen.Width - Width
'Remove all lines over 30
vbwProfiler.vbwExecuteLine 955
    LineCount = GetLineCount(Text1)
'vbwLine 956:    Do Until LineCount < 30
    Do Until vbwProfiler.vbwExecuteLine(956) Or LineCount < 30
vbwProfiler.vbwExecuteLine 957
        Text1 = Mid$(Text1, InStr(1, Text1, vbCrLf) + 2)
'        Height = Text1.Height + FrmBorderHeight + 100    '50 each side
vbwProfiler.vbwExecuteLine 958
        LineCount = GetLineCount(Text1)
vbwProfiler.vbwExecuteLine 959
    Loop
'We now have all the text we wish to display in the textbox
vbwProfiler.vbwExecuteLine 960
    Select Case TextHeight(Text1)
'vbwLine 961:    Case Is < MIN_TEXT_HEIGHT
    Case Is < IIf(vbwProfiler.vbwExecuteLine(961), VBWPROFILER_EMPTY, _
        MIN_TEXT_HEIGHT)
vbwProfiler.vbwExecuteLine 962
        Text1.Height = MIN_TEXT_HEIGHT
'vbwLine 963:    Case Is > MaxTextHeight
    Case Is > IIf(vbwProfiler.vbwExecuteLine(963), VBWPROFILER_EMPTY, _
        MaxTextHeight)
vbwProfiler.vbwExecuteLine 964
        Text1.Height = MaxTextHeight
    Case Else
vbwProfiler.vbwExecuteLine 965 'B
vbwProfiler.vbwExecuteLine 966
        Text1.Height = TextHeight(Text1)
    End Select
vbwProfiler.vbwExecuteLine 967 'B
vbwProfiler.vbwExecuteLine 968
    Height = Text1.Height + FrmBorderHeight + TEXTBOX_PADDING * 2
vbwProfiler.vbwExecuteLine 969
    Top = Screen.Height - Height
vbwProfiler.vbwExecuteLine 970
    ret = SendMessage(Text1.hwnd, EM_LINESCROLL, 0, 100)
vbwProfiler.vbwProcOut 54
vbwProfiler.vbwExecuteLine 971
    Exit Function
Error_RefreshDisplay:
vbwProfiler.vbwExecuteLine 972
    Select Case Err.Number
'vbwLine 973:    Case Is = 401 'cant show nonmodal when modal displayed
    Case Is = IIf(vbwProfiler.vbwExecuteLine(973), VBWPROFILER_EMPTY, _
        401 )'cant show nonmodal when modal displayed
                    'retry until modal form is closed
'vbwLine 974:    Case Is = 6 'overflow (text1.text too big)
    Case Is = IIf(vbwProfiler.vbwExecuteLine(974), VBWPROFILER_EMPTY, _
        6 )'overflow (text1.text too big)
vbwProfiler.vbwExecuteLine 975
        MsgBuffer = ""
vbwProfiler.vbwExecuteLine 976
        Text1 = ""
vbwProfiler.vbwExecuteLine 977
        Text1 = "RefreshDisplay Error " & Str(Err.Number) & " " & Err.Description & vbCrLf
    Case Else
vbwProfiler.vbwExecuteLine 978 'B
vbwProfiler.vbwExecuteLine 979
        Text1 = Text1 & "RefreshDisplay Error " & Str(Err.Number) & " " & Err.Description & vbCrLf
    End Select
vbwProfiler.vbwExecuteLine 980 'B
vbwProfiler.vbwProcOut 54
vbwProfiler.vbwExecuteLine 981
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
vbwProfiler.vbwProcIn 55
Dim lCount As Long
'The EM_GETLINECOUNT message retrieves the total number of text
'lines, not just the number of lines that are currently visible.
'If the Wordwrap feature is enabled, the number of lines can
'change when the dimensions of the editing window change.

vbwProfiler.vbwExecuteLine 982
        lCount = SendMessage(Textbox.hwnd, EM_GETLINECOUNT, 0, 0)
vbwProfiler.vbwExecuteLine 983
    GetLineCount = lCount
vbwProfiler.vbwProcOut 55
vbwProfiler.vbwExecuteLine 984
End Function
    
Private Sub Form_Load()
'You must NOT set scale explicitly as we are calculationg the height
vbwProfiler.vbwProcIn 56

'The icon needs setting up from a file. If you try loading it
'at design time - it just keeps the file location. This will
'be different when a user tries.
'    Me.Icon = LoadPicture(NmeaRouterIcon)

vbwProfiler.vbwExecuteLine 985
    FrmBorderWidth = Width - ScaleWidth
vbwProfiler.vbwExecuteLine 986
    FrmBorderHeight = Height - ScaleHeight
vbwProfiler.vbwExecuteLine 987
    Text1.Top = TEXTBOX_PADDING
vbwProfiler.vbwExecuteLine 988
    Text1.Left = TEXTBOX_PADDING
vbwProfiler.vbwExecuteLine 989
    MaxTextHeight = Screen.Height / 2 - FrmBorderHeight - TEXTBOX_PADDING * 2
vbwProfiler.vbwExecuteLine 990
    MaxTextWidth = Screen.Width / 2 - FrmBorderWidth - TEXTBOX_PADDING * 2
vbwProfiler.vbwExecuteLine 991
    If MaxTextWidth < MIN_TEXT_WIDTH Then
vbwProfiler.vbwExecuteLine 992
         MaxTextWidth = MIN_TEXT_WIDTH
    End If
vbwProfiler.vbwExecuteLine 993 'B
vbwProfiler.vbwExecuteLine 994
    Me.BackColor = vbWhite
vbwProfiler.vbwExecuteLine 995
    Text1.BorderStyle = vbBSNone    'must be set at design time
vbwProfiler.vbwProcOut 56
vbwProfiler.vbwExecuteLine 996
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
vbwProfiler.vbwProcIn 57
vbwProfiler.vbwExecuteLine 997
    If UnloadMode = vbFormControlMenu Then
vbwProfiler.vbwExecuteLine 998
        Call HideMe
    End If
vbwProfiler.vbwExecuteLine 999 'B
vbwProfiler.vbwProcOut 57
vbwProfiler.vbwExecuteLine 1000
End Sub

Private Sub RefreshTimer_Timer()
vbwProfiler.vbwProcIn 58
Dim SavedWnd As Long

   'save the window handle of the control that currently has focus
vbwProfiler.vbwExecuteLine 1001
    SavedWnd = Screen.ActiveControl.hwnd
vbwProfiler.vbwExecuteLine 1002
    RefreshTimer.Enabled = False
vbwProfiler.vbwExecuteLine 1003
    HideMeTimer.Enabled = True   'restart
vbwProfiler.vbwExecuteLine 1004
    Call RefreshDisplay
vbwProfiler.vbwExecuteLine 1005
    Call PutFocus(SavedWnd)
vbwProfiler.vbwProcOut 58
vbwProfiler.vbwExecuteLine 1006
End Sub

Private Sub HideMeTimer_Timer()
vbwProfiler.vbwProcIn 59
vbwProfiler.vbwExecuteLine 1007
    HideMeTimer.Enabled = False
'Dont hide if were just waiting for the next update
vbwProfiler.vbwExecuteLine 1008
    If RefreshTimer.Enabled = False Then
vbwProfiler.vbwExecuteLine 1009
        Call HideMe
    End If
vbwProfiler.vbwExecuteLine 1010 'B
vbwProfiler.vbwProcOut 59
vbwProfiler.vbwExecuteLine 1011
End Sub

Private Sub HideMe()
vbwProfiler.vbwProcIn 60
vbwProfiler.vbwExecuteLine 1012
    On Error GoTo Modal_Error
vbwProfiler.vbwExecuteLine 1013
    Me.Hide
vbwProfiler.vbwExecuteLine 1014
    MsgBuffer = ""
vbwProfiler.vbwExecuteLine 1015
    Text1 = ""
vbwProfiler.vbwProcOut 60
vbwProfiler.vbwExecuteLine 1016
    Exit Sub

Modal_Error:
vbwProfiler.vbwExecuteLine 1017
    HideMeTimer.Enabled = True  'retry
vbwProfiler.vbwProcOut 60
vbwProfiler.vbwExecuteLine 1018
End Sub




