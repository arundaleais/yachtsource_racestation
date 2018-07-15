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
vbwProfiler.vbwProcIn 47
    Dim SavedWnd As Long

   'save the window handle of the control that currently has focus
vbwProfiler.vbwExecuteLine 779
    On Error Resume Next    'may be no window
vbwProfiler.vbwExecuteLine 780
    SavedWnd = Screen.ActiveControl.hwnd
vbwProfiler.vbwExecuteLine 781
    On Error GoTo 0
'Buffer the message while the display timer is enabled
'Cant put into Text1 as it would force a refresh
vbwProfiler.vbwExecuteLine 782
    If strCaption <> "" Then
vbwProfiler.vbwExecuteLine 783
        Caption = strCaption
    Else
vbwProfiler.vbwExecuteLine 784 'B
vbwProfiler.vbwExecuteLine 785
        Caption = "Message"
    End If
vbwProfiler.vbwExecuteLine 786 'B
vbwProfiler.vbwExecuteLine 787
    If DisplaySecs = 0 Then
vbwProfiler.vbwExecuteLine 788
         DisplaySecs = 5
    End If
vbwProfiler.vbwExecuteLine 789 'B
vbwProfiler.vbwExecuteLine 790
    MsgBuffer = MsgBuffer + Message
vbwProfiler.vbwExecuteLine 791
    HideMeTimer.Interval = DisplaySecs * 1000
vbwProfiler.vbwExecuteLine 792
    HideMeTimer.Enabled = True   'restart timer
vbwProfiler.vbwExecuteLine 793
    If RefreshTimer.Enabled = True Then
vbwProfiler.vbwProcOut 47
vbwProfiler.vbwExecuteLine 794
         Exit Sub
    End If
vbwProfiler.vbwExecuteLine 795 'B
vbwProfiler.vbwExecuteLine 796
    Text1.SelStart = Len(Text1.Text)    'causes less flicker than above
vbwProfiler.vbwExecuteLine 797
    Text1.SelText = MsgBuffer
vbwProfiler.vbwExecuteLine 798
    MsgBuffer = ""
vbwProfiler.vbwExecuteLine 799
    Call RefreshDisplay
'To slow it down to debug RefreshTimer.Interval = 4000
vbwProfiler.vbwExecuteLine 800
    RefreshTimer.Enabled = True 'comment out to debug
'    Call SetForegroundWindow(Me.hWnd)
vbwProfiler.vbwExecuteLine 801
    On Error Resume Next    'may not be a window
vbwProfiler.vbwExecuteLine 802
    Call PutFocus(SavedWnd)
vbwProfiler.vbwExecuteLine 803
    On Error GoTo 0
vbwProfiler.vbwProcOut 47
vbwProfiler.vbwExecuteLine 804
End Sub

Private Function RefreshDisplay()
vbwProfiler.vbwProcIn 48
Dim LineCount As Long
Dim ret As Long

vbwProfiler.vbwExecuteLine 805
    On Error GoTo Error_RefreshDisplay
'    Text1 = Text1 & Message

'CANT SHOW if a MODAL form is currently displayed
vbwProfiler.vbwExecuteLine 806
    Show

'MsgBox TextWidth(Text1)
vbwProfiler.vbwExecuteLine 807
    Select Case TextWidth(Text1)
'vbwLine 808:    Case Is < MIN_TEXT_WIDTH
    Case Is < IIf(vbwProfiler.vbwExecuteLine(808), VBWPROFILER_EMPTY, _
        MIN_TEXT_WIDTH)
vbwProfiler.vbwExecuteLine 809
        Text1.Width = MIN_TEXT_WIDTH
'vbwLine 810:    Case Is > MaxTextWidth
    Case Is > IIf(vbwProfiler.vbwExecuteLine(810), VBWPROFILER_EMPTY, _
        MaxTextWidth)
vbwProfiler.vbwExecuteLine 811
        Text1.Width = MaxTextWidth
    Case Else
vbwProfiler.vbwExecuteLine 812 'B
vbwProfiler.vbwExecuteLine 813
        Text1.Width = TextWidth(Text1)
    End Select
vbwProfiler.vbwExecuteLine 814 'B

'Make the textbox the maximum size so we can calculate the no of lines
vbwProfiler.vbwExecuteLine 815
    Width = Text1.Width + FrmBorderWidth + TEXTBOX_PADDING * 2
vbwProfiler.vbwExecuteLine 816
    Height = Text1.Height + FrmBorderHeight + TEXTBOX_PADDING * 2
vbwProfiler.vbwExecuteLine 817
    Top = Screen.Height - Height
vbwProfiler.vbwExecuteLine 818
    Left = Screen.Width - Width
'Remove all lines over 30
vbwProfiler.vbwExecuteLine 819
    LineCount = GetLineCount(Text1)
'vbwLine 820:    Do Until LineCount < 30
    Do Until vbwProfiler.vbwExecuteLine(820) Or LineCount < 30
vbwProfiler.vbwExecuteLine 821
        Text1 = Mid$(Text1, InStr(1, Text1, vbCrLf) + 2)
'        Height = Text1.Height + FrmBorderHeight + 100    '50 each side
vbwProfiler.vbwExecuteLine 822
        LineCount = GetLineCount(Text1)
vbwProfiler.vbwExecuteLine 823
    Loop
'We now have all the text we wish to display in the textbox
vbwProfiler.vbwExecuteLine 824
    Select Case TextHeight(Text1)
'vbwLine 825:    Case Is < MIN_TEXT_HEIGHT
    Case Is < IIf(vbwProfiler.vbwExecuteLine(825), VBWPROFILER_EMPTY, _
        MIN_TEXT_HEIGHT)
vbwProfiler.vbwExecuteLine 826
        Text1.Height = MIN_TEXT_HEIGHT
'vbwLine 827:    Case Is > MaxTextHeight
    Case Is > IIf(vbwProfiler.vbwExecuteLine(827), VBWPROFILER_EMPTY, _
        MaxTextHeight)
vbwProfiler.vbwExecuteLine 828
        Text1.Height = MaxTextHeight
    Case Else
vbwProfiler.vbwExecuteLine 829 'B
vbwProfiler.vbwExecuteLine 830
        Text1.Height = TextHeight(Text1)
    End Select
vbwProfiler.vbwExecuteLine 831 'B
vbwProfiler.vbwExecuteLine 832
    Height = Text1.Height + FrmBorderHeight + TEXTBOX_PADDING * 2
vbwProfiler.vbwExecuteLine 833
    Top = Screen.Height - Height
vbwProfiler.vbwExecuteLine 834
    ret = SendMessage(Text1.hwnd, EM_LINESCROLL, 0, 100)
vbwProfiler.vbwProcOut 48
vbwProfiler.vbwExecuteLine 835
    Exit Function
Error_RefreshDisplay:
vbwProfiler.vbwExecuteLine 836
    Select Case Err.Number
'vbwLine 837:    Case Is = 401 'cant show nonmodal when modal displayed
    Case Is = IIf(vbwProfiler.vbwExecuteLine(837), VBWPROFILER_EMPTY, _
        401 )'cant show nonmodal when modal displayed
                    'retry until modal form is closed
'vbwLine 838:    Case Is = 6 'overflow (text1.text too big)
    Case Is = IIf(vbwProfiler.vbwExecuteLine(838), VBWPROFILER_EMPTY, _
        6 )'overflow (text1.text too big)
vbwProfiler.vbwExecuteLine 839
        MsgBuffer = ""
vbwProfiler.vbwExecuteLine 840
        Text1 = ""
vbwProfiler.vbwExecuteLine 841
        Text1 = "RefreshDisplay Error " & Str(Err.Number) & " " & Err.Description & vbCrLf
    Case Else
vbwProfiler.vbwExecuteLine 842 'B
vbwProfiler.vbwExecuteLine 843
        Text1 = Text1 & "RefreshDisplay Error " & Str(Err.Number) & " " & Err.Description & vbCrLf
    End Select
vbwProfiler.vbwExecuteLine 844 'B
vbwProfiler.vbwProcOut 48
vbwProfiler.vbwExecuteLine 845
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
vbwProfiler.vbwProcIn 49
Dim lCount As Long
'The EM_GETLINECOUNT message retrieves the total number of text
'lines, not just the number of lines that are currently visible.
'If the Wordwrap feature is enabled, the number of lines can
'change when the dimensions of the editing window change.

vbwProfiler.vbwExecuteLine 846
        lCount = SendMessage(Textbox.hwnd, EM_GETLINECOUNT, 0, 0)
vbwProfiler.vbwExecuteLine 847
    GetLineCount = lCount
vbwProfiler.vbwProcOut 49
vbwProfiler.vbwExecuteLine 848
End Function
    
Private Sub Form_Load()
'You must NOT set scale explicitly as we are calculationg the height
vbwProfiler.vbwProcIn 50

'The icon needs setting up from a file. If you try loading it
'at design time - it just keeps the file location. This will
'be different when a user tries.
'    Me.Icon = LoadPicture(NmeaRouterIcon)

vbwProfiler.vbwExecuteLine 849
    FrmBorderWidth = Width - ScaleWidth
vbwProfiler.vbwExecuteLine 850
    FrmBorderHeight = Height - ScaleHeight
vbwProfiler.vbwExecuteLine 851
    Text1.Top = TEXTBOX_PADDING
vbwProfiler.vbwExecuteLine 852
    Text1.Left = TEXTBOX_PADDING
vbwProfiler.vbwExecuteLine 853
    MaxTextHeight = Screen.Height / 2 - FrmBorderHeight - TEXTBOX_PADDING * 2
vbwProfiler.vbwExecuteLine 854
    MaxTextWidth = Screen.Width / 2 - FrmBorderWidth - TEXTBOX_PADDING * 2
vbwProfiler.vbwExecuteLine 855
    If MaxTextWidth < MIN_TEXT_WIDTH Then
vbwProfiler.vbwExecuteLine 856
         MaxTextWidth = MIN_TEXT_WIDTH
    End If
vbwProfiler.vbwExecuteLine 857 'B
vbwProfiler.vbwExecuteLine 858
    Me.BackColor = vbWhite
vbwProfiler.vbwExecuteLine 859
    Text1.BorderStyle = vbBSNone    'must be set at design time
vbwProfiler.vbwProcOut 50
vbwProfiler.vbwExecuteLine 860
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
vbwProfiler.vbwProcIn 51
vbwProfiler.vbwExecuteLine 861
    If UnloadMode = vbFormControlMenu Then
vbwProfiler.vbwExecuteLine 862
        Call HideMe
    End If
vbwProfiler.vbwExecuteLine 863 'B
vbwProfiler.vbwProcOut 51
vbwProfiler.vbwExecuteLine 864
End Sub

Private Sub RefreshTimer_Timer()
vbwProfiler.vbwProcIn 52
Dim SavedWnd As Long

   'save the window handle of the control that currently has focus
vbwProfiler.vbwExecuteLine 865
    SavedWnd = Screen.ActiveControl.hwnd
vbwProfiler.vbwExecuteLine 866
    RefreshTimer.Enabled = False
vbwProfiler.vbwExecuteLine 867
    HideMeTimer.Enabled = True   'restart
vbwProfiler.vbwExecuteLine 868
    Call RefreshDisplay
vbwProfiler.vbwExecuteLine 869
    Call PutFocus(SavedWnd)
vbwProfiler.vbwProcOut 52
vbwProfiler.vbwExecuteLine 870
End Sub

Private Sub HideMeTimer_Timer()
vbwProfiler.vbwProcIn 53
vbwProfiler.vbwExecuteLine 871
    HideMeTimer.Enabled = False
'Dont hide if were just waiting for the next update
vbwProfiler.vbwExecuteLine 872
    If RefreshTimer.Enabled = False Then
vbwProfiler.vbwExecuteLine 873
        Call HideMe
    End If
vbwProfiler.vbwExecuteLine 874 'B
vbwProfiler.vbwProcOut 53
vbwProfiler.vbwExecuteLine 875
End Sub

Private Sub HideMe()
vbwProfiler.vbwProcIn 54
vbwProfiler.vbwExecuteLine 876
    On Error GoTo Modal_Error
vbwProfiler.vbwExecuteLine 877
    Me.Hide
vbwProfiler.vbwExecuteLine 878
    MsgBuffer = ""
vbwProfiler.vbwExecuteLine 879
    Text1 = ""
vbwProfiler.vbwProcOut 54
vbwProfiler.vbwExecuteLine 880
    Exit Sub

Modal_Error:
vbwProfiler.vbwExecuteLine 881
    HideMeTimer.Enabled = True  'retry
vbwProfiler.vbwProcOut 54
vbwProfiler.vbwExecuteLine 882
End Sub




