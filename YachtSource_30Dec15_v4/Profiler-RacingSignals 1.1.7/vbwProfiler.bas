Attribute VB_Name = "vbwProfiler"
' vbwNoProfileProc
' vbwNoProfileLine

Option Explicit

Public VBWPROFILER_EMPTY As Variant ' for use with vbwExecuteLine() in IIf structures

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Private Const vbMsgBoxSetTopMost = &H40000
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long

Dim ProfilerHW As Long
Dim ProfilerID As Long

' messages and udt's used to communicate with the profiler
Private Const WM_USER = &H400
Private Const WM_USER_PROFILE_DECLARE = WM_USER + 1
Private Const WM_USER_PROFILE_EXECUTE_LINE = WM_USER + 2
Private Const WM_USER_PROFILE_ENTER_PROCEDURE = WM_USER + 3
Private Const WM_USER_PROFILE_EXIT_PROCEDURE = WM_USER + 4
Private Const WM_USER_PROFILE_OVERHEAD = WM_USER + 5
Private Const WM_USER_PROFILE_START_OVERHEAD = WM_USER + 6
Private Const WM_USER_PROFILE_END_OVERHEAD = WM_USER + 7
Private Const WM_USER_PROFILE_END_SESSION = WM_USER + 8

Private Type ProfDeclareSession
    ThreadID As Long
    ProcessID As Long
    ProfilerID As Long
    Name As String * 128
    Vbp As String * 512
    Vbg As String * 512
    IniFile As String * 512
    AppMajor As Long
    AppMinor As Long
    AppRev As Long
    LinesNumber As Long
    ProceduresNumber As Long
    Reserved As String * 1664
End Type

Dim c As CvbwProfile


Public Sub vbwInitializeProfiler()

    Static vbwIsInitialized As Boolean

    If vbwIsInitialized Then
        Exit Sub
    End If
    
    vbwIsInitialized = True
   
   If GetPrivateProfileInt("VB Watch", "CancelOperations", 0, "C:\Program Files\VB Watch 2\VBWatch.ini") Then
        Exit Sub
    End If
    
    
    ' run the profiler
    On Error Resume Next
    Const ProfilerPath = "C:\Program Files\VB Watch 2\VB Watch Profiler.exe"
    ShellExecute 0&, "open", ProfilerPath, "", "", 2
    On Error GoTo 0
    
    ' check that the profiler is running
retry:
    Dim tim As Double
    tim = CDbl(Now)
    Do
        DoEvents
        ' retrieve the profiler handle
        ProfilerHW = GetPrivateProfileInt("Profiler", "hWnd", 0, "C:\Program Files\VB Watch 2\VBWatch.ini")
    Loop Until GetWindowTextLength(ProfilerHW) > 0 Or CDbl(Now) > tim + 0.00010
    
    If GetWindowTextLength(ProfilerHW) = 0 Then
        If App.UnattendedApp = False Then
            If MessageBox(0&, "Can't find profiler at " & ProfilerPath & " !" & vbCrLf & "Please run the VB Watch Profiler manually and press OK...", App.Title, vbOKCancel + vbCritical + vbMsgBoxSetTopMost) = vbOK Then   '
                GoTo retry
            End If
        End If
        Exit Sub
    End If
    
    ' retrieve a unique ID so the profiler can identify us
    ProfilerID = GetPrivateProfileInt("Profiler", "ID", 1, "C:\Program Files\VB Watch 2\VBWatch.ini")
    ' increment ID for other apps
    WritePrivateProfileString "Profiler", "ID", CStr(ProfilerID + 1), "C:\Program Files\VB Watch 2\VBWatch.ini"
    
    ' declare our session to the profiler
    Dim pds As ProfDeclareSession
    pds.ProcessID = GetCurrentProcessId
    pds.ThreadID = App.ThreadID
    pds.ProfilerID = ProfilerID
    pds.Name = "RacingSignals.exe" & Chr(0)
    pds.Vbp = "E:\My Documents\ais\YachtSource\Profiler-RacingSignals 1.1.7\RacingSignals.vbp" & Chr(0)
    pds.Vbg = "" & Chr(0)
    pds.IniFile = "E:\My Documents\ais\YachtSource\vbw2Options_RacingSignals.ini" & Chr(0)
    pds.AppMajor = App.Major
    pds.AppMinor = App.Minor
    pds.AppRev = App.Revision
    pds.LinesNumber = CLng("1522")
    pds.ProceduresNumber = CLng("107")
    SendMessage ProfilerHW, WM_USER_PROFILE_DECLARE, pds.ProcessID, ByVal VarPtr(pds)
        
    ' send fake messages to compute overhead
    Const n As Long = 100
    Dim i As Long
    SendMessage ProfilerHW, WM_USER_PROFILE_START_OVERHEAD, ProfilerID, ByVal n
    For i = 1 To n
        vbwOverhead -1
    Next i
    SendMessage ProfilerHW, WM_USER_PROFILE_END_OVERHEAD, ProfilerID, ByVal n
    
    ' this class will go out of scope when the app terminates
    ' so this will be helpful to track the end of the session
    Set c = New CvbwProfile
    
End Sub

Public Function vbwOverhead(ByRef lLineID As Long) As Boolean
    SendMessage ProfilerHW, WM_USER_PROFILE_OVERHEAD, ProfilerID, ByVal lLineID
End Function

Public Function vbwExecuteLine(ByRef lLineID As Long) As Boolean
    SendMessage ProfilerHW, WM_USER_PROFILE_EXECUTE_LINE, ProfilerID, ByVal lLineID
    ' This function always returns false
End Function

Public Sub vbwProcIn(ByRef lProcID As Long)
    SendMessage ProfilerHW, WM_USER_PROFILE_ENTER_PROCEDURE, ProfilerID, ByVal lProcID
End Sub

Public Sub vbwProcOut(ByRef lProcID As Long)
    SendMessage ProfilerHW, WM_USER_PROFILE_EXIT_PROCEDURE, ProfilerID, ByVal lProcID
End Sub

Public Sub vbwEndSession()
    SendMessage ProfilerHW, WM_USER_PROFILE_END_SESSION, ProfilerID, 0&
End Sub

' Call this anytime in your code when you want to force the session to finish
Sub vbwFinishSession()
    Set c = Nothing
End Sub


