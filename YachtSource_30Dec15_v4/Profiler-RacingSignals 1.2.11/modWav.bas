Attribute VB_Name = "modWav"
Option Explicit
 'See http://msdn.microsoft.com/en-us/library/ms712587.aspx
 'see http://www.vbforfree.com/mci-multimedia-command-string-tutorial-a-step-by-step-guide/
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long

Private WavFile As String
Private Command As String
Private Success As Boolean
Private retVal As Long
Private returnData As String


Public SoundFilePath As String
Private WavMem As String

Public Function OpenWav(ByVal FileName As String)
vbwProfiler.vbwProcIn 117

vbwProfiler.vbwExecuteLine 2277
    If WavFile <> "" Then
vbwProfiler.vbwExecuteLine 2278
        Call CloseWav
    End If
vbwProfiler.vbwExecuteLine 2279 'B

vbwProfiler.vbwExecuteLine 2280
    WavFile = Chr(34) & SoundFilePath & FileName & Chr(34)
'make the buffer 128 characters
vbwProfiler.vbwExecuteLine 2281
    returnData = String(128, 0) 'Space(128)
vbwProfiler.vbwExecuteLine 2282
    Command = "open " & WavFile & " type waveaudio alias mysound" ' buffer 6"
vbwProfiler.vbwExecuteLine 2283
    retVal = mciSendString(Command, 0, 0, 0)
vbwProfiler.vbwExecuteLine 2284
    Success = mciGetErrorString(retVal, returnData, Len(returnData))
vbwProfiler.vbwExecuteLine 2285
    If Success = False Then
vbwProfiler.vbwExecuteLine 2286
        WavFile = ""
vbwProfiler.vbwExecuteLine 2287
        frmMain.StatusBar1.Panels(3).Picture = Nothing
    Else
vbwProfiler.vbwExecuteLine 2288 'B
vbwProfiler.vbwExecuteLine 2289
        frmMain.StatusBar1.Panels(3).Picture = LoadPicture(SignalImageFilePath & "speaker.gif")
    End If
vbwProfiler.vbwExecuteLine 2290 'B
vbwProfiler.vbwProcOut 117
vbwProfiler.vbwExecuteLine 2291
End Function

Public Function PlayWav()
vbwProfiler.vbwProcIn 118

'May be a controller but no Sound Card
vbwProfiler.vbwExecuteLine 2292
    If WavFile = "" Then
vbwProfiler.vbwProcOut 118
vbwProfiler.vbwExecuteLine 2293
         Exit Function
    End If
vbwProfiler.vbwExecuteLine 2294 'B

vbwProfiler.vbwExecuteLine 2295
    Command = "play mysound from 0"    ' & " from 0 to 500"
vbwProfiler.vbwExecuteLine 2296
    retVal = mciSendString(Command, 0, 0, 0)
'    Debug.Print "play " & retVal
vbwProfiler.vbwExecuteLine 2297
    Success = mciGetErrorString(retVal, returnData, Len(returnData))
vbwProfiler.vbwExecuteLine 2298
    If Success = False Then
vbwProfiler.vbwExecuteLine 2299
         MsgBox Trim(returnData), , Command
    End If
vbwProfiler.vbwExecuteLine 2300 'B
vbwProfiler.vbwProcOut 118
vbwProfiler.vbwExecuteLine 2301
End Function

Public Function PauseWav()
vbwProfiler.vbwProcIn 119

'May be a controller but no Sound Card
vbwProfiler.vbwExecuteLine 2302
    If WavFile = "" Then
vbwProfiler.vbwProcOut 119
vbwProfiler.vbwExecuteLine 2303
         Exit Function
    End If
vbwProfiler.vbwExecuteLine 2304 'B

vbwProfiler.vbwExecuteLine 2305
    retVal = mciSendString("pause mysound", 0, 0, 0)
'    Debug.Print "pause " & retVal
vbwProfiler.vbwExecuteLine 2306
    Success = mciGetErrorString(retVal, returnData, Len(returnData))
vbwProfiler.vbwExecuteLine 2307
    If Success = False Then
vbwProfiler.vbwExecuteLine 2308
         MsgBox Trim(returnData), , Command
    End If
vbwProfiler.vbwExecuteLine 2309 'B
vbwProfiler.vbwProcOut 119
vbwProfiler.vbwExecuteLine 2310
End Function

Public Function CloseWav()
vbwProfiler.vbwProcIn 120

'May be a controller but no Sound Card
vbwProfiler.vbwExecuteLine 2311
    If WavFile = "" Then
vbwProfiler.vbwProcOut 120
vbwProfiler.vbwExecuteLine 2312
         Exit Function
    End If
vbwProfiler.vbwExecuteLine 2313 'B

vbwProfiler.vbwExecuteLine 2314
    Command = "close mysound"
vbwProfiler.vbwExecuteLine 2315
    retVal = mciSendString(Command, 0, 0, 0)
'    Debug.Print "close " & retVal
vbwProfiler.vbwExecuteLine 2316
    Success = mciGetErrorString(retVal, returnData, 128)
vbwProfiler.vbwExecuteLine 2317
    If Success = False Then
vbwProfiler.vbwExecuteLine 2318
         MsgBox Trim(returnData), , Command
    End If
vbwProfiler.vbwExecuteLine 2319 'B
vbwProfiler.vbwProcOut 120
vbwProfiler.vbwExecuteLine 2320
End Function



