Attribute VB_Name = "modWav"
Option Explicit

Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" _
    (ByVal lpszName As String, _
     ByVal hModule As Long, _
     ByVal dwFlags As Long) As Long

Private Const SND_APPLICATION As Long = &H80
Private Const SND_ALIAS As Long = &H10000
Private Const SND_ALIAS_ID As Long = &H110000
Private Const SND_ASYNC As Long = &H1
Private Const SND_FILENAME As Long = &H20000
Private Const SND_LOOP As Long = &H8
Private Const SND_MEMORY As Long = &H4
Private Const SND_NODEFAULT As Long = &H2
Private Const SND_NOSTOP As Long = &H10
Private Const SND_NOWAIT As Long = &H2000
Private Const SND_PURGE As Long = &H40
Private Const SND_RESOURCE As Long = &H40004
Private Const SND_SYNC As Long = &H0

Private Declare Function waveOutGetNumDevs Lib "winmm" () As Long

Public SoundFilePath As String
Private WavMem As String

Public Function HasSound() As Boolean
vbwProfiler.vbwProcIn 103
Dim i As Long

vbwProfiler.vbwExecuteLine 1491
      i = waveOutGetNumDevs()
vbwProfiler.vbwExecuteLine 1492
      If i > 0 Then         ' There is at least one device.
vbwProfiler.vbwExecuteLine 1493
            HasSound = True
      End If
vbwProfiler.vbwExecuteLine 1494 'B
vbwProfiler.vbwProcOut 103
vbwProfiler.vbwExecuteLine 1495
End Function

Public Function PlayWav(ByVal FileName As String)
vbwProfiler.vbwProcIn 104

vbwProfiler.vbwExecuteLine 1496
    If HasSound Then
vbwProfiler.vbwExecuteLine 1497
        If WavMem = "" Then
vbwProfiler.vbwExecuteLine 1498
            WavMem = ReadWav(SoundFilePath & FileName)
        End If
vbwProfiler.vbwExecuteLine 1499 'B
vbwProfiler.vbwExecuteLine 1500
        Call PlaySound(WavMem, 0, SND_NODEFAULT Or SND_ASYNC Or SND_MEMORY Or SND_LOOP)

'        Call PlaySound(SoundFilePath & FileName, 0, SND_FILENAME Or SND_NODEFAULT Or SND_ASYNC Or SND_LOOP)
'        Call PlaySound(vbNull, 0, 0)
    End If
vbwProfiler.vbwExecuteLine 1501 'B
vbwProfiler.vbwProcOut 104
vbwProfiler.vbwExecuteLine 1502
End Function
Public Function StopWav()
vbwProfiler.vbwProcIn 105

vbwProfiler.vbwExecuteLine 1503
    If HasSound Then
'        Call PlaySound(FileName, 0, SND_FILENAME Or SND_NODEFAULT Or SND_ASYNC Or SND_LOOP)
vbwProfiler.vbwExecuteLine 1504
        Call PlaySound(vbNull, 0, SND_NODEFAULT Or SND_ASYNC Or SND_MEMORY Or SND_LOOP)
    End If
vbwProfiler.vbwExecuteLine 1505 'B
vbwProfiler.vbwProcOut 105
vbwProfiler.vbwExecuteLine 1506
End Function

Private Function ReadWav(sFile As String) As String
vbwProfiler.vbwProcIn 106
Dim b() As Byte
    Dim nFile       As Integer

vbwProfiler.vbwExecuteLine 1507
    nFile = FreeFile
vbwProfiler.vbwExecuteLine 1508
    Open sFile For Binary Access Read As #nFile
vbwProfiler.vbwExecuteLine 1509
    If LOF(nFile) > 0 Then
vbwProfiler.vbwExecuteLine 1510
        ReDim b(0 To LOF(nFile) - 1)
vbwProfiler.vbwExecuteLine 1511
        Get nFile, , b
    End If
vbwProfiler.vbwExecuteLine 1512 'B
vbwProfiler.vbwExecuteLine 1513
    Close #nFile
vbwProfiler.vbwExecuteLine 1514
    ReadWav = StrConv(b, vbUnicode)
vbwProfiler.vbwProcOut 106
vbwProfiler.vbwExecuteLine 1515
End Function

Private Function ReadFile1(sFile As String) As String
vbwProfiler.vbwProcIn 107
    Dim nFile       As Integer

vbwProfiler.vbwExecuteLine 1516
    nFile = FreeFile
vbwProfiler.vbwExecuteLine 1517
    Open sFile For Input Access Read As #nFile
vbwProfiler.vbwExecuteLine 1518
    If LOF(nFile) > 0 Then
vbwProfiler.vbwExecuteLine 1519
        ReadFile1 = InputB(LOF(nFile), nFile)
    End If
vbwProfiler.vbwExecuteLine 1520 'B
vbwProfiler.vbwExecuteLine 1521
    Close #nFile
vbwProfiler.vbwProcOut 107
vbwProfiler.vbwExecuteLine 1522
End Function


