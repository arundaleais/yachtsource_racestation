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
vbwProfiler.vbwProcIn 88
Dim i As Long

vbwProfiler.vbwExecuteLine 1629
      i = waveOutGetNumDevs()
vbwProfiler.vbwExecuteLine 1630
      If i > 0 Then         ' There is at least one device.
vbwProfiler.vbwExecuteLine 1631
            HasSound = True
      End If
vbwProfiler.vbwExecuteLine 1632 'B
vbwProfiler.vbwProcOut 88
vbwProfiler.vbwExecuteLine 1633
End Function

Public Function PlayWav(ByVal FileName As String)
vbwProfiler.vbwProcIn 89

vbwProfiler.vbwExecuteLine 1634
    If HasSound Then
vbwProfiler.vbwExecuteLine 1635
        If WavMem = "" Then
vbwProfiler.vbwExecuteLine 1636
            WavMem = ReadWav(SoundFilePath & FileName)
        End If
vbwProfiler.vbwExecuteLine 1637 'B
vbwProfiler.vbwExecuteLine 1638
        Call PlaySound(WavMem, 0, SND_NODEFAULT Or SND_ASYNC Or SND_MEMORY Or SND_LOOP)

'        Call PlaySound(SoundFilePath & FileName, 0, SND_FILENAME Or SND_NODEFAULT Or SND_ASYNC Or SND_LOOP)
'        Call PlaySound(vbNull, 0, 0)
    End If
vbwProfiler.vbwExecuteLine 1639 'B
vbwProfiler.vbwProcOut 89
vbwProfiler.vbwExecuteLine 1640
End Function
Public Function StopWav()
vbwProfiler.vbwProcIn 90

vbwProfiler.vbwExecuteLine 1641
    If HasSound Then
'        Call PlaySound(FileName, 0, SND_FILENAME Or SND_NODEFAULT Or SND_ASYNC Or SND_LOOP)
vbwProfiler.vbwExecuteLine 1642
        Call PlaySound(vbNull, 0, SND_NODEFAULT Or SND_ASYNC Or SND_MEMORY Or SND_LOOP)
    End If
vbwProfiler.vbwExecuteLine 1643 'B
vbwProfiler.vbwProcOut 90
vbwProfiler.vbwExecuteLine 1644
End Function

Private Function ReadWav(sFile As String) As String
vbwProfiler.vbwProcIn 91
Dim b() As Byte
    Dim nFile       As Integer

vbwProfiler.vbwExecuteLine 1645
    nFile = FreeFile
vbwProfiler.vbwExecuteLine 1646
    Open sFile For Binary Access Read As #nFile
vbwProfiler.vbwExecuteLine 1647
    If LOF(nFile) > 0 Then
vbwProfiler.vbwExecuteLine 1648
        ReDim b(0 To LOF(nFile) - 1)
vbwProfiler.vbwExecuteLine 1649
        Get nFile, , b
    End If
vbwProfiler.vbwExecuteLine 1650 'B
vbwProfiler.vbwExecuteLine 1651
    Close #nFile
vbwProfiler.vbwExecuteLine 1652
    ReadWav = StrConv(b, vbUnicode)
vbwProfiler.vbwProcOut 91
vbwProfiler.vbwExecuteLine 1653
End Function

Private Function ReadFile1(sFile As String) As String
vbwProfiler.vbwProcIn 92
    Dim nFile       As Integer

vbwProfiler.vbwExecuteLine 1654
    nFile = FreeFile
vbwProfiler.vbwExecuteLine 1655
    Open sFile For Input Access Read As #nFile
vbwProfiler.vbwExecuteLine 1656
    If LOF(nFile) > 0 Then
vbwProfiler.vbwExecuteLine 1657
        ReadFile1 = InputB(LOF(nFile), nFile)
    End If
vbwProfiler.vbwExecuteLine 1658 'B
vbwProfiler.vbwExecuteLine 1659
    Close #nFile
vbwProfiler.vbwProcOut 92
vbwProfiler.vbwExecuteLine 1660
End Function


