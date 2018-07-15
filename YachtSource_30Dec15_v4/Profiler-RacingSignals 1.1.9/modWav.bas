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
vbwProfiler.vbwProcIn 119
Dim i As Long

vbwProfiler.vbwExecuteLine 1855
      i = waveOutGetNumDevs()
vbwProfiler.vbwExecuteLine 1856
      If i > 0 Then         ' There is at least one device.
vbwProfiler.vbwExecuteLine 1857
            HasSound = True
      End If
vbwProfiler.vbwExecuteLine 1858 'B
vbwProfiler.vbwProcOut 119
vbwProfiler.vbwExecuteLine 1859
End Function

Public Function PlayWav(ByVal FileName As String)
vbwProfiler.vbwProcIn 120

vbwProfiler.vbwExecuteLine 1860
    If HasSound Then
vbwProfiler.vbwExecuteLine 1861
        If WavMem = "" Then
vbwProfiler.vbwExecuteLine 1862
            WavMem = ReadWav(SoundFilePath & FileName)
        End If
vbwProfiler.vbwExecuteLine 1863 'B
vbwProfiler.vbwExecuteLine 1864
        Call PlaySound(WavMem, 0, SND_NODEFAULT Or SND_ASYNC Or SND_MEMORY Or SND_LOOP)

'        Call PlaySound(SoundFilePath & FileName, 0, SND_FILENAME Or SND_NODEFAULT Or SND_ASYNC Or SND_LOOP)
'        Call PlaySound(vbNull, 0, 0)
    End If
vbwProfiler.vbwExecuteLine 1865 'B
vbwProfiler.vbwProcOut 120
vbwProfiler.vbwExecuteLine 1866
End Function
Public Function StopWav()
vbwProfiler.vbwProcIn 121

vbwProfiler.vbwExecuteLine 1867
    If HasSound Then
'        Call PlaySound(FileName, 0, SND_FILENAME Or SND_NODEFAULT Or SND_ASYNC Or SND_LOOP)
vbwProfiler.vbwExecuteLine 1868
        Call PlaySound(vbNull, 0, SND_NODEFAULT Or SND_ASYNC Or SND_MEMORY Or SND_LOOP)
    End If
vbwProfiler.vbwExecuteLine 1869 'B
vbwProfiler.vbwProcOut 121
vbwProfiler.vbwExecuteLine 1870
End Function

Private Function ReadWav(sFile As String) As String
vbwProfiler.vbwProcIn 122
Dim b() As Byte
    Dim nFile       As Integer

vbwProfiler.vbwExecuteLine 1871
    nFile = FreeFile
vbwProfiler.vbwExecuteLine 1872
    Open sFile For Binary Access Read As #nFile
vbwProfiler.vbwExecuteLine 1873
    If LOF(nFile) > 0 Then
vbwProfiler.vbwExecuteLine 1874
        ReDim b(0 To LOF(nFile) - 1)
vbwProfiler.vbwExecuteLine 1875
        Get nFile, , b
    End If
vbwProfiler.vbwExecuteLine 1876 'B
vbwProfiler.vbwExecuteLine 1877
    Close #nFile
vbwProfiler.vbwExecuteLine 1878
    ReadWav = StrConv(b, vbUnicode)
vbwProfiler.vbwProcOut 122
vbwProfiler.vbwExecuteLine 1879
End Function

Private Function ReadFile1(sFile As String) As String
vbwProfiler.vbwProcIn 123
    Dim nFile       As Integer

vbwProfiler.vbwExecuteLine 1880
    nFile = FreeFile
vbwProfiler.vbwExecuteLine 1881
    Open sFile For Input Access Read As #nFile
vbwProfiler.vbwExecuteLine 1882
    If LOF(nFile) > 0 Then
vbwProfiler.vbwExecuteLine 1883
        ReadFile1 = InputB(LOF(nFile), nFile)
    End If
vbwProfiler.vbwExecuteLine 1884 'B
vbwProfiler.vbwExecuteLine 1885
    Close #nFile
vbwProfiler.vbwProcOut 123
vbwProfiler.vbwExecuteLine 1886
End Function


