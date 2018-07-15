Attribute VB_Name = "modEncrypt"
Option Explicit

'Requires CAPICOM V2.1 Project > Reference including
Public Function Decrypt(kb As String) As String
vbwProfiler.vbwProcIn 121
    Dim Secret As EncryptedData

vbwProfiler.vbwExecuteLine 2321
    Set Secret = New EncryptedData
vbwProfiler.vbwExecuteLine 2322
    Secret.Algorithm.Name = CAPICOM_ENCRYPTION_ALGORITHM_AES
vbwProfiler.vbwExecuteLine 2323
    Secret.Algorithm.KeyLength = CAPICOM_ENCRYPTION_KEY_LENGTH_256_BITS
vbwProfiler.vbwExecuteLine 2324
    Secret.SetSecret "My Secret Encryption Key"
'    Secret.Content = "password" ' just so we know that this is being reset by decryption
vbwProfiler.vbwExecuteLine 2325
    On Error Resume Next    'errror if .secret differs
vbwProfiler.vbwExecuteLine 2326
    Secret.Decrypt kb
vbwProfiler.vbwExecuteLine 2327
    On Error GoTo 0
vbwProfiler.vbwExecuteLine 2328
    Decrypt = Secret.Content
'MsgBox Decrypt
vbwProfiler.vbwProcOut 121
vbwProfiler.vbwExecuteLine 2329
End Function

Public Function Encrypt(kb As String) As String
vbwProfiler.vbwProcIn 122
Dim Secret As EncryptedData
vbwProfiler.vbwExecuteLine 2330
    Set Secret = New EncryptedData
vbwProfiler.vbwExecuteLine 2331
    Secret.Algorithm.Name = CAPICOM_ENCRYPTION_ALGORITHM_AES
vbwProfiler.vbwExecuteLine 2332
    Secret.Algorithm.KeyLength = CAPICOM_ENCRYPTION_KEY_LENGTH_256_BITS
vbwProfiler.vbwExecuteLine 2333
    Secret.SetSecret "My Secret Encryption Key"
vbwProfiler.vbwExecuteLine 2334
    Secret.Content = kb ' what we want to encrypt
vbwProfiler.vbwExecuteLine 2335
    Encrypt = Secret.Encrypt
'For Password encryption (AisDecoder)
'we must remove the split lines secret.content includes
'    Encrypt = Replace(Secret.Encrypt, vbCrLf, "")
'    MsgBox Encrypt
vbwProfiler.vbwProcOut 122
vbwProfiler.vbwExecuteLine 2336
End Function

Public Function DecryptFile(EncryptedFileName As String, DecryptedFileName As String)
vbwProfiler.vbwProcIn 123
Dim EncryptedLines As String
Dim DecryptedLines As String
Dim Ch As Long

'MsgBox "Decrypting " & EncryptedFileName & vbCrLf & "to " & DecryptedFileName
vbwProfiler.vbwExecuteLine 2337
    Ch = FreeFile
vbwProfiler.vbwExecuteLine 2338
    Open EncryptedFileName For Input As #Ch
vbwProfiler.vbwExecuteLine 2339
    EncryptedLines = StrConv(InputB(LOF(Ch), Ch), vbUnicode)
vbwProfiler.vbwExecuteLine 2340
    Close Ch
vbwProfiler.vbwExecuteLine 2341
    DecryptedLines = Decrypt(EncryptedLines)
vbwProfiler.vbwExecuteLine 2342
    Open DecryptedFileName For Output As #Ch Len = Len(DecryptedLines)
vbwProfiler.vbwExecuteLine 2343
    Print #Ch, DecryptedLines
vbwProfiler.vbwExecuteLine 2344
    Close #Ch
vbwProfiler.vbwProcOut 123
vbwProfiler.vbwExecuteLine 2345
End Function

Public Function EncryptFile(DecryptedFileName As String, EncryptedFileName As String)
vbwProfiler.vbwProcIn 124
Dim EncryptedLines As String
Dim DecryptedLines As String
Dim Ch As Long
Dim l As Integer

'MsgBox "Encrypting " & DecryptedFileName & vbCrLf & "to " & EncryptedFileName
vbwProfiler.vbwExecuteLine 2346
    Ch = FreeFile
vbwProfiler.vbwExecuteLine 2347
    Open DecryptedFileName For Input As #Ch
vbwProfiler.vbwExecuteLine 2348
    DecryptedLines = StrConv(InputB(LOF(Ch), Ch), vbUnicode)
vbwProfiler.vbwExecuteLine 2349
    Close Ch
vbwProfiler.vbwExecuteLine 2350
    EncryptedLines = Encrypt(DecryptedLines)
vbwProfiler.vbwExecuteLine 2351
    Open EncryptedFileName For Output As #Ch '    Len = Len(EncryptedLines)
vbwProfiler.vbwExecuteLine 2352
    Print #Ch, EncryptedLines
vbwProfiler.vbwExecuteLine 2353
    Close #Ch
vbwProfiler.vbwProcOut 124
vbwProfiler.vbwExecuteLine 2354
End Function

Public Function EncryptFiles(FilePath As String, DecryptedExt As String, EncryptedExt As String)
vbwProfiler.vbwProcIn 125
Dim DecryptedFileName As String
Dim EncryptedFileName As String

'MsgBox FilePath
vbwProfiler.vbwExecuteLine 2355
    DecryptedFileName = Dir$(FilePath & "*" & DecryptedExt)
'vbwLine 2356:    Do While DecryptedFileName > ""
    Do While vbwProfiler.vbwExecuteLine(2356) Or DecryptedFileName > ""
vbwProfiler.vbwExecuteLine 2357
        If Right$(DecryptedFileName, Len(DecryptedExt)) = DecryptedExt Then
vbwProfiler.vbwExecuteLine 2358
            EncryptedFileName = Replace(DecryptedFileName, DecryptedExt, EncryptedExt)
vbwProfiler.vbwExecuteLine 2359
            Call EncryptFile(FilePath & DecryptedFileName, FilePath & EncryptedFileName)
        End If
vbwProfiler.vbwExecuteLine 2360 'B
vbwProfiler.vbwExecuteLine 2361
        DecryptedFileName = Dir$
vbwProfiler.vbwExecuteLine 2362
    Loop
vbwProfiler.vbwProcOut 125
vbwProfiler.vbwExecuteLine 2363
End Function

