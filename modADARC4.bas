Attribute VB_Name = "modADARC4"

'------------------------------------------------------------
'
'              ADARCFOUR v2.0 TEXT ENCRYPTION CLASS MODULE
'
'------------------------------------------------------------
'
' retStr = EncodeString(myText, myKey)
'
' retStr = DecodeString(myText, myKey)
'
' Where:
'
' - myText [string] = string to encrypt/decrypt
' - myKey  [string] = key or key phrase
'
' Version Info
'
' v1.0 had no Key IV - No longer in use
' v2.0 has Key IV.
' v2.0 is downwards compatible: v2.0 can decrypt v1.0 data
' and will automatically save in v2.0 on next encryption!
'
'------------------------------------------------------------
Option Explicit


Private sBox(255)   As Integer
Private P1          As Integer
Private P2          As Integer

'err handling
Public ErrorFlag As Boolean
Public ErrorDescription As String



' ------------------------------------------------------------
'                   Encryption algorithm functions
' ------------------------------------------------------------

Private Sub SetKey(ByVal aKey As String)
'initialize the PRNG key
Dim i       As Long
Dim j       As Long
Dim k       As Integer

Dim KeyLen  As Long
Dim key()   As Byte
Dim tmp     As Integer

'convert key to array
key() = StrConv(aKey, vbFromUnicode)
KeyLen = Len(aKey)

'fill sBox array
For i = 0 To 255
    sBox(i) = i
Next i

'transpose sBox-shedule 24 times with key
For k = 1 To 24
    For i = 0 To 255
        j = (j + sBox(i) + key(i Mod KeyLen)) Mod 256
        tmp = sBox(i)
        sBox(i) = sBox(j)
        sBox(j) = tmp
    Next
Next

P1 = 0
P2 = 0

End Sub

Public Function EncodeString(ByVal aText As String, ByVal KeyString As String) As String
Dim OFB As Byte
Dim k As Double
Dim dataIV As String
Dim aByte As Integer
Dim aCode As Integer
Dim KeyIV As String

ErrorFlag = False
ErrorDescription = ""

aText = TextTrim(aText)

If aText = "" Then
    ErrorDescription = "No data to encrypt"
    ErrorFlag = True
    Screen.MousePointer = 0
    Exit Function
    End If

If IsGoodKey(KeyString) = False Then
    ErrorDescription = "Key too small or containing repetitions"
    ErrorFlag = True
    Screen.MousePointer = 0
    Exit Function
    End If

Screen.MousePointer = 11

'add header 64 random bytes and trailer checkbytes to data
Randomize
For k = 1 To 64
    dataIV = dataIV & Chr(GetRND)
Next
aText = dataIV & aText & Chr(255) & Chr(255)

'create Init Vector
KeyIV = ""
For k = 1 To 16
    KeyIV = KeyIV & Chr(GetRND)
Next
'add Init Vector to key
KeyString = KeyString & KeyIV
'set key
Call SetKey(KeyString)

'encode cycle
OFB = 0
For k = 1 To Len(aText)
    aByte = Asc(Mid(aText, k, 1))
    aCode = GetCSPRNG(OFB)
    OFB = aByte
    aByte = aByte Xor aCode
    Mid(aText, k, 1) = Chr(aByte)
Next

'add version info (v2.0) and IV to encrypted string
EncodeString = "V20" & KeyIV & aText
Screen.MousePointer = 0

End Function

Public Function DecodeString(ByVal aText As String, ByVal KeyString As String) As String
Dim OFB As Byte
Dim k As Double
Dim aByte As Integer
Dim aCode As Integer
Dim KeyIV As String

ErrorFlag = False
ErrorDescription = ""

If Left(aText, 8) = "[NOTE]" & vbCrLf Then
    'non encrypted data
    DecodeString = aText
    Exit Function
    End If

If Len(aText) < 81 Then
    ErrorDescription = "No data to encrypt"
    ErrorFlag = True
    Screen.MousePointer = 0
    Exit Function
    End If

If IsGoodKey(KeyString) = False Then
    ErrorDescription = "Key too small or containing repetitions"
    ErrorFlag = True
    Screen.MousePointer = 0
    Exit Function
    End If

Screen.MousePointer = 11

'check version (<v2.0 has no version info!)
If Left(aText, 3) = "V20" Then
    'retrieve and cut off IV
    KeyIV = Mid(aText, 4, 16)
    aText = Mid(aText, 20)
    Else
    'v1.0 (without key IV)
    KeyIV = ""
    MsgBox "This file was encrypted with an earlier version (ADARCFOUR v1.0)." & vbCrLf & "Any changes to the file will be encrypted and saved with v2.0", vbInformation + vbOKOnly
End If
'add Init Vector to key
KeyString = KeyString & KeyIV

'set data CSPRNG key
Call SetKey(KeyString)

OFB = 0
For k = 1 To Len(aText)
    aByte = Asc(Mid(aText, k, 1)) Xor GetCSPRNG(OFB)
    OFB = aByte
    Mid(aText, k, 1) = Chr(aByte)
Next

If Right(aText, 2) <> Chr(255) & Chr(255) Then
    ErrorDescription = "Checksum Error (wrong key or corrupted data)"
    ErrorFlag = True
    Exit Function
    End If

'trim check bytes
aText = Left(aText, Len(aText) - 2)

DecodeString = Mid(aText, 65)
Screen.MousePointer = 0

End Function

Private Function GetCSPRNG(feedBack As Byte) As Byte
'generate next byte (with feedback from data)
Dim tmp As Integer
P1 = (P1 + 1) Mod 256
P2 = (P2 + sBox(P1) + feedBack) Mod 256
tmp = sBox(P1)
sBox(P1) = sBox(P2)
sBox(P2) = tmp
GetCSPRNG = sBox((sBox(P1) + sBox(P2)) Mod 256)
End Function

Private Function GetRND() As Integer
'get a non-secure rnd byte
GetRND = Int((256 * Rnd) + 0)
End Function

Public Function IsGoodKey(ByVal aKey As String) As Boolean
'check if key is at least 5 char long, and doesn't repeat
Dim tmp As String
Dim Wid As Integer
Dim i As Integer
Dim Repro As Boolean
If Len(aKey) < 5 Then Exit Function
For Wid = 1 To Int(Len(aKey) / 2)
    IsGoodKey = False
    For i = Wid + 1 To Len(aKey) Step Wid
        If Mid(aKey, 1, Wid) <> Mid(aKey, i, Wid) Then IsGoodKey = True: Exit For
    Next
If IsGoodKey = False Then Exit For
Next
End Function

Public Function TextTrim(ByVal aText As String) As String
'cut off all heading and trailing spaces,tabs,CR's and LF's
Dim tmp As String
BeginCutL:
tmp = Left(aText, 1)
If tmp = Chr(32) Or tmp = Chr(9) Or tmp = Chr(13) Or tmp = Chr(10) Then
    aText = Mid(aText, 2)
    GoTo BeginCutL
    End If
BeginCutR:
tmp = Right(aText, 1)
If tmp = Chr(32) Or tmp = Chr(9) Or tmp = Chr(13) Or tmp = Chr(10) Then
    aText = Left(aText, Len(aText) - 1)
    GoTo BeginCutR
    End If
TextTrim = aText
End Function






