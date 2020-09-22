Attribute VB_Name = "ModFiles"
Option Explicit

Public Function OpenFile(tmpFileName) As Integer
Dim FileNr As Integer
Dim x As Integer
Dim tmpData As String
Dim retval As Integer
Dim tmp As String
Dim fileBuffer() As Byte

On Error Resume Next

If tmpFileName <> "" Then GoTo skipDialog

'first save and set current free if owner
Call StoreCurrentNote
If gblnSbookChanged = True Then
    x = SaveFile(gstrCurrentFile)
    If x <> 0 Then
        'warn if not saved
        retval = MsgBox("Unable to save '" & StripFileName(gstrCurrentFile) & "'" & vbCrLf & vbCrLf & "If another file is opened, all changes in the current S-Book will be lost." & vbCrLf & vbCrLf & "Do you want to continue opening another S-Book?", vbYesNo + vbExclamation, " Open S-Book file...")
        If retval = vbNo Then Exit Function
    End If
End If

With frmMain.dlgOpen
    .filename = ""
    .DialogTitle = " Open S-Book File..."
    .Flags = &H1000 Or &H4
    .DefaultExt = ".bsb"
    .InitDir = gstrCurrentFile
    .Filter = "S-Book files (*.bsb)|*.bsb"
    .FilterIndex = 1
    .ShowOpen
    If Err = 32755 Or .filename = "" Then Exit Function
    tmpFileName = .filename
End With
    
skipDialog:

'get key
frmKeyDirect.Caption = " Open S-Book '" & StripFileName(tmpFileName) & "'"
frmKeyDirect.lblCode.Caption = "Enter the Passphrase"
frmKeyDirect.Show (vbModal)
If gblnCancelKey = True Or gstrKeyInput = "" Then
    OpenFile = -1
    Exit Function
    End If

Err.Clear
On Error GoTo ErrHandler

'open the file
Screen.MousePointer = 11
FileNr = FreeFile
Open tmpFileName For Binary Access Read As #FileNr
frmMain.txtBox.Text = vbCrLf & "   Opening File..."
frmMain.txtBox.Refresh
ReDim fileBuffer(LOF(FileNr) - 3) '(disregard crlf)
Get #FileNr, , fileBuffer
Close #FileNr
tmpData = StrConv(fileBuffer(), vbUnicode)
ReDim fileBuffer(0)

'decrypt data
frmMain.txtBox.Text = vbCrLf & "   Decrypting file, please wait..."
frmMain.txtBox.Refresh
tmpData = DecodeString(tmpData, gstrKeyInput)
If ErrorFlag = False Then
    gstrCurrentKey = gstrKeyInput
    Else
    MsgBox "Faild Decrypting the S-Book File:" & vbCrLf & vbCrLf & ErrorDescription, vbCritical
    Screen.MousePointer = 0
    frmMain.txtBox.Text = ""
    OpenFile = -1
    Call ShowCurrentNote
    Call CheckButtons
    Exit Function
End If

Screen.MousePointer = 0

'check attribute read only
If GetAttr(tmpFileName) And 1 Then
    gblnReadOnly = True
    Else
    gblnReadOnly = False
    End If

gstrSbook = tmpData
frmMain.Caption = " S-Book - " & StripFileName(tmpFileName)
frmMain.txtBox.Locked = True
If gblnReadOnly = True Then frmMain.Caption = frmMain.Caption & " (Read Only)"
gstrCurrentFile = tmpFileName

'show last note
curPos = Len(gstrSbook)
gblnSbookChanged = False
Call GetNote(Down)
Call ShowCurrentNote
Call CheckButtons
Exit Function

ErrHandler:
    MsgBox "Unable to open '" & StripFileName(tmpFileName) & "'" & vbCrLf & vbCrLf & "Error: " & Error, vbCritical
    Close #FileNr
    Err.Clear
    OpenFile = -1
    Screen.MousePointer = 0
    frmMain.txtBox.Text = ""
    Call ShowCurrentNote
    Call CheckButtons
End Function

Public Function SaveFile(tmpFileName As String) As Integer
Dim FileNr
Dim tmpData As String

'check for key
If gstrKeyInput = "" Then
    frmKey.Caption = " Save S-book '" & StripFileName(gstrCurrentFile) & "'"
    frmKey.Show (vbModal)
    gstrCurrentKey = gstrKeyInput
    If gblnCancelKey = True Or gstrCurrentKey = "" Then
        SaveFile = -1
        Exit Function
    End If
End If

'prepare save
Call StoreCurrentNote
If gblnReadOnly = True Then SaveFile = -1: Exit Function
If gblnSbookChanged = False Then Exit Function

frmMain.txtBox.Text = vbCrLf & "   Encrypting file, please wait..."
frmMain.txtBox.Refresh

'encrypt data
Screen.MousePointer = 11
tmpData = EncodeString(gstrSbook, gstrCurrentKey)
If ErrorFlag = True Then
    frmMain.txtBox.Text = vbCrLf & ""
    MsgBox "Failed encrypting the S-Book File:" & vbCrLf & vbCrLf & ErrorDescription, vbCritical
    Screen.MousePointer = 0
    SaveFile = -1
    Call ShowCurrentNote
    Call CheckButtons
    Exit Function
    End If
    
'save note
frmMain.txtBox.Text = vbCrLf & "   Saving File..."
frmMain.txtBox.Refresh

On Error GoTo errorHandle
FileNr = FreeFile
Open tmpFileName For Output As #FileNr
Print #FileNr, tmpData
Close #FileNr

'set parm after good save
Screen.MousePointer = 0
gblnSbookChanged = False
gstrOldNote = ""
frmMain.txtBox.Text = ""
Call ShowCurrentNote
Call CheckButtons
Exit Function

errorHandle:
    Close #FileNr
    MsgBox "Failed saving '" & StripFileName(tmpFileName) & "'" & vbCrLf & vbCrLf & "Error: " & Error, vbCritical
    frmMain.txtBox.Text = ""
    Call ShowCurrentNote
    Screen.MousePointer = 0
    SaveFile = -1
    Call ShowCurrentNote
    Call CheckButtons
End Function

Public Sub CreateFile()
Dim FileNr
Dim x As Integer
Dim tmpName As String
Dim retval As Integer
Dim oldFileName As String
On Error Resume Next

'first save and set current free if owner
Call StoreCurrentNote
If gblnSbookChanged = True Then
    x = SaveFile(gstrCurrentFile)
    If x <> 0 Then
        'warn if not saved
        retval = MsgBox("Unable to save '" & StripFileName(gstrCurrentFile) & "'" & vbCrLf & vbCrLf & "If another file is opened, all changes in the current S-Book will be lost." & vbCrLf & vbCrLf & "Do you want to continue opening another S-Book?", vbYesNo + vbExclamation, " Open S-Book file...")
        If retval = vbNo Then Exit Sub
    End If
End If

With frmMain.dlgOpen
    .filename = ""
    .DialogTitle = " Create new S-Book file..."
    .Flags = &H2 Or &H4
    .DefaultExt = ".bsb"
    .InitDir = gstrCurrentFile
    .Filter = "S-Book files (*.bsb)|*.bsb"
    .FilterIndex = 1
    .ShowSave
    If Err = 32755 Or .filename = "" Then Exit Sub
    tmpName = .filename
    oldFileName = gstrCurrentFile
End With

'set new s-book
gstrSbook = ""
gblnSbookChanged = True
gstrCurrentFile = tmpName
frmMain.Caption = " S-Book - " & StripFileName(gstrCurrentFile)
frmMain.txtBox.Text = ""
gstrOldNote = ""
curStop = 0
curStart = 0
curPos = 0
Call FindNote(Down)
Call ShowCurrentNote
Call CheckButtons
gstrKeyInput = ""
End Sub
