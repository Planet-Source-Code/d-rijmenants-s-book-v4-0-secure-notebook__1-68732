Attribute VB_Name = "ModMain"

'-------------------------------------
'                                    '
'            S-Book v4.0             '
'                                    '
'      (c) D. Rijmenants 2005        '
'                                    '
'-------------------------------------

Option Explicit

Public gstrSbook        As String
Public gstrCurrentFile  As String
Public gstrOldNote      As String
Public gstrPrevSearch   As String

Public gstrKeyInput     As String
Public gstrCurrentKey   As String
Public gblnCancelKey    As Boolean


Public gblnSbookChanged As Boolean
Public gblnReadOnly     As Boolean
Public gblnSearchBusy   As Boolean
Public gblnEditNewNote  As Boolean
Public gblnNewSearch    As Boolean
Public gblnSearchLock   As Boolean
Public gblnSearchHit    As Boolean
Public gcontinueFind    As Boolean

Public gblnChangeMarges As Boolean
Public gblnPrintBusy    As Boolean
Public gblnCancelPrint  As Boolean

Public gintArguments    As Integer
Public gstrArg1         As String
Public gstrArg2         As String
Public gstrArg3         As String

Public curPos           As Long
Public curStart         As Long
Public curStop          As Long
Public oldStart         As Long
Public oldStop          As Long
Public lastFound        As Long

Public glngLmPrint      As Long
Public glngRmPrint      As Long
Public glngTmPrint      As Long
Public glngBmPrint      As Long

Public Enum Direction
    Up
    Down
End Enum

Public Const COL_WHITE = &H8000000E
Public Const COL_GRAY = &H8000000B

Sub Main()
Dim tmp     As String
Dim x       As Integer
Load frmMain
Load frmPage
Load frmAbout
Load frmKeyDirect
Load frmKey
Load frmHelp

'settings menu
On Error Resume Next

'printer check
tmp = Printer.DeviceName
If Err Or tmp = "" Then
    frmMain.Toolbar1.Buttons("print").Enabled = False
    frmMain.mniPrint.Enabled = False
    Err.Clear
    Else
    frmMain.Toolbar1.Buttons("print").Enabled = True
    frmMain.mniPrint.Enabled = True
    End If

' marges printer
glngTmPrint = Val(GetSetting(App.EXEName, "CONFIG", "PrinterTop", "5"))
glngBmPrint = Val(GetSetting(App.EXEName, "CONFIG", "PrinterBottom", "5"))
glngLmPrint = Val(GetSetting(App.EXEName, "CONFIG", "PrinterLeft", "5"))
glngRmPrint = Val(GetSetting(App.EXEName, "CONFIG", "PrinterRight", "5"))
frmPage.txtPrintHead = GetSetting(App.EXEName, "CONFIG", "Headtext", "")
frmPage.chkPrint.Value = Val(GetSetting(App.EXEName, "CONFIG", "Chkhead", "false"))
Call GetWindowPos
frmMain.Show

'snb filetype
tmp = ReadKey(HKEY_CLASSES_ROOT, ".bsb", "", "")
If tmp = "" Then
    frmMain.mniFileType.Caption = "&Register .bsb filetype"
    Else
    frmMain.mniFileType.Caption = "&Unregister .bsb filetype"
    End If

' current note file
tmp = GetSetting(App.EXEName, "CONFIG", "File", "")
If Command <> "" Then
    If Right(Command, 4) = ".bsb" Then
        tmp = Command
        End If
    End If

'check file
If tmp = "" Or Dir(tmp) = "" Then
    tmp = ""
    'open-file window if no file
    'x = OpenFile("")
    Else
    x = OpenFile(tmp)
    End If
If x = 0 Then
    'show last note
    curPos = Len(gstrSbook)
    Call FindNote(Down)
    Call GetWindowPos
    Else
    gstrCurrentFile = ""
    frmMain.txtBox.Text = vbCrLf & " Please open or create S-Book file."
    End If
Call CheckButtons
frmMain.Show
End Sub

Public Sub GetNote(UpDn As Direction)
Dim sta1 As Long
Dim stp1 As Long
Dim revPos As Long
Dim allArg As String
If UpDn = Direction.Up Then
    'search up
    sta1 = InStr(curPos + 1, gstrSbook, "[NOTE]" & vbCrLf)
    If sta1 > 0 Then stp1 = InStr(sta1 + 1, gstrSbook, "[NOTE]" & vbCrLf)
    Else
    'search down
    For revPos = curPos - 1 To 1 Step -1
        If revPos < 1 Then revPos = 0: Exit For
        If Mid(gstrSbook, revPos, 8) = "[NOTE]" & vbCrLf Then Exit For
        Next revPos
    sta1 = revPos
    If sta1 < 0 Then sta1 = 0
    stp1 = InStr(sta1 + 1, gstrSbook, "[NOTE]" & vbCrLf)
End If
If sta1 > 0 Then
    curStart = sta1
    curStop = stp1
    curPos = sta1
    Else
    curPos = 0
    End If
Call SetPosBar(curPos)
End Sub

Public Sub FindNote(UpDn As Direction)
Dim k           As Integer
Dim tmp         As String
Dim blnA        As Boolean
Dim allArg      As String
Dim blnFound    As Boolean
oldStart = curStart
oldStop = curStop
'continue same search?
If gstrPrevSearch = frmMain.Combo1.Text Then
    gblnNewSearch = False
    Else
    gblnNewSearch = True
    gstrPrevSearch = frmMain.Combo1.Text
    End If
' check for limit begin or end
If UpDn = Direction.Down And curStop > Len(gstrSbook) + 1 Then Beep: Exit Sub
If UpDn = Direction.Up And curStart < 1 Then Beep: Exit Sub
If Not gblnSearchBusy Then
    Call GetNote(UpDn) '<<<<<<
    ShowCurrentNote
    Exit Sub
    End If
If gblnNewSearch Then
    'add new search to combo list
    For k = 0 To frmMain.Combo1.ListCount
    If frmMain.Combo1.Text = frmMain.Combo1.List(k) Then blnA = True
    Next k
    If Not blnA Then frmMain.Combo1.AddItem (frmMain.Combo1.Text)
    End If
'if new search, start at begin or end
If gblnNewSearch = True And UpDn = Direction.Up Then curPos = 0: gblnNewSearch = False
If gblnNewSearch = True And UpDn = Direction.Down Then curPos = Len(gstrSbook): gblnNewSearch = False
Call SetArguments
Screen.MousePointer = 11
frmMain.Combo1.SetFocus
gcontinueFind = True
'block all menu's during search
With frmMain
.mnuFile.Enabled = False
.mnuEdit.Enabled = False
.mnuInfo.Enabled = False
.Toolbar1.Enabled = False
.Toolbar2.Enabled = False
.txtBox.Enabled = False
End With
Do
'get next note
Call GetNote(UpDn)
If curPos = 0 Then gcontinueFind = False: Exit Do
'check if match
If curStop <> 0 Then
    tmp = UCase(Mid(gstrSbook, curStart, (curStop - curStart)))
    Else
    tmp = UCase(Mid(gstrSbook, curStart))
    End If
blnFound = FindMatch(Mid(tmp, 9))
'make interrupt possible when searching next hit
DoEvents
If gcontinueFind = False Then blnFound = True
Loop While Not blnFound
With frmMain
.mnuFile.Enabled = True
.mnuEdit.Enabled = True
.mnuInfo.Enabled = True
.Toolbar1.Enabled = True
.Toolbar2.Enabled = True
.txtBox.Enabled = True
End With
Screen.MousePointer = 0
If gintArguments = 0 Then blnFound = True
If blnFound Then
    'show hit
    Call ShowCurrentNote
    gblnSearchHit = True
    Else
    'not found,set argument for search next hit
    If gintArguments = 3 Then allArg = gstrArg1 & " + " & gstrArg2 & " + " & gstrArg3
    If gintArguments = 2 Then allArg = gstrArg1 & " + " & gstrArg2
    If gintArguments = 1 Then allArg = gstrArg1
    If gblnSearchHit = True Then
        'if previous was hit
        MsgBox "Search for '" & allArg & "'  completed.", vbInformation, " Find"
        Else
        'if no previours hit
        MsgBox "'" & allArg & "' not found.", vbInformation, " Find"
        End If
    'set pointers to last hit
    gblnSearchHit = False
    gstrPrevSearch = ""
    curPos = lastFound
    SetPosBar curPos
    curStart = oldStart
    curStop = oldStop
    frmMain.Combo1.SetFocus
    End If
gcontinueFind = False
End Sub

Private Function FindMatch(strA As String) As Boolean
Dim ar1 As String
Dim ar2 As String
Dim ar3 As String
Dim blnFound As Boolean
ar1 = UCase(gstrArg1)
ar2 = UCase(gstrArg2)
ar3 = UCase(gstrArg3)
strA = UCase(strA)
Select Case gintArguments
Case 0
    blnFound = True
Case 1
    If InStr(1, strA, ar1) <> 0 Then FindMatch = True
Case 2
    If InStr(1, strA, ar1) <> 0 _
    And InStr(1, strA, ar2) <> 0 Then FindMatch = True
Case 3
    If InStr(1, strA, ar1) <> 0 _
    And InStr(1, strA, ar2) <> 0 _
    And InStr(1, strA, ar3) <> 0 Then FindMatch = True
End Select
End Function

Sub ShowCurrentNote()
Dim tmp
Dim Header As String
If gstrCurrentFile = "" Then
    frmMain.txtBox.Text = vbCrLf & " Please open or create S-Book file."
    Exit Sub
    End If
If Len(gstrSbook) < 10 Then
    'new s-book, then tempty textfield and add new
    gstrSbook = ""
    frmMain.txtBox.Text = ""
    frmMain.picPointer.Left = 0
    Call AddNewNote
    curPos = 0
    Exit Sub
    End If
If curPos = 0 Then Beep: curPos = curStart: Exit Sub
frmMain.txtBox.Enabled = True
'get current note
If curStop <> 0 Then
    tmp = Mid(gstrSbook, curStart, (curStop - curStart))
    Else
    tmp = Mid(gstrSbook, curStart)
    End If
'get header
Header = Left(tmp, 8)
frmMain.txtBox.Text = Mid(tmp, 9) & vbCrLf
'strip text
beginCut:
If Right(frmMain.txtBox.Text, 4) = vbCrLf & vbCrLf Then frmMain.txtBox.Text = Left(frmMain.txtBox.Text, Len(frmMain.txtBox.Text) - 2): GoTo beginCut
'get date

'update toolbar attrib
gstrOldNote = frmMain.txtBox.Text
curPos = curStart
lastFound = curStart
oldStart = curStart
oldStop = curStop
SetPosBar curPos
End Sub

Sub StoreCurrentNote()
Dim tmp As String
'if no s-book or text, don't store
If gstrCurrentFile = "" Then Exit Sub
If frmMain.txtBox.Text = "" And gstrOldNote = "" Then Exit Sub
If gstrOldNote = frmMain.txtBox.Text Or frmMain.txtBox.Text = vbCrLf & "   No Notes in this S-Book" Then
    gblnEditNewNote = False
    Exit Sub
    End If
gblnSbookChanged = True
If gblnEditNewNote Then
    'store new note in gstrSbook
    gblnEditNewNote = False
    gstrOldNote = frmMain.txtBox.Text
    curStart = Len(gstrSbook) + 1
    gstrSbook = gstrSbook & ComposeNote
Else
    If curStop = 0 And Len(gstrSbook) > 8 Then
        'store last note again
        gstrSbook = Left(gstrSbook, curStart - 1) & ComposeNote
        Else
        If Len(gstrSbook) > 8 Then
            'store note inside gstrSbook
            gstrSbook = Left(gstrSbook, curStart - 1) & ComposeNote & Mid(gstrSbook, curStop)
            Else
            'store first note
            gstrSbook = ComposeNote
        End If
    End If
End If
curPos = curStart
curStop = InStr(curPos + 1, gstrSbook, "[NOTE]" & vbCrLf)
gblnEditNewNote = False
gstrOldNote = frmMain.txtBox.Text
frmMain.txtBox.Locked = True
Call ShowCurrentNote
End Sub

Public Function ComposeNote() As String
'write note to string
Dim tmpT As String
Dim AttribA As String
Dim AttribL As String
Dim AttribO As String
Dim q As Long
Dim qS As Long
Dim k As Long
tmpT = RTrim(frmMain.txtBox.Text)
beginCut:
If Right(tmpT, 2) = vbCrLf Then tmpT = Left(tmpT, Len(tmpT) - 2): GoTo beginCut
tmpT = RTrim(tmpT)
'eliminate headers "[NOTE]" in text
qS = 1
Do
q = InStr(qS, tmpT, "[NOTE]")
If q > 0 Then Mid(tmpT, q, 6) = "[note]"
qS = q + 1
If qS > Len(tmpT) Then q = 0
Loop While q > 0
'compose update string
ComposeNote = "[NOTE]" & vbCrLf & tmpT & vbCrLf
End Function

Sub AddNewNote()
If gblnReadOnly = True Then
    MsgBox "The S-Book file is Read Only!", vbExclamation
    Exit Sub
    End If
frmMain.txtBox.Locked = False
gblnEditNewNote = True
gstrOldNote = ""
With frmMain
.txtBox.Enabled = True
.txtBox.Text = ""
.txtBox.SetFocus
End With
curStart = Len(gstrSbook)
curPos = curStart
curStop = 0
Call CheckButtons
End Sub

Sub DeleteCurrentNote()
Dim retval As Integer
If gblnReadOnly = True Then
    MsgBox "The S-Book file is Read Only!", vbExclamation
    Exit Sub
    End If
If curPos = 0 Then Beep: Exit Sub
retval = MsgBox("Are you sure you want to delete this note?", vbQuestion + vbYesNo, " Delete note")
If retval <> vbYes Then Exit Sub
gblnSbookChanged = True
If Len(gstrSbook) < 13 Then gstrSbook = "": Beep: Exit Sub
If curStop <> 0 Then
    gstrSbook = Left(gstrSbook, curStart - 1) + Mid(gstrSbook, curStop)
    Else
    gstrSbook = Left(gstrSbook, curStart - 1)
    End If
If Len(gstrSbook) > 8 Then
    curPos = curStart - 1
    Call GetNote(Up)
    If curPos = 0 Then curPos = Len(gstrSbook): Call GetNote(Down)
    End If
Call ShowCurrentNote
Call CheckButtons
End Sub

Sub SetArguments()
Dim tmp As String
Dim endPos As Integer
tmp = Trim(frmMain.Combo1.Text)
If tmp = "" Then gintArguments = 0: Exit Sub
endPos = InStr(1, tmp, " ")
gintArguments = 1
If endPos = 0 Then gstrArg1 = Trim(tmp): Exit Sub
gstrArg1 = Trim(Left(tmp, endPos - 1))
If Len(Mid(tmp, endPos)) = 1 Then Exit Sub
tmp = Trim(Mid(tmp, endPos + 1))
endPos = InStr(1, tmp, " ")
gintArguments = 2
If endPos = 0 Then gstrArg2 = Trim(tmp): Exit Sub
gstrArg2 = Trim(Left(tmp, endPos - 1))
If Len(Mid(tmp, endPos)) = 1 Then Exit Sub
gstrArg3 = Trim(Mid(tmp, endPos + 1))
gintArguments = 3
'more than 3 argument = msgbox
If InStr(1, gstrArg3, " ") Then gintArguments = 0: _
    MsgBox "Up to 3 keywords, seperated by a space, are allowed.", vbExclamation, " Find Notes"
End Sub

Public Sub FormResize()
Dim th
Dim ph
With frmMain
th = .Toolbar1.Height
ph = .picPosBar.Height
If .WindowState <> 1 And .ScaleHeight > 1700 And .ScaleWidth > 1700 Then
    If .ScaleHeight - th - 150 > 0 Then
        .txtBox.Height = .ScaleHeight - th - ph - 350
        .txtBox.Top = th
    .picPosBar.Width = .ScaleWidth
    .picPosBar.Top = .ScaleHeight - .Toolbar2.Height - ph
        End If
    .txtBox.Width = .ScaleWidth
    If .ScaleWidth > 1980 Then .Combo1.Width = .ScaleWidth - 1925
    SetPosBar curPos
    End If
End With
End Sub

Function StripFileName(ByVal strNameA As String) As String
Dim pos As Long
Do
pos = InStr(1, strNameA, "\")
strNameA = Right(strNameA, Len(strNameA) - pos)
Loop Until pos = 0
strNameA = Left(strNameA, Len(strNameA) - 4)
StripFileName = strNameA
End Function

Sub SaveWindowPos()
With frmMain
SaveSetting App.EXEName, "CONFIG", "WindowState", .WindowState
If .WindowState <> vbNormal Then Exit Sub
SaveSetting App.EXEName, "CONFIG", "WindowHeight", .Height
SaveSetting App.EXEName, "CONFIG", "WindowWidth", .Width
SaveSetting App.EXEName, "CONFIG", "WindowLeft", .Left
SaveSetting App.EXEName, "CONFIG", "WindowTop", .Top
End With
End Sub

Sub GetWindowPos()
With frmMain
.WindowState = CSng(GetSetting(App.EXEName, "Config", "Windowstate", vbNormal))
If .WindowState <> vbNormal Then Exit Sub
.Height = CSng(GetSetting(App.EXEName, "CONFIG", "WindowHeight", .Height))
.Width = CSng(GetSetting(App.EXEName, "CONFIG", "WindowWidth", .Width))
.Left = CSng(GetSetting(App.EXEName, "CONFIG", "WindowLeft", ((Screen.Width - .Width) / 2)))
.Top = CSng(GetSetting(App.EXEName, "CONFIG", "WindowTop", ((Screen.Height - .Height) / 2)))
End With
End Sub

Public Sub ToggleFind()
If gblnEditNewNote Then Exit Sub
gstrPrevSearch = ""
If Len(gstrSbook) < 13 Then Beep: Exit Sub
If gblnSearchBusy = False Then
    FindOn
    frmMain.mniFind.Caption = "&Stop Find"
    Else
    FindOff
    frmMain.mniFind.Caption = "&Find"
    End If
End Sub

Public Sub FindOn()
If gblnEditNewNote Then Exit Sub
gblnSearchBusy = True
gblnNewSearch = True
With frmMain
    .Combo1.Enabled = True
    .Combo1.BackColor = COL_WHITE
    .Toolbar2.Buttons("find").Image = "stopfind"
    .Toolbar2.Buttons("find").ToolTipText = " Stop Find [ESC] "
    .Toolbar2.Buttons("first").ToolTipText = " Find First (CTRL+Home) "
    .Toolbar2.Buttons("previous").ToolTipText = " Find Previous (CTRL+Page Down)"
    .Toolbar2.Buttons("next").ToolTipText = " Find Next (CTRL+Page Up) "
    .Toolbar2.Buttons("last").ToolTipText = " Find Last (CTRL+End) "
    .Combo1.SetFocus
End With
End Sub

Public Sub FindOff()
If gblnEditNewNote Then Exit Sub
gblnSearchBusy = False
With frmMain
    .Combo1.SelLength = 0
    .Combo1.Enabled = False
    .Combo1.BackColor = COL_GRAY
    .Toolbar2.Buttons("find").Image = "find"
    .Toolbar2.Buttons("find").ToolTipText = " Find F3 "
    .Toolbar2.Buttons("first").ToolTipText = " First (CTRL+Home) "
    .Toolbar2.Buttons("previous").ToolTipText = " Previous (CTRL+Page Down) "
    .Toolbar2.Buttons("next").ToolTipText = " Next (CTRL+Page Up) "
    .Toolbar2.Buttons("last").ToolTipText = " Last (CTRL+End) "
    .txtBox.SetFocus
End With
End Sub

Public Sub SetPosBar(posn As Long)
Dim barScale As Single
Dim bLen
bLen = Len(gstrSbook)
With frmMain
If bLen < 13 Then .picPointer.Left = 0: Exit Sub
If posn = 0 Then Exit Sub
barScale = (.picPosBar.Width - .picPointer.Width) / bLen
.picPointer.Left = posn * barScale
.picPosBar.Refresh
End With
End Sub

Public Sub CheckButtons()
With frmMain
If Len(gstrCurrentFile) > 8 Then  '<<<
    .Toolbar1.Buttons("save").Enabled = True
    .Toolbar1.Buttons("new").Enabled = True
    .Toolbar1.Buttons("delete").Enabled = True
    .Toolbar1.Buttons("lock").Enabled = True
    .Toolbar1.Buttons("print").Enabled = True
    .Toolbar1.Buttons("cut").Enabled = True
    .Toolbar1.Buttons("copy").Enabled = True
    .Toolbar1.Buttons("paste").Enabled = True
    .Toolbar1.Buttons("undo").Enabled = True
    .Toolbar2.Buttons("first").Enabled = True
    .Toolbar2.Buttons("previous").Enabled = True
    .Toolbar2.Buttons("next").Enabled = True
    .Toolbar2.Buttons("last").Enabled = True
    .Toolbar2.Buttons("find").Enabled = True
    .mnuRepair.Enabled = True
    .mnuDelet.Enabled = True
    .mniDelNote.Enabled = True
    .mniCut.Enabled = True
    .mniFind.Enabled = True
    .mniPrint.Enabled = True
    .mnuEdit.Enabled = True
    .txtBox.Enabled = True
    .mniSave.Enabled = True
    .mnuRepair.Enabled = True
    .mniClipboard.Enabled = True
    Else
    .Toolbar1.Buttons("save").Enabled = False
    .Toolbar1.Buttons("new").Enabled = False
    .Toolbar1.Buttons("delete").Enabled = False
    .Toolbar1.Buttons("lock").Enabled = False
    .Toolbar1.Buttons("print").Enabled = False
    .Toolbar1.Buttons("cut").Enabled = False
    .Toolbar1.Buttons("copy").Enabled = False
    .Toolbar1.Buttons("paste").Enabled = False
    .Toolbar1.Buttons("undo").Enabled = False
    .Toolbar2.Buttons("first").Enabled = False
    .Toolbar2.Buttons("previous").Enabled = False
    .Toolbar2.Buttons("next").Enabled = False
    .Toolbar2.Buttons("last").Enabled = False
    .Toolbar2.Buttons("find").Enabled = False
    .mniDelNote.Enabled = False
    .mnuEdit.Enabled = False
    .txtBox.Enabled = False
    .mniChangeKey.Enabled = False
    .txtBox.Locked = True
    .mniSave.Enabled = False
    .mnuRepair.Enabled = False
    .mniPrint.Enabled = False
    .mniClipboard.Enabled = False
    Exit Sub
    End If
'check for readonly
If gblnReadOnly = False Then
    .Toolbar1.Buttons("save").Enabled = True
    .Toolbar1.Buttons("new").Enabled = True
    .Toolbar1.Buttons("delete").Enabled = True
    .Toolbar1.Buttons("lock").Enabled = True
    .Toolbar1.Buttons("paste").Enabled = True
    .Toolbar1.Buttons("cut").Enabled = True
    .Toolbar1.Buttons("undo").Enabled = True
    .mniCut.Enabled = True
    .mniLock.Enabled = True
    .mniCopy.Enabled = True
    .mniPaste.Enabled = True
    .mniAddNote.Enabled = True
    .mniDelNote.Enabled = True
    .mnuDelet.Enabled = True
    .mniUndo.Enabled = True
    .mniSave.Enabled = True
    .mnuRepair.Enabled = True
    If gstrCurrentKey <> "" Then
        .mniChangeKey.Enabled = True
        Else
        .mniChangeKey.Enabled = False
        End If
    Else
    .Toolbar1.Buttons("save").Enabled = False
    .Toolbar1.Buttons("new").Enabled = False
    .Toolbar1.Buttons("delete").Enabled = False
    .Toolbar1.Buttons("lock").Enabled = False
    .Toolbar1.Buttons("paste").Enabled = False
    .Toolbar1.Buttons("cut").Enabled = False
    .Toolbar1.Buttons("undo").Enabled = False
    .mniCut.Enabled = False
    .mniLock.Enabled = False
    .mniPaste.Enabled = False
    .mniAddNote.Enabled = False
    .mniDelNote.Enabled = False
    .mnuDelet.Enabled = False
    .mniUndo.Enabled = False
    .mniSave.Enabled = False
    .mnuRepair.Enabled = False
    .txtBox.Locked = True
    .mniChangeKey.Enabled = False
    End If
If .txtBox.Locked = True Then
    .Toolbar1.Buttons("paste").Enabled = False
    .Toolbar1.Buttons("cut").Enabled = False
    .Toolbar1.Buttons("undo").Enabled = False
    .mniCut.Enabled = False
    .mniPaste.Enabled = False
    .mnuDelet.Enabled = False
    .mniUndo.Enabled = False
    .Toolbar1.Buttons("lock").Image = "lock"
    .Toolbar1.Buttons("lock").ToolTipText = " Edit Note "
    .mniLock.Caption = "&Edit Note"
    Else
    .Toolbar1.Buttons("paste").Enabled = True
    .Toolbar1.Buttons("cut").Enabled = True
    .Toolbar1.Buttons("undo").Enabled = True
    .mniCut.Enabled = True
    .mniPaste.Enabled = True
    .mnuDelet.Enabled = True
    .mniUndo.Enabled = True
    .Toolbar1.Buttons("lock").Image = "free"
    .Toolbar1.Buttons("lock").ToolTipText = " Lock Note "
    .mniLock.Caption = "&Lock note "
    End If
If Clipboard.GetText = "" Then
    .Toolbar1.Buttons("paste").Enabled = False
    .mniPaste.Enabled = False
    End If
If .txtBox.SelLength = 0 Then
    .Toolbar1.Buttons("copy").Enabled = False
    .mniCopy.Enabled = False
    .Toolbar1.Buttons("cut").Enabled = False
    .mniCut.Enabled = False
    End If
End With
End Sub

Public Sub ShowHelpFile()
frmHelp.Show
End Sub
