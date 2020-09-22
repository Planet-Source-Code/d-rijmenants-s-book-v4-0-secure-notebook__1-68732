Attribute VB_Name = "Modsubs"

'-------------------------------------
'                                    '
'     Memo v3.0 memo organizer       '
'                                    '
'  (c) RDDS - RD Data Systems 2002   '
'                                    '
'-------------------------------------

Public gstrMemo         As String
Public gstrCurrFileName As String
Public gstrOldMemo      As String
Public gstrPrevSearch   As String

Public gblnMemoChanged  As Boolean
Public gblnReadOnly     As Boolean
Public gblnSearchBusy   As Boolean
Public gblnEditNewMemo      As Boolean
Public gblnNewSearch    As Boolean
Public blnCancelPrint   As Boolean
Public gblnSearchHit    As Boolean
Public gblnLock          As Boolean
Public gblnOldLock      As Boolean
Public gblnAlarm         As Boolean
Public gblnOldAlarm     As Boolean
Public continueFind     As Boolean

Public gintTextCol      As Integer
Public gintTextFont     As Integer
Public gintTextSize     As Integer

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

Public lmPrint          As Long
Public rmPrint          As Long
Public tmPrint          As Long
Public bmPrint          As Long

Public Enum Direction
    Up
    Down
End Enum

Public Const colWhite = &H8000000E
Public Const colGray = &H8000000B

Sub Main()
Dim W1 As Long
Dim W2 As Long
Dim wTxt As String
Load frmMain
Load frmFile
Load frmPage
Load frmInfo
'settings menu
On Error Resume Next
gintTextCol = Val(GetSetting(App.EXEName, "CONFIG", "TextCol", "1"))
gintTextSize = Val(GetSetting(App.EXEName, "CONFIG", "TextSize", "3"))
gintTextFont = Val(GetSetting(App.EXEName, "CONFIG", "TextFont", "1"))
If gintTextCol < 1 Or gintTextCol > 7 Then gintTextCol = 1
If gintTextSize < 1 Or gintTextSize > 6 Then gintTextSize = 3
If gintTextFont < 1 Or gintTextFont > 4 Then gintTextFont = 1
frmMain.mniCol(gintTextCol).Checked = True
frmMain.mniSize(gintTextSize).Checked = True
frmMain.mniFont(gintTextFont).Checked = True
'view textfield
Call SetColor
With frmMain
    txtField1.Font = Mid(.mniFont(Index).Caption, 2)
    txtField1.FontSize = Val(.mniSize(Index).Caption)
    txtField1.FontBold = False
    txtField1.FontItalic = False
End With
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
tmPrint = Val(GetSetting(App.EXEName, "CONFIG", "PrinterTop", "5"))
bmPrint = Val(GetSetting(App.EXEName, "CONFIG", "PrinterBottom", "5"))
lmPrint = Val(GetSetting(App.EXEName, "CONFIG", "PrinterLeft", "5"))
rmPrint = Val(GetSetting(App.EXEName, "CONFIG", "PrinterRight", "5"))
frmPage.txtPrintHead = GetSetting(App.EXEName, "CONFIG", "Headtext", "")
frmPage.chkPrint.Value = Val(GetSetting(App.EXEName, "CONFIG", "Chkhead", "false"))
' current memo file
gstrCurrFileName = GetSetting(App.EXEName, "CONFIG", "File", "")
If Command <> "" Then
    If Right(Command, 4) = ".mem" Then
        gstrCurrFileName = Command
        End If
    End If
'check file
getMemoFile:
If gstrCurrFileName = "" Or Dir(gstrCurrFileName) = "" Then
    gstrCurrFileName = ""
    'open-file window if no file
    frmFile.Show (vbModal)
    End If
If gstrCurrFileName = "" Then
    'no file, exit prg
    Unload frmMain
    Exit Sub
    End If
Call GetWindowPos
frmMain.Show
frmMain.txtField1.Text = vbCrLf & "   BESTAND WORDT GEOPEND..."
frmMain.Refresh
'load file
x = OpenMemoFile(gstrCurrFileName)
'if error, get again file
If x <> 0 Then gstrCurrFileName = "": GoTo getMemoFile
'show last memo
curPos = Len(gstrMemo)
Call FindMemo(Down)
Call GetWindowPos
frmMain.Show
'search important memo's
frmMain.Combo1.AddItem ("Belangrijke memo's")
frmMain.Combo1.Text = "Belangrijke memo's"
W1 = InStr(1, gstrMemo, "[~A")
If W1 <> 0 Then
    'more than one?
    W2 = InStr(W1 + 1, gstrMemo, "[~A")
    If W2 = 0 Then
        wTxt = "Gelieve deze belangrijke memo eerst te lezen !"
        Else
        wTxt = "Gelieve deze belangrijke memo's eerst te doorbladeren !"
        End If
    FindOn
    Call FindMemo(Up)
    MsgBox wTxt, vbExclamation, " Memo"
    frmMain.Combo1.SetFocus
    'only one = no search
    If W2 = 0 Then frmMain.Combo1.Text = "": FindOff
    End If
End Sub

Public Sub GetMemo(UpDn As Direction)
Dim sta1 As Long
Dim stp1 As Long
Dim revPos As Long
Dim allArg As String
If UpDn = Direction.Up Then
    'search up
    sta1 = InStr(curPos + 1, gstrMemo, "[~")
    If sta1 > 0 Then stp1 = InStr(sta1 + 1, gstrMemo, "[~")
    Else
    'search down
    For revPos = curPos - 1 To 1 Step -1
        If revPos < 1 Then revPos = 0: Exit For
        If Mid(gstrMemo, revPos, 2) = "[~" Then Exit For
        Next revPos
    sta1 = revPos
    If sta1 < 0 Then sta1 = 0
    stp1 = InStr(sta1 + 1, gstrMemo, "[~")
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


Public Sub FindMemo(UpDn As Direction)
Dim blnFound As Boolean
Dim tmp As String
Dim blnA As Boolean
oldStart = curStart
oldStop = curStop
If frmMain.Combo1.Text = "!" Then frmMain.Combo1.Text = "Belangrijke memo's"
'continue same search?
If gstrPrevSearch = frmMain.Combo1.Text Then
    gblnNewSearch = False
    Else
    gblnNewSearch = True
    gstrPrevSearch = frmMain.Combo1.Text
    End If
' check for limit begin or end
If UpDn = Direction.Down And curStop > Len(gstrMemo) + 1 Then Beep: Exit Sub
If UpDn = Direction.Up And curStart < 1 Then Beep: Exit Sub
If Not gblnSearchBusy Then
    Call GetMemo(UpDn) '<<<<<<
    ShowCurrentMemo
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
If gblnNewSearch = True And UpDn = Direction.Down Then curPos = Len(gstrMemo): gblnNewSearch = False
Call SetArguments
Screen.MousePointer = 11
frmMain.Combo1.SetFocus
continueFind = True
Do
'get next memo
Call GetMemo(UpDn)
If curPos = 0 Then continueFind = False: Exit Do
'check if match
If curStop <> 0 Then
    tmp = UCase(Mid(gstrMemo, curStart, (curStop - curStart)))
    Else
    tmp = UCase(Mid(gstrMemo, curStart))
    End If
blnFound = FindMatch(tmp)
'make interrupt possible when searching next hit
DoEvents
If continueFind = False Then blnFound = True
Loop While Not blnFound
Screen.MousePointer = 0
If gintArguments = 0 Then blnFound = True
If blnFound Then
    'show hit
    Call ShowCurrentMemo
    gblnSearchHit = True
    Else
    'not found,set argument for search next hit
    If gintArguments = 3 Then allArg = gstrArg1 & " + " & gstrArg2 & " + " & gstrArg3
    If gintArguments = 2 Then allArg = gstrArg1 & " + " & gstrArg2
    If gintArguments = 1 Then allArg = gstrArg1
    If frmMain.Combo1.Text = "Belangrijke memo's" Then allArg = "Belangrijke memo's"
    If gblnSearchHit = True Then
        'if previous was hit
        MsgBox "Zoeken naar  " & allArg & "  is voltooid.", vbInformation, " Memo Zoeken"
        Else
        'if no previours hit
        MsgBox allArg & "  niet gevonden.", vbInformation, " Memo Zoeken"
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
continueFind = False
End Sub

Private Function FindMatch(strA As String) As Boolean
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

Sub ShowCurrentMemo()
Dim tmpCr
Dim Header As String
If Len(gstrMemo) < 13 Then
    'no memos = textfield empty
    gstrMemo = ""
    frmMain.txtField1.Text = vbCrLf & "   Bestand bevat geen memo's."
    frmMain.txtField1.Enabled = False
    frmMain.txtDate.Text = ""
    frmMain.txtDate.Enabled = False
    frmMain.picPointer.Left = 0
    curPos = 0
    Exit Sub
    End If
If curPos = 0 Then Beep: curPos = curStart: Exit Sub
frmMain.txtField1.Enabled = True
frmMain.txtDate.Enabled = True
'get current memo
If curStop <> 0 Then
    tmp = Mid(gstrMemo, curStart, (curStop - curStart))
    Else
    tmp = Mid(gstrMemo, curStart)
    End If
'get header
Header = Left(tmp, 17)
'get all after header [~XXX~31/12/2000] & 13/10
frmMain.txtField1.Text = Mid(tmp, 20) & vbCrLf
'strip text
beginCut:
If Right(frmMain.txtField1.Text, 4) = vbCrLf & vbCrLf Then frmMain.txtField1.Text = Left(frmMain.txtField1.Text, Len(frmMain.txtField1.Text) - 2): GoTo beginCut
'get date
frmMain.txtDate.Text = Mid(Header, 7, 10)
'get attrib alarm
If Mid(Header, 3, 1) = "A" Then
    gblnAlarm = True
    Else
    gblnAlarm = False
    End If
'get attrib Lock
If Mid(Header, 4, 1) = "L" Then
    gblnLock = True
    Else
    gblnLock = False
    End If
'update toolbar attrib
CheckAttrib
gblnOldLock = gblnLock
gblnOldAlarm = gblnAlarm
gstrOldMemo = frmMain.txtField1.Text
curPos = curStart
lastFound = curStart
oldStart = curStart
oldStop = curStop
SetPosBar curPos
End Sub

Sub StoreCurrentMemo()
Dim tmp As String
'if no text, don't store
If frmMain.txtField1.Text = "" And gstrOldMemo = "" Then Exit Sub
If (gstrOldMemo = frmMain.txtField1.Text And gblnLock = gblnOldLock And gblnAlarm = gblnOldAlarm) _
    Or frmMain.txtField1.Text = vbCrLf & "   Bestand bevat geen memo's." Then
    gblnEditNewMemo = False
    Exit Sub
    End If
gblnMemoChanged = True
If gblnEditNewMemo Then
    'store new memo in gstrmemo
    gblnEditNewMemo = False
    gstrOldMemo = frmMain.txtField1.Text
    curStart = Len(gstrMemo) + 1
    gstrMemo = gstrMemo & vbCrLf & ComposeMemo
Else
    If curStop = 0 And Len(gstrMemo) > 17 Then
        'store last memo again
        gstrMemo = Left(gstrMemo, curStart - 1) & ComposeMemo
        Else
        If Len(gstrMemo) > 17 Then
            'store memo inside gstrmemo
            gstrMemo = Left(gstrMemo, curStart - 1) & ComposeMemo & vbCrLf & Mid(gstrMemo, curStop)
            Else
            'store first memo
            gstrMemo = ComposeMemo
        End If
    End If
End If
curPos = curStart
curStop = InStr(curPos + 1, gstrMemo, "[~")
gblnEditNewMemo = False
gstrOldMemo = frmMain.txtField1.Text
Call ShowCurrentMemo
End Sub

Public Function ComposeMemo() As String
'write memo to string
Dim tmpT As String
Dim AttribA As String
Dim AttribL As String
Dim AttribO As String
Dim q As Long
Dim qS As Long
Dim k As Long
tmpT = RTrim(frmMain.txtField1.Text)
beginCut:
If Right(tmpT, 2) = vbCrLf Then tmpT = Left(tmpT, Len(tmpT) - 2): GoTo beginCut
tmpT = RTrim(tmpT)
'get attrib alarm
If gblnAlarm = True Then AttribA = "A" Else AttribA = "X"
'get attrib lock
If gblnLock = True Then AttribL = "L" Else AttribL = "X"
'get attrib other options
AttribO = "X"
'eliminate headers "[~" in text
qS = 1
Do
q = InStr(qS, tmpT, "[~")
If q > 0 Then Mid(tmpT, q, 2) = "[-"
qS = q + 1
If qS > Len(tmpT) Then q = 0
Loop While q > 0
'compose update string
ComposeMemo = "[~" & AttribA & AttribL & AttribO & "~" & Format(Date, "dd/mm/yyyy") & "]" & vbCrLf & tmpT & vbCrLf
End Function

Sub AddNewMemo()
gblnEditNewMemo = True
gstrOldMemo = ""
gblnLock = False
gblnOldLock = False
Call CheckAttrib
With frmMain
.txtField1.Enabled = True
.txtField1.Text = ""
.txtDate.Enabled = True
.txtDate.Text = Format(Date, "dd/mm/yyyy")
.txtField1.SetFocus
End With
curStart = Len(gstrMemo)
curPos = curStart
curStop = 0
Call CheckButtons
End Sub

Sub DeleteCurrentMemo()
If gblnReadOnly = True Then
    MsgBox "Dit memo-bestand kan enkel gelezen worden !", vbExclamation
    Exit Sub
    End If
If gblnLock = True Then
    MsgBox "Deze memo is vergrendeld en kan enkel gelezen worden !", vbExclamation
    Exit Sub
    End If
If curPos = 0 Then Beep: Exit Sub
retval = MsgBox("Bent u zeker dat u de memo van " & frmMain.txtDate.Text & " wilt wissen?", vbQuestion + vbYesNo, " Wissen memo")
If retval <> vbYes Then Exit Sub
gblnMemoChanged = True
If Len(gstrMemo) < 13 Then gstrMemo = "": Beep: Exit Sub
If curStop <> 0 Then
    gstrMemo = Left(gstrMemo, curStart - 1) + Mid(gstrMemo, curStop)
    Else
    gstrMemo = Left(gstrMemo, curStart - 1)
    End If
If Len(gstrMemo) > 13 Then
    curPos = curStart - 1
    Call GetMemo(Up)
    If curPos = 0 Then curPos = Len(gstrMemo): Call GetMemo(Down)
    End If
Call ShowCurrentMemo
Call CheckButtons
End Sub

Sub SetArguments()
Dim tmp As String
Dim endPos As Integer
tmp = Trim(frmMain.Combo1.Text)
If tmp = "Belangrijke memo's" Or tmp = "!" Then gintArguments = 1: gstrArg1 = "[~A": Exit Sub
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
    MsgBox "Maximum 3 trefwoorden, gescheiden door spaties toegestaan." _
    , vbExclamation, " Memo zoeken"
End Sub

Public Sub FormResize()
Dim th
Dim ph
With frmMain
th = .Toolbar1.Height
ph = .picPosBar.Height
If .WindowState <> 1 And .ScaleHeight > 1700 And .ScaleWidth > 1700 Then
    If .ScaleHeight - th - 150 > 0 Then
        .txtField1.Height = .ScaleHeight - th - ph - 350
        .txtField1.Top = th
    .picPosBar.Width = .ScaleWidth
    .picPosBar.Top = .ScaleHeight - .Toolbar2.Height - ph
        End If
    .txtField1.Width = .ScaleWidth
    If .ScaleWidth > 1980 Then .Combo1.Width = .ScaleWidth - 1925
    SetPosBar curPos
    End If
End With
End Sub

Public Function OpenMemoFile(strName As String) As Integer
Dim fAttrib As Integer
Screen.MousePointer = 11
On Error GoTo ErrorHandling
FileNr = FreeFile
gblnReadOnly = False
'
Open strName For Input As #FileNr
frmMain.txtField1.Text = vbCrLf & "   BESTAND WORDT GEOPEND (" & Format(Int(LOF(FileNr) / 1024), "###,###,##0") & " kB)..."
frmMain.txtField1.Refresh
gstrMemo = Input(LOF(FileNr), FileNr)
Close #FileNr
'
fAttrib = GetAttr(strName)
If fAttrib And 1 Then gblnReadOnly = True
'
cutCRLF:
If Right(gstrMemo, 2) = vbCrLf Then gstrMemo = Left(gstrMemo, Len(gstrMemo) - 2): GoTo cutCRLF
gstrMemo = gstrMemo & vbCrLf
'
Screen.MousePointer = 0
frmMain.txtField1.Text = ""
frmMain.Caption = NameToCaption(strName)
gstrCurrFileName = strName
'
curPos = Len(gstrMemo)
Call GetMemo(Down)
Call ShowCurrentMemo
Call CheckButtons
OpenMemoFile = 0
Exit Function
'
ErrorHandling:
    frmMain.txtField1.Text = ""
    MsgBox NameToCaption(strName) & " kan niet geopend worden." & vbCr & vbCr & "Error: " & Error, vbCritical
    Close #FileNr
    Err.Clear
    Screen.MousePointer = 0
    OpenMemoFile = Err.Number
End Function

Public Function SaveMemoFile(strName As String) As Integer
Dim FileNr
Dim WriteText As String
If strName = "" Then Exit Function
On Error GoTo errorHandle
Screen.MousePointer = 11
FileNr = FreeFile
WriteText = gstrMemo
cutCRLF:
If Right(gstrMemo, 2) = vbCrLf Then gstrMemo = Left(gstrMemo, Len(gstrMemo) - 2): GoTo cutCRLF
'
Open strName For Output As #FileNr
Print #FileNr, WriteText
Close #FileNr
'
Screen.MousePointer = 0
SaveMemoFile = 0
Exit Function
'
errorHandle:
    MsgBox "Opslaan memo-bestand mislukt." & vbCr & vbCr & "Error: " & Error, 48
    Close #FileNr
    Screen.MousePointer = 0
    SaveMemoFile = Err.Number
End Function

Function NameToCaption(strNameA) As String
xx = strNameA
Do
pos = InStr(1, xx, "\")
xx = Right(xx, Len(xx) - pos)
Loop Until pos = 0
xx = Left(xx, Len(xx) - 4)
NameToCaption = " Memo - " + xx
If gblnReadOnly = True Then NameToCaption = NameToCaption + " [ Alleen lezen ]"
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
If gblnEditNewMemo Then Exit Sub
gstrPrevSearch = ""
If Len(gstrMemo) < 13 Then Beep: Exit Sub
If gblnSearchBusy = False Then
    FindOn
    Else
    FindOff
    End If
End Sub

Public Sub FindOn()
If gblnEditNewMemo Then Exit Sub
gblnSearchBusy = True
gblnNewSearch = True
With frmMain
    .Combo1.Enabled = True
    .Combo1.BackColor = colWhite
    .Toolbar2.Buttons("find").Image = "stopfind"
    .Toolbar2.Buttons("find").ToolTipText = " Stop zoeken [ESC] "
    .Toolbar2.Buttons("first").ToolTipText = " Eerste zoeken [Home] "
    .Toolbar2.Buttons("previous").ToolTipText = " Vorige zoeken [Page Down]"
    .Toolbar2.Buttons("next").ToolTipText = " Volgende zoeken [Page Up] "
    .Toolbar2.Buttons("last").ToolTipText = " Laatste zoeken [End] "
    .Combo1.SetFocus
End With
End Sub

Public Sub FindOff()
If gblnEditNewMemo Then Exit Sub
gblnSearchBusy = False
With frmMain
    .Combo1.SelLength = 0
    .Combo1.Enabled = False
    .Combo1.BackColor = colGray
    .Toolbar2.Buttons("find").Image = "find"
    .Toolbar2.Buttons("find").ToolTipText = " Zoeken F3 "
    .Toolbar2.Buttons("first").ToolTipText = " Eerste [Home] "
    .Toolbar2.Buttons("previous").ToolTipText = " Vorige [Page Down] "
    .Toolbar2.Buttons("next").ToolTipText = " Volgende [Page Up] "
    .Toolbar2.Buttons("last").ToolTipText = " Laatste [End] "
    .txtField1.SetFocus
End With
End Sub

Public Sub CheckAttrib()
With frmMain
If gblnLock = True Then
    .Toolbar1.Buttons("lock").Image = "lock"
    .Toolbar1.Buttons("lock").ToolTipText = " Memo vrijgeven "
    .txtField1.Locked = True
    Else
    .Toolbar1.Buttons("lock").Image = "unlock"
    .Toolbar1.Buttons("lock").ToolTipText = " Memo vergrendelen "
    .txtField1.Locked = False
    End If
If gblnAlarm = True Then
    .Toolbar1.Buttons("alarm").Image = "alarm"
    Else
    .Toolbar1.Buttons("alarm").Image = "alarmoff"
    End If
If gblnReadOnly = True Then
    .txtField1.Locked = True
    Else
    .txtField1.Locked = False
    End If
End With
End Sub

Public Sub SetPosBar(posn As Long)
Dim barScale As Single
bl = Len(gstrMemo)
With frmMain
If bl < 13 Then .picPointer.Left = 0: Exit Sub
If posn = 0 Then Exit Sub
barScale = (.picPosBar.Width - .picPointer.Width) / bl
.picPointer.Left = posn * barScale
.picPosBar.Refresh
End With
End Sub

Public Sub CheckButtons()
With frmMain
If Len(gstrMemo) > 19 Or gblnEditNewMemo = True Then '<<<
    .Toolbar1.Buttons("delete").Enabled = True
    .Toolbar1.Buttons("lock").Enabled = True
    .Toolbar1.Buttons("print").Enabled = True
    .Toolbar1.Buttons("cut").Enabled = True
    .Toolbar1.Buttons("copy").Enabled = True
    .Toolbar1.Buttons("paste").Enabled = True
    .Toolbar1.Buttons("undo").Enabled = True
    .Toolbar1.Buttons("alarm").Enabled = True
    .Toolbar2.Buttons("first").Enabled = True
    .Toolbar2.Buttons("previous").Enabled = True
    .Toolbar2.Buttons("next").Enabled = True
    .Toolbar2.Buttons("last").Enabled = True
    Else
    .Toolbar1.Buttons("delete").Enabled = False
    .Toolbar1.Buttons("lock").Enabled = False
    .Toolbar1.Buttons("print").Enabled = False
    .Toolbar1.Buttons("cut").Enabled = False
    .Toolbar1.Buttons("copy").Enabled = False
    .Toolbar1.Buttons("paste").Enabled = False
    .Toolbar1.Buttons("undo").Enabled = False
    .Toolbar1.Buttons("alarm").Enabled = False
    .Toolbar2.Buttons("first").Enabled = False
    .Toolbar2.Buttons("previous").Enabled = False
    .Toolbar2.Buttons("next").Enabled = False
    .Toolbar2.Buttons("last").Enabled = False
    End If
If Len(gstrMemo) > 12 Then
    .Toolbar2.Buttons("find").Enabled = True
    Else
    .Toolbar2.Buttons("find").Enabled = False
    End If
End With
End Sub

Public Sub SaveCurrentMemoFile()
Dim x As Integer
Call StoreCurrentMemo
If gblnReadOnly = True Then Exit Sub
If gblnMemoChanged = False Then Exit Sub
frmMain.txtField1.Text = vbCrLf & "  BESTAND WORDT BIJGEWERKT..."
frmMain.txtField1.Refresh
x = SaveMemoFile(gstrCurrFileName)
If x = 0 Then
    gblnMemoChanged = False
    gstrOldMemo = ""
    End If
frmMain.txtField1.Text = ""
Call ShowCurrentMemo
End Sub

Public Sub SetColor()
With frmMain.txtField1
Select Case gintTextCol
Case 1 'Zwart wit
    .BackColor = &HFFFFFF
    .ForeColor = &H0
Case 2 'Lilablauw
    .BackColor = &HFFDDE0
    .ForeColor = &HC00000
Case 3 'Pastelgeel
    .BackColor = &HCEFFFE
    .ForeColor = &HFF0000
Case 4 'Pastelgroen
    .BackColor = &HDFFFDD
    .ForeColor = &HC00000
Case 5 'Woestijnrood
    .BackColor = &H80C0FF
    .ForeColor = &HFF&
Case 6 'Contrast zwart
    .BackColor = &H0
    .ForeColor = &HFF00&
Case 7 'Contrast blauw
    .BackColor = &HFF0000
    .ForeColor = &HFFFFFF
End Select
End With
End Sub
