VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   " S-Book"
   ClientHeight    =   5325
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7395
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   7395
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlgOpen 
      Left            =   6600
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.PictureBox picPosBar 
      BackColor       =   &H00C0C0C0&
      Height          =   150
      Left            =   0
      ScaleHeight     =   90
      ScaleMode       =   0  'User
      ScaleWidth      =   7231.729
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4680
      Width           =   7215
      Begin VB.PictureBox picPointer 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FF0000&
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   0
         ScaleHeight     =   240
         ScaleWidth      =   135
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   0
         Width           =   158
      End
   End
   Begin MSComctlLib.Toolbar Toolbar2 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   2
      Top             =   4980
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   609
      ButtonWidth     =   609
      ButtonHeight    =   556
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "first"
            Object.ToolTipText     =   " First [CTRL+Home] "
            ImageKey        =   "first"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "previous"
            Object.ToolTipText     =   " Previous [CTRL+Page Down] "
            ImageKey        =   "previous"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "next"
            Object.ToolTipText     =   " Next [CTRL+Page Up] "
            ImageKey        =   "next"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "last"
            Object.ToolTipText     =   " Last [CTRL+End] "
            ImageKey        =   "last"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "find"
            Object.ToolTipText     =   " Find F3 "
            ImageKey        =   "find"
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.ComboBox Combo1 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmMain.frx":030A
         Left            =   1875
         List            =   "frmMain.frx":030C
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   0
         Width           =   3015
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6600
      Top             =   420
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":030E
            Key             =   "open"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0722
            Key             =   "save"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0C66
            Key             =   "new"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11AA
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12BE
            Key             =   "print"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1802
            Key             =   "find"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1916
            Key             =   "stopfind"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A2A
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F6E
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":24B2
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":29F6
            Key             =   "undo"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2F3A
            Key             =   "first"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":304E
            Key             =   "previous"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3162
            Key             =   "next"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3276
            Key             =   "last"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":338A
            Key             =   "lock"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":38CE
            Key             =   "free"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtBox 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   4170
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   360
      Width           =   6495
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   609
      ButtonWidth     =   609
      ButtonHeight    =   556
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "open"
            Object.ToolTipText     =   " Open S-Book "
            ImageKey        =   "open"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "save"
            Object.ToolTipText     =   " Save S-Book "
            ImageKey        =   "save"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "new"
            Object.ToolTipText     =   " Add New Note "
            ImageKey        =   "new"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "delete"
            Object.ToolTipText     =   " Delete Note "
            ImageKey        =   "delete"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "lock"
            ImageKey        =   "lock"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "print"
            Object.ToolTipText     =   " Print Current Note "
            ImageKey        =   "print"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cut"
            Object.ToolTipText     =   " Cut "
            ImageKey        =   "cut"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "copy"
            Object.ToolTipText     =   " Copy "
            ImageKey        =   "copy"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "paste"
            Object.ToolTipText     =   " Paste "
            ImageKey        =   "paste"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "undo"
            Object.ToolTipText     =   " Undo "
            ImageKey        =   "undo"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mniFile 
         Caption         =   "&Open..."
         Shortcut        =   ^B
      End
      Begin VB.Menu mniSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mniCreate 
         Caption         =   "&Create..."
      End
      Begin VB.Menu ln5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRepair 
         Caption         =   "&Restore..."
      End
      Begin VB.Menu mniClipboard 
         Caption         =   "&Send to Clipboard"
      End
      Begin VB.Menu ln8 
         Caption         =   "-"
      End
      Begin VB.Menu mniPrint 
         Caption         =   "&Print..."
         Shortcut        =   ^P
      End
      Begin VB.Menu ln2 
         Caption         =   "-"
      End
      Begin VB.Menu mniExit 
         Caption         =   "&Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mniUndo 
         Caption         =   "&Undo"
      End
      Begin VB.Menu ln3 
         Caption         =   "-"
      End
      Begin VB.Menu mniCut 
         Caption         =   "Cu&t"
      End
      Begin VB.Menu mniCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mniPaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnuDelet 
         Caption         =   "&Delete"
      End
      Begin VB.Menu ln4 
         Caption         =   "-"
      End
      Begin VB.Menu mniFind 
         Caption         =   "&Find"
         Shortcut        =   ^F
      End
      Begin VB.Menu ln6 
         Caption         =   "-"
      End
      Begin VB.Menu mniAddNote 
         Caption         =   "&New note"
         Shortcut        =   ^N
      End
      Begin VB.Menu mniDelNote 
         Caption         =   "D&elete note"
      End
      Begin VB.Menu ln7 
         Caption         =   "-"
      End
      Begin VB.Menu mniLock 
         Caption         =   "&Lock note"
         Shortcut        =   ^L
      End
   End
   Begin VB.Menu mnuExtra 
      Caption         =   "&Extra"
      Begin VB.Menu mniChangeKey 
         Caption         =   "&Change Passphrase"
      End
      Begin VB.Menu ln10 
         Caption         =   "-"
      End
      Begin VB.Menu mniFileType 
         Caption         =   "&Register .bsb Filetype"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mniHelp 
         Caption         =   "&Help Topics"
         Shortcut        =   {F1}
      End
      Begin VB.Menu ln9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInfo 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
Call FormResize
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Dim x As Integer
If gcontinueFind = True Then KeyAscii = 0: Exit Sub
Select Case KeyAscii
Case 14 'ctrl+n nieuw
    KeyAscii = 0
    StoreCurrentNote
    If gblnSearchBusy Then FindOff
    Call AddNewNote
Case 15 'ctrl+o open
    If gblnEditNewNote = True Then
        StoreCurrentNote
        curPos = Len(gstrSbook)
        Call FindNote(Down)
        End If
    If gblnSearchBusy = True Then ToggleFind
    x = OpenFile("")
    KeyAscii = 0
End Select
End Sub

Private Sub mniChangeKey_Click()
frmKeyDirect.Caption = " Change Passphrase for '" & StripFileName(gstrCurrentFile) & "'"
frmKeyDirect.lblCode.Caption = "Enter old Passphrase"
frmKeyDirect.Show (vbModal)
If gblnCancelKey = True Then Exit Sub
If gstrCurrentKey <> gstrKeyInput Then
    MsgBox "Wrong Passphrase!", vbCritical, " Change Passphrase"
    Exit Sub
    End If
frmKey.Caption = " Change Passphrase for '" & StripFileName(gstrCurrentFile) & "'"
frmKey.Show (vbModal)
If gblnCancelKey = True Then Exit Sub
If gstrCurrentKey <> gstrKeyInput Then
    gstrCurrentKey = gstrKeyInput
    gblnSbookChanged = True
    End If
End Sub

Private Sub mniClipboard_Click()
Dim ret As String
ret = MsgBox("Are you sure you want to send the complete S-Book without encryption to the clipboard?" & vbCrLf & "(the clipboard will be cleared on program exit)", vbQuestion + vbYesNo, "S-book to Clipboard")
If ret = vbNo Then Exit Sub
If gblnEditNewNote = True Then
    Call StoreCurrentNote
    Call FindNote(Up)
    End If
On Error Resume Next
Clipboard.Clear
Clipboard.SetText gstrSbook & vbCrLf
End Sub

Private Sub mniCreate_Click()
Call CreateFile
End Sub

Private Sub mniExit_Click()
Unload Me
End Sub

Private Sub mniFile_Click()
Dim x As Integer
If gcontinueFind = True Then Exit Sub
If gblnEditNewNote = True Then
    StoreCurrentNote
    curPos = Len(gstrSbook)
    Call FindNote(Down)
    End If
If gblnSearchBusy = True Then ToggleFind
x = OpenFile("")
End Sub

Private Sub mniFileType_Click()
Dim tmp As String
Dim ret As String
tmp = ReadKey(HKEY_CLASSES_ROOT, ".bsb", "", "")
'update menu
If tmp = "" Then
    Call MakeFileAssociation("bsb", App.Path, App.EXEName, "BIC encrypted S-Book File", App.Path & "\" & "bsb.ico")
    MsgBox "The .bsb filetype will be recognized after restarting the computer.", vbInformation, " S-Book"
    Me.mniFileType.Caption = "&Unregister .bsb filetype"
    Else
    ret = MsgBox("Are you sure you want to unregister the bsb filetype?", vbQuestion + vbYesNo, " S-Book")
    If ret = vbNo Then Exit Sub
    Call DeleteFileAssociation("bsb")
    MsgBox "The .bsb filetype is unregistred after restarting the computer.", vbInformation, " S-Book"
    Me.mniFileType.Caption = "&Register .bsb filetype"
    End If
End Sub

Private Sub mniFind_Click()
If gcontinueFind = True Then Exit Sub
ToggleFind
End Sub

Private Sub mniHelp_Click()
Call ShowHelpFile
End Sub

Private Sub mniLock_Click()
If Me.txtBox.Enabled = False Then Exit Sub
If Me.txtBox.Locked = True Then
    Me.txtBox.Locked = False
    Else
    Me.txtBox.Locked = True
    End If
CheckButtons
End Sub

Private Sub mniAddNote_Click()
If gcontinueFind = True Then Exit Sub
StoreCurrentNote
If gblnSearchBusy Then FindOff
Call AddNewNote
End Sub

Private Sub mniDelNote_Click()
If gcontinueFind = True Then Exit Sub
If gblnEditNewNote = True Then
    gblnEditNewNote = False
    curPos = Len(gstrSbook)
    Call FindNote(Down)
    Else
    DeleteCurrentNote
End If
End Sub

Private Sub mniPrint_Click()
If gcontinueFind = True Then Exit Sub
StoreCurrentNote
frmPage.Show (vbModal)
End Sub

Private Sub mniSave_Click()
Dim x As Integer
If gcontinueFind = True Then Exit Sub
If gblnEditNewNote = True Then
    Call StoreCurrentNote
    Call FindNote(Up)
    End If
x = SaveFile(gstrCurrentFile)
End Sub

Private Sub mniCut_Click()
If gcontinueFind = True Then Exit Sub
SendKeys "^{x}"
End Sub

Private Sub mniCopy_Click()
If gcontinueFind = True Then Exit Sub
SendKeys "^{c}"
End Sub

Private Sub mniPaste_Click()
If gcontinueFind = True Then Exit Sub
SendKeys "^{v}"
End Sub

Private Sub mniSaveAs_Click()
'
End Sub

Private Sub mniUndo_Click()
If gcontinueFind = True Then Exit Sub
SendKeys "^{z}"
End Sub

Private Sub mnuDelet_Click()
If gcontinueFind = True Then Exit Sub
SendKeys "^{Del}"
End Sub

Private Sub mnuFile_Click()
If gcontinueFind = True Then Exit Sub
End Sub

Private Sub mnuInfo_Click()
If gcontinueFind = True Then Exit Sub
frmAbout.Show (vbModal)
End Sub

Private Sub mnuRepair_Click()
If gcontinueFind = True Then Exit Sub
Dim x As Integer
Dim retval As Integer
retval = MsgBox("Do you want to undo all changes and restore the S-Book file as before opening?", vbYesNo + vbExclamation, " Restore S-Book file")
If retval <> vbYes Then Exit Sub
x = OpenFile(gstrCurrentFile)
curPos = Len(gstrSbook)
ShowCurrentNote
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim x As Integer
If gcontinueFind = True Then Exit Sub
If (Button.key <> "new" And Button.key <> "open") And frmMain.txtBox.Enabled = False Then Exit Sub
Select Case Button.key
Case "cut"
    SendKeys "^{x}"
Case "copy"
    SendKeys "^{c}"
Case "paste"
    SendKeys "^{v}"
Case "undo"
    SendKeys "^{z}"
Case "print"
    StoreCurrentNote
    frmPage.Show (vbModal)
Case "open"
    If gblnEditNewNote = True Then
        StoreCurrentNote
        curPos = Len(gstrSbook)
        Call FindNote(Down)
        End If
    If gblnSearchBusy = True Then ToggleFind
    x = OpenFile("")
Case "save"
    If gblnEditNewNote = True Then
        Call StoreCurrentNote
        Call FindNote(Up)
        End If
    x = SaveFile(gstrCurrentFile)
Case "new"
    If gblnReadOnly = True Then
        MsgBox "This S-Book is Read-Only!", vbExclamation
        Exit Sub
        End If
    StoreCurrentNote
    If gblnSearchBusy Then FindOff
    Call AddNewNote
Case "delete"
    If gblnReadOnly = True Then
        MsgBox "This S-Book is Read-Only!", vbExclamation
        Exit Sub
        End If
    If gblnEditNewNote = True Then
        gblnEditNewNote = False
        curPos = Len(gstrSbook)
        Call FindNote(Down)
    Else
        DeleteCurrentNote
    End If
Case "lock"
    If frmMain.txtBox.Enabled = False Then Exit Sub
    If frmMain.txtBox.Locked = True Then
        frmMain.txtBox.Locked = False
        Else
        frmMain.txtBox.Locked = True
        End If
End Select
CheckButtons
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
If gcontinueFind = True Then Me.Combo1.SetFocus: Exit Sub
If frmMain.txtBox.Enabled = False Then Exit Sub
StoreCurrentNote
gblnEditNewNote = False
Select Case Button.key
Case "first"
    curPos = 0
    Call FindNote(Up)
    frmMain.txtBox.Locked = True
Case "previous"
    Call FindNote(Down)
    frmMain.txtBox.Locked = True
Case "next"
    Call FindNote(Up)
    frmMain.txtBox.Locked = True
Case "last"
    curPos = Len(gstrSbook)
    Call FindNote(Down)
    frmMain.txtBox.Locked = True
Case "find"
    ToggleFind
    frmMain.txtBox.Locked = True
End Select
CheckButtons
If frmMain.txtBox.Enabled = False Then Exit Sub
If Not gblnSearchBusy Then
    frmMain.txtBox.SetFocus
    Else
    frmMain.Combo1.SetFocus
    End If
End Sub

Private Sub txtBox_KeyDown(KeyCode As Integer, Shift As Integer)
If gcontinueFind = True Then KeyCode = 0: Exit Sub
Select Case KeyCode
Case 114 'f3
    StoreCurrentNote
    Call FindOn
    KeyCode = 0
Case 27 'esc
    KeyCode = 0
    FindOff
Case 46 ' del
    If gblnReadOnly = True Then
        MsgBox "This S-Book is Read-Only!", vbExclamation
        KeyCode = 0
        Exit Sub
        End If
End Select
If Shift <> 2 Then Exit Sub
Select Case KeyCode
Case 36  'HOME
    StoreCurrentNote
    curPos = 0
    Call FindNote(Up)
Case 34  'PDOWN
    StoreCurrentNote
    Call FindNote(Down)
Case 33  'PUP
    StoreCurrentNote
    Call FindNote(Up)
Case 35  'END
    StoreCurrentNote
    curPos = Len(gstrSbook)
    Call FindNote(Down)
End Select
Call CheckButtons
End Sub

Private Sub txtBox_GotFocus()
If gcontinueFind = True And Me.Combo1.Enabled = True Then Me.Combo1.SetFocus
End Sub

Private Sub txtBox_KeyPress(KeyAscii As Integer)
Dim x As Integer
Dim retval As Integer
If gcontinueFind = True And KeyAscii = 27 Then
    'exit search
    FindOff
    KeyAscii = 0
    gcontinueFind = False
    Exit Sub
    End If
Select Case KeyAscii
Case 1 ' ctrl+A
    frmMain.txtBox.SelStart = 0
    frmMain.txtBox.SelLength = Len(frmMain.txtBox)
    KeyAscii = 0
Case 3 'ctrl+c copy
    'dummy
Case 4 'ctrl+d delete
    KeyAscii = 0
    If gblnEditNewNote = True Then
        gblnEditNewNote = False
        curPos = Len(gstrSbook)
        Call FindNote(Down)
    Else
        DeleteCurrentNote
    End If
Case 12 'ctrl+l (lock)
    If frmMain.txtBox.Locked = False Then
        frmMain.txtBox.Locked = True
        Else
        frmMain.txtBox.Locked = False
        End If
    KeyAscii = 0
Case 14 'ctrl+n new
    KeyAscii = 0
    StoreCurrentNote
    If gblnSearchBusy Then FindOff
    Call AddNewNote
Case 16 'ctrl+p print
    frmPage.Show vbModal
    KeyAscii = 0
Case 19 'ctrl+s save
    If gblnEditNewNote = True Then
        Call StoreCurrentNote
        Call FindNote(Up)
        End If
    x = SaveFile(gstrCurrentFile)
    KeyAscii = 0
Case 6 'ctrl+f find
    FindOn
    KeyAscii = 0
Case Else
    If gblnReadOnly = True Then
        MsgBox "This S-Book is Read-Only!", vbExclamation
        KeyAscii = 0
        Exit Sub
        End If
    If Me.txtBox.Locked = True Then
        retval = MsgBox("This note is locked!" & vbCrLf & vbCrLf & "Do you want to unlock this note for editing?", vbYesNo + vbQuestion)
        If retval = vbYes Then
            Call mniLock_Click
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Select
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
Dim x As Integer
If gcontinueFind = True Then KeyAscii = 0: Exit Sub
Select Case KeyAscii
Case 4 'ctrl+d
    KeyAscii = 0
    If gblnEditNewNote = True Then
        StoreCurrentNote
        curPos = Len(gstrSbook)
        Call FindNote(Down)
        End If
    DeleteCurrentNote
Case 12 'ctrl+l (lock)
    If frmMain.txtBox.Locked = False Then
        frmMain.txtBox.Locked = True
        Else
        frmMain.txtBox.Locked = False
        End If
    KeyAscii = 0
Case 16 'ctrl+p print
    frmPage.Show vbModal
    KeyAscii = 0
Case 19 'ctrl+s save
    If gblnEditNewNote = True Then
        Call StoreCurrentNote
        Call FindNote(Up)
        End If
    x = SaveFile(gstrCurrentFile)
    KeyAscii = 0
Case 21
    KeyAscii = 0
    Me.Combo1.SetFocus
Case 14 'ctrl+n (nieuw)
    KeyAscii = 0
    StoreCurrentNote
    If gblnSearchBusy Then FindOff
    Call AddNewNote
Case 15 'ctrl+o (open)
    KeyAscii = 0
    StoreCurrentNote
    If gblnSearchBusy = True Then ToggleFind
    x = OpenFile("")
Case 6
    KeyAscii = 0: ToggleFind
End Select
End Sub

Private Sub Combo1_GotFocus()
Me.Combo1.SelStart = 0
Me.Combo1.SelLength = Len(Me.Combo1.Text)
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 And gcontinueFind = True Then
    gcontinueFind = False
    KeyCode = 0
    Exit Sub
    End If
If gcontinueFind = True Then KeyCode = 0: Exit Sub
Select Case KeyCode
Case 27
    FindOff
    KeyCode = 0
Case 13, 114
    Call StoreCurrentNote
    If gblnNewSearch = True Then curPos = 0
    Call FindNote(Up)
    Me.Combo1.SelStart = 0
    Me.Combo1.SelLength = Len(Me.Combo1.Text)
    KeyCode = 0
End Select
If Shift <> 2 Then Exit Sub
Select Case KeyCode
Case 36  'HOME
    StoreCurrentNote
    curPos = 0
    Call FindNote(Up)
Case 34  'PDOWN
    StoreCurrentNote
    Call FindNote(Down)
Case 33  'PUP
    StoreCurrentNote
    Call FindNote(Up)
Case 35  'END
    StoreCurrentNote
    curPos = Len(gstrSbook)
    Call FindNote(Down)
End Select
End Sub

Private Sub txtBox_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Call CheckButtons
End Sub

Private Sub txtDate_GotFocus()
If Me.txtBox.Enabled = True Then Me.txtBox.SetFocus
End Sub

Private Sub picPosBar_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim newPos As Long
Call FindOff
If Len(gstrSbook) < 20 Then
    If Me.txtBox.Enabled = True Then Me.txtBox.SetFocus
    Exit Sub
    End If
newPos = Int((Len(gstrSbook) / Me.picPosBar.Width) * x)
If newPos > curPos Then
    curPos = newPos
    Call FindNote(Up)
    Else
    curPos = newPos
    Call FindNote(Down)
    End If
If Me.txtBox.Enabled = True Then Me.txtBox.SetFocus
End Sub

Private Sub picPosBar_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim newPos As Long
newPos = Int((Len(gstrSbook) / Me.picPosBar.Width) * x)
Me.picPosBar.ToolTipText = " " & Format(newPos, "###,###,###,###") & " "
End Sub

Private Sub picPointer_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Me.picPointer.ToolTipText = " >> " & Format(curPos, "###,###,###,###") & " << "
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim x As Integer
Dim retval As Integer
'no exit during search
If gcontinueFind = True Then Cancel = True: Exit Sub
' save changes
If gblnReadOnly = False Then
    'save note if changed
    Call StoreCurrentNote
    If gblnSbookChanged = True And gstrCurrentFile <> "" Then x = SaveFile(gstrCurrentFile)
    If x <> 0 Then
        retval = MsgBox("Do you want to exit without saving '" & StripFileName(gstrCurrentFile) & "' ?", vbYesNo + vbDefaultButton2 + vbQuestion, " Exit S-Book")
        If retval = vbNo Then Cancel = True: Exit Sub
        End If
    End If
'save all settings
SaveWindowPos
If gstrCurrentFile <> "" Then
    SaveSetting App.EXEName, "CONFIG", "File", gstrCurrentFile
    End If
Unload frmPage
Unload frmAbout
Unload frmKey
Unload frmKeyDirect
Unload frmHelp
On Error Resume Next
Clipboard.Clear
End Sub


