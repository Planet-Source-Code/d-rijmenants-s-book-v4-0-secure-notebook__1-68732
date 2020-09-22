VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPage 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "  Print Note"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   5415
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog dlgPage 
      Left            =   4800
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   4305
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdPrinter 
      Caption         =   "&Printer..."
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   4620
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   4005
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   5175
      Begin VB.Frame Frame1 
         Height          =   50
         Left            =   240
         TabIndex        =   24
         Top             =   2880
         Width           =   4695
      End
      Begin VB.CheckBox chkPrint 
         Caption         =   "&Print Header"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   3120
         Width           =   2655
      End
      Begin VB.TextBox txtPrintHead 
         Height          =   285
         Left            =   360
         TabIndex        =   7
         Top             =   3465
         Width           =   4455
      End
      Begin VB.PictureBox picBlad 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2175
         Left            =   2880
         ScaleHeight     =   2175
         ScaleWidth      =   1575
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   480
         Width           =   1575
         Begin VB.Image imgtekst 
            Height          =   1905
            Left            =   120
            Picture         =   "PageSetup.frx":0000
            Stretch         =   -1  'True
            Top             =   120
            Width           =   1320
         End
      End
      Begin VB.TextBox txtMarge 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   3
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   5
         Text            =   "0"
         Top             =   1920
         Width           =   375
      End
      Begin VB.TextBox txtMarge 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   2
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   4
         Text            =   "0"
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox txtMarge 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   1
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   3
         Text            =   "0"
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox txtMarge 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   0
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   2
         Text            =   "0"
         Top             =   840
         Width           =   375
      End
      Begin VB.PictureBox PicShadow 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2175
         Left            =   3000
         ScaleHeight     =   2175
         ScaleWidth      =   1575
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lblMarges 
         Alignment       =   2  'Center
         Caption         =   "Marges"
         Height          =   255
         Left            =   840
         TabIndex        =   21
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblVoorbeeld 
         Alignment       =   2  'Center
         Caption         =   "Example"
         Height          =   255
         Left            =   2880
         TabIndex        =   20
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Right"
         Height          =   255
         Index           =   7
         Left            =   360
         TabIndex        =   17
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Left"
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   16
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Bottom"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   15
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Top"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   14
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "%"
         Height          =   255
         Index           =   3
         Left            =   1680
         TabIndex        =   13
         Top             =   1965
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "%"
         Height          =   255
         Index           =   2
         Left            =   1680
         TabIndex        =   12
         Top             =   1605
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "%"
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   11
         Top             =   1245
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "%"
         Height          =   255
         Index           =   0
         Left            =   1680
         TabIndex        =   10
         Top             =   885
         Width           =   135
      End
   End
   Begin VB.CommandButton cmdAnnul 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   4620
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   4620
      Width           =   1215
   End
   Begin VB.Label lblPrinterName 
      Appearance      =   0  'Flat
      Caption         =   "Printer Name"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   4095
      Width           =   5175
   End
End
Attribute VB_Name = "frmPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
Me.lblPrinterName.Caption = Printer.DeviceName
Me.txtMarge(0).Text = Trim(Str(glngTmPrint))
Me.txtMarge(1).Text = Trim(Str(glngBmPrint))
Me.txtMarge(2).Text = Trim(Str(glngLmPrint))
Me.txtMarge(3).Text = Trim(Str(glngRmPrint))
Me.cmdOK.SetFocus
gblnChangeMarges = True
If Me.chkPrint.Value = 1 Then
    Me.txtPrintHead.Enabled = True
    Me.txtPrintHead.BackColor = &H80000005
    Else
    Me.txtPrintHead.Enabled = False
    Me.txtPrintHead.BackColor = &H8000000F
    End If
End Sub

Private Sub cmdOK_Click()
gblnCancelPrint = False
Dim tmp As String
Printer.FontName = frmMain.txtBox.Font
Printer.FontSize = 11
Printer.FontBold = False
Printer.FontItalic = False
glngTmPrint = Val(Me.txtMarge(0).Text)
glngBmPrint = Val(Me.txtMarge(1).Text)
glngLmPrint = Val(Me.txtMarge(2).Text)
glngRmPrint = Val(Me.txtMarge(3).Text)
If gblnChangeMarges = True Then
    SaveSetting App.EXEName, "CONFIG", "PrinterLeft", glngLmPrint
    SaveSetting App.EXEName, "CONFIG", "PrinterRight", glngRmPrint
    SaveSetting App.EXEName, "CONFIG", "PrinterTop", glngTmPrint
    SaveSetting App.EXEName, "CONFIG", "PrinterBottom", glngBmPrint
    SaveSetting App.EXEName, "CONFIG", "Headtext", Me.txtPrintHead.Text
    SaveSetting App.EXEName, "CONFIG", "Chkhead", Me.chkPrint.Value
    End If
If Me.chkPrint.Value = 1 Then
    tmp = Me.txtPrintHead.Text & vbCrLf & vbCrLf & frmMain.txtBox.Text
    Else
    tmp = frmMain.txtBox.Text
    End If
PrintString tmp, glngLmPrint, glngRmPrint, glngTmPrint, glngBmPrint
Me.Hide
End Sub

Private Sub cmdPrinter_Click()
On Error Resume Next
frmPage.dlgPage.Flags = &H4 Or &H100000
frmPage.dlgPage.ShowPrinter
Me.lblPrinterName.Caption = Printer.DeviceName
Call SetExample
End Sub

Private Sub txtMarge_Change(Index As Integer)
Call DrawMargesExample
End Sub

Private Sub txtMarge_GotFocus(Index As Integer)
Me.txtMarge(Index).SelStart = 0
Me.txtMarge(Index).SelLength = Len(Me.txtMarge(Index))
End Sub

Private Sub txtPrintHead_Change()
gblnChangeMarges = True
End Sub

Private Sub txtPrintHead_GotFocus()
Me.txtPrintHead.SelStart = 0
Me.txtPrintHead.SelLength = Len(Me.txtPrintHead.Text)
End Sub

Private Sub chkPrint_Click()
gblnChangeMarges = True
If Me.chkPrint.Value = 1 Then
    Me.txtPrintHead.Enabled = True
    Me.txtPrintHead.BackColor = &H80000005
    Else
    Me.txtPrintHead.Enabled = False
    Me.txtPrintHead.BackColor = &H8000000F
    End If
End Sub

Private Sub cmdAnnul_Click()
gblnCancelPrint = True
If Not gblnPrintBusy Then
    Me.Hide
    End If
End Sub

Private Sub SetExample()
With Me
If Printer.Orientation = vbPRORPortrait Then
    'portrait
    .picBlad.Top = 480
    .picBlad.Left = 2840
    .picBlad.Width = 1575
    .picBlad.Height = 2175
    .PicShadow.Width = 1575
    .PicShadow.Height = 2175
    .PicShadow.Top = 600
    .PicShadow.Left = 2960
    Else
    'landscape
    .picBlad.Top = 720
    .picBlad.Left = 2480
    .picBlad.Width = 2175
    .picBlad.Height = 1575
    .PicShadow.Width = 2175
    .PicShadow.Height = 1575
    .PicShadow.Top = 840
    .PicShadow.Left = 2600
    End If
End With
Call DrawMargesExample
End Sub

Private Sub DrawMargesExample()
Dim tm
Dim bm
Dim lm
Dim rm
Dim SheetWidht
Dim SheetHeight
tm = Val(Me.txtMarge(0).Text)
bm = Val(Me.txtMarge(1).Text)
lm = Val(Me.txtMarge(2).Text)
rm = Val(Me.txtMarge(3).Text)
If tm + bm > 95 Then
    tm = 5
    bm = 5
    Me.txtMarge(0).Text = "5"
    Me.txtMarge(1).Text = "5"
    End If
If lm + rm > 95 Then
    lm = 5
    rm = 5
    Me.txtMarge(2).Text = "5"
    Me.txtMarge(3).Text = "5"
    End If
SheetWidht = Me.picBlad.Width
SheetHeight = Me.picBlad.Height
Me.imgtekst.Width = Int(SheetWidht / 100 * (100 - lm - rm))
Me.imgtekst.Height = Int(SheetHeight / 100 * (100 - tm - bm))
Me.imgtekst.Top = Int(SheetHeight / 100 * tm)
Me.imgtekst.Left = Int(SheetWidht / 100 * lm)
gblnChangeMarges = True
End Sub

