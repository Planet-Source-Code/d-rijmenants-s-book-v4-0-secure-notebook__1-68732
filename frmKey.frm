VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmKey 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Save Current S-Book"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4785
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   195
      Left            =   1995
      TabIndex        =   7
      Top             =   855
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CheckBox Check1 
      Caption         =   "&Hide typing"
      Height          =   225
      Left            =   210
      TabIndex        =   2
      Top             =   2220
      Width           =   1800
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   2115
      Width           =   1170
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2100
      TabIndex        =   3
      Top             =   2115
      Width           =   1170
   End
   Begin VB.TextBox txtConfirm 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   210
      TabIndex        =   1
      Top             =   1590
      Width           =   4320
   End
   Begin VB.TextBox txtCode 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   210
      TabIndex        =   0
      Top             =   435
      Width           =   4320
   End
   Begin VB.Label lblQuality 
      Alignment       =   1  'Right Justify
      Caption         =   "Passphrase Quality"
      Height          =   225
      Left            =   210
      TabIndex        =   8
      Top             =   855
      Width           =   1695
   End
   Begin VB.Label lblConfirm 
      Caption         =   "Confirm the Passphrase"
      Height          =   225
      Left            =   210
      TabIndex        =   6
      Top             =   1380
      Width           =   4320
   End
   Begin VB.Label lblCode 
      Caption         =   "Enter the Passphrase"
      Height          =   225
      Left            =   210
      TabIndex        =   5
      Top             =   225
      Width           =   3900
   End
End
Attribute VB_Name = "frmKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
gblnCancelKey = False
Me.cmdOK.Enabled = False
Me.txtCode.Text = ""
Me.txtConfirm.Text = ""
Me.txtCode.PasswordChar = "*"
Me.txtConfirm.PasswordChar = "*"
Me.Check1.Value = 1
Me.ProgressBar1.Value = 0
Me.txtCode.SetFocus
If gstrKeyInput <> "" Then
    Me.txtCode.SelStart = 0
    Me.txtCode.SelLength = Len(Me.txtCode.Text)
    Me.txtConfirm.SelStart = 0
    Me.txtConfirm.SelLength = Len(Me.txtConfirm.Text)
End If
Me.ProgressBar1.Value = KeyQuality
End Sub

Private Sub Check1_Click()
If Me.Check1.Value = 1 Then
    Me.txtCode.PasswordChar = "*"
    Me.txtConfirm.PasswordChar = "*"
    Else
    Me.txtCode.PasswordChar = ""
    Me.txtConfirm.PasswordChar = ""
    End If
End Sub

Private Sub cmdOK_Click()
If Me.txtCode.Text <> Me.txtConfirm.Text Or Me.txtCode.Text = "" Then
    MsgBox "The passphrase and the confirmation do not match." & vbCrLf & "Please enter the passphrase again.", vbCritical, "Ultra"
    Me.txtCode.Text = ""
    Me.txtConfirm.Text = ""
    Me.txtCode.SetFocus
    Exit Sub
    End If
If IsGoodKey(Me.txtCode.Text) = False Then
    MsgBox "The passphrase is too small or contains repetitions and did not meet the minimum security requirements. Please enter another passphrase.", vbCritical, "Ultra"
    Me.txtCode.Text = ""
    Me.txtConfirm.Text = ""
    Me.txtCode.SetFocus
    gstrKeyInput = ""
    Exit Sub
    End If
gstrKeyInput = Me.txtCode.Text
Me.txtCode.Text = ""
Me.txtConfirm.Text = ""
Me.Hide
End Sub

Private Sub cmdCancel_Click()
Me.txtCode.Text = ""
Me.txtConfirm.Text = ""
gblnCancelKey = True
Me.Hide
End Sub

Private Sub txtCode_Change()
Me.ProgressBar1.Value = KeyQuality
Me.txtConfirm.Text = ""
If Len(Me.txtCode.Text) > 4 Then
    Me.cmdOK.Enabled = True
    Else
    Me.cmdOK.Enabled = False
    End If
End Sub

Private Sub txtCode_GotFocus()
Me.txtCode.SelStart = 0
Me.txtCode.SelLength = Len(Me.txtCode.Text)
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    If Me.txtCode <> "" Then Me.txtConfirm.SetFocus
    End If
End Sub

Private Sub txtConfirm_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    If Me.txtConfirm <> "" And Me.cmdOK.Enabled = True Then cmdOK_Click
    End If
End Sub

Private Sub txtConfirm_GotFocus()
Me.txtConfirm.SelStart = 0
Me.txtConfirm.SelLength = Len(Me.txtConfirm.Text)
End Sub


Private Function KeyQuality() As Integer
'check quality of the key
Dim LN As Integer
Dim k As Integer
Dim Uc As Boolean
Dim Lc As Boolean
LN = Len(Me.txtCode.Text)
KeyQuality = LN * 3
'check ucases
For k = 1 To Len(Me.txtCode.Text)
    If Asc(Mid(Me.txtCode.Text, k, 1)) > 64 And Asc(Mid(Me.txtCode.Text, k, 1)) < 91 Then Uc = True
Next k
'check ucases and lcases
For k = 1 To Len(Me.txtCode.Text)
    If Asc(Mid(Me.txtCode.Text, k, 1)) > 96 And Asc(Mid(Me.txtCode.Text, k, 1)) < 123 Then Lc = True
Next k
If Uc = True And Lc = True Then KeyQuality = KeyQuality * 1.2
'check numbers
For k = 1 To Len(Me.txtCode.Text)
    If Asc(Mid(Me.txtCode.Text, k, 1)) > 47 And Asc(Mid(Me.txtCode.Text, k, 1)) < 58 Then
        If Uc = True Or Lc = True Then KeyQuality = KeyQuality * 1.4
        Exit For
        End If
Next k
'check signs
For k = 1 To Len(Me.txtCode.Text)
    If Asc(Mid(Me.txtCode.Text, k, 1)) < 48 Or Asc(Mid(Me.txtCode.Text, k, 1)) > 122 Or (Asc(Mid(Me.txtCode.Text, k, 1)) > 57 And Asc(Mid(Me.txtCode.Text, k, 1)) < 65) Then KeyQuality = KeyQuality * 1.5: Exit For
Next k
If KeyQuality > 100 Then KeyQuality = 100
End Function
