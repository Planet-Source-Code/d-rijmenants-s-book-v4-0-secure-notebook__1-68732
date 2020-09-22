VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "S-Book Help"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8265
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   8265
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtHelp 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   8250
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim tmp As String
On Error Resume Next
'open the file
FileNr = FreeFile
Open App.Path & "/Readme.txt" For Input As #FileNr
tmp = Input(LOF(FileNr), FileNr)
Close #FileNr
Me.txtHelp.Text = tmp
End Sub

