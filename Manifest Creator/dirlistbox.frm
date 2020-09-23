VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Directory"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4575
   Icon            =   "dirlistbox.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   4575
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Directory"
      Height          =   4215
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4335
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   3840
         Width           =   4095
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   3495
         Left            =   120
         ScaleHeight     =   3495
         ScaleWidth      =   4095
         TabIndex        =   3
         Top             =   240
         Width           =   4095
         Begin VB.DriveListBox Drive1 
            Height          =   315
            Left            =   0
            TabIndex        =   5
            Top             =   0
            Width           =   4095
         End
         Begin VB.DirListBox Dir1 
            Height          =   3465
            Left            =   0
            TabIndex        =   4
            Top             =   360
            Width           =   4095
         End
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   4440
      Width           =   1215
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Text2.Text = Text1.Text & "\"
Me.Hide
End Sub

Private Sub Command2_Click()
Me.Hide
End Sub

Private Sub Dir1_Change()
Text1.Text = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error GoTo errhandler
Dir1.Path = Drive1.Drive
errhandler:
If Err.Number <> 0 Then
MsgBox "Could not find drive.", vbCritical
End If
End Sub

Private Sub Form_Initialize()
Text1.Text = Dir1.Path
End Sub

