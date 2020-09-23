VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Usage Notes"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5880
   Icon            =   "usage.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   5880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Notes on XP Controls"
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   4335
         Left            =   120
         ScaleHeight     =   4335
         ScaleWidth      =   5415
         TabIndex        =   1
         Top             =   240
         Width           =   5415
         Begin VB.TextBox Text1 
            BackColor       =   &H8000000F&
            Height          =   1095
            Left            =   0
            MultiLine       =   -1  'True
            TabIndex        =   2
            Text            =   "usage.frx":038A
            Top             =   840
            Width           =   5415
         End
         Begin VB.Label Label4 
            Caption         =   $"usage.frx":0427
            Height          =   615
            Left            =   0
            TabIndex        =   6
            Top             =   3720
            Width           =   5415
         End
         Begin VB.Label Label3 
            Caption         =   $"usage.frx":04F7
            Height          =   855
            Left            =   0
            TabIndex        =   5
            Top             =   2760
            Width           =   5415
         End
         Begin VB.Label Label1 
            Caption         =   $"usage.frx":060A
            Height          =   855
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   5415
         End
         Begin VB.Label Label2 
            Caption         =   $"usage.frx":06FA
            Height          =   615
            Left            =   0
            TabIndex        =   3
            Top             =   2040
            Width           =   5415
         End
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

