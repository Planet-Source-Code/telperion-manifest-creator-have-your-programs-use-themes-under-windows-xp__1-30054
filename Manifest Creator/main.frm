VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manifest Creator"
   ClientHeight    =   4740
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "&Usage Notes"
      Height          =   375
      Left            =   2280
      TabIndex        =   18
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   120
      TabIndex        =   17
      Text            =   "Temp"
      Top             =   4320
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "Windows Executable Files (*.exe)|*.exe"
   End
   Begin VB.Frame Frame2 
      Caption         =   "Other Options"
      Height          =   2055
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   4935
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   1695
         Left            =   120
         ScaleHeight     =   1695
         ScaleWidth      =   4695
         TabIndex        =   10
         Top             =   240
         Width           =   4695
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   0
            TabIndex        =   15
            Top             =   960
            Width           =   3495
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   0
            TabIndex        =   14
            Top             =   240
            Width           =   3495
         End
         Begin VB.CheckBox Check1 
            Caption         =   "&Create as Hidden"
            Height          =   255
            Left            =   0
            TabIndex        =   11
            Top             =   1440
            Value           =   1  'Checked
            Width           =   1575
         End
         Begin VB.Label Label5 
            Caption         =   "Note that the controls are on a picturebox on a frame, not on the frame itself."
            Height          =   975
            Left            =   1200
            TabIndex        =   19
            Top             =   600
            Visible         =   0   'False
            Width           =   3375
         End
         Begin VB.Label Label4 
            Caption         =   "Company Name:"
            Height          =   255
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Project Name:"
            Height          =   255
            Left            =   0
            TabIndex        =   12
            Top             =   720
            Width           =   1095
         End
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   3615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Path Options"
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4935
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1575
         Left            =   120
         ScaleHeight     =   1575
         ScaleWidth      =   4695
         TabIndex        =   2
         Top             =   240
         Width           =   4695
         Begin VB.CheckBox Check2 
            Caption         =   "Same as EXE Path"
            Height          =   255
            Left            =   0
            TabIndex        =   16
            Top             =   1320
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Browse"
            Height          =   255
            Left            =   3720
            TabIndex        =   8
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   0
            TabIndex        =   7
            Top             =   960
            Width           =   3615
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Browse"
            Height          =   255
            Left            =   3720
            TabIndex        =   4
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Destination Path:"
            Height          =   255
            Left            =   0
            TabIndex        =   6
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "EXE Name"
            Height          =   255
            Left            =   0
            TabIndex        =   5
            Top             =   0
            Width           =   975
         End
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Create Manifest"
      Default         =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Menu FileMenu 
      Caption         =   "&File"
      Begin VB.Menu CreateManifestItem 
         Caption         =   "&Create Manifest"
      End
      Begin VB.Menu Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu ExitItem 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu HelpMenu 
      Caption         =   "&Help"
      Begin VB.Menu UsageNotesItem 
         Caption         =   "&Usage Notes"
      End
      Begin VB.Menu Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu AboutItem 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub AboutItem_Click()
MsgBox "Manifest Creator v1.0 created by Telperion (www.echelonstudio.com).  If you like this program, vote for me at Planet Source Code!", vbInformation
End Sub

Private Sub Command1_Click()
CreateManifest

End Sub

Private Sub Command2_Click()
On Error GoTo errhandler
CommonDialog1.ShowOpen 'Show the open dialog
Text1.Text = StripPath(CommonDialog1.FileName) 'Strip the filename from the full path

If Check2.Value = 1 Then 'If user wants "Same as EXE Path"
lookingfor = Text1.Text
Text5.Text = CommonDialog1.FileName 'Load filename into temp text box
lookingin = Text5.Text
mypos = InStr(lookingin, lookingfor) 'Find where the filename starts
Text5.SelStart = mypos - 1
Text5.SelLength = Len(Text1.Text) 'Highlight it in the temp box
Text5.SelText = "" 'Delete it
Text2.Text = Text5.Text 'Move the result (only path) to the visible Destination Path text box
End If

errhandler:
Exit Sub 'If user pressed cancel
End Sub

Private Sub Command3_Click()
Form3.Show 1
End Sub

Private Sub Command4_Click()
Form2.Show 1
End Sub

Private Sub CreateManifestItem_Click()
CreateManifest
End Sub

Private Sub ExitItem_Click()
Unload Me
End Sub

Private Sub Form_Initialize()
'For XP Controls
Dim XP As Long
XP = InitCommonControls
End Sub

Private Sub UsageNotesItem_Click()
Form2.Show 1
End Sub
