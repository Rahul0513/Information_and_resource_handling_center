VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form SPLASH_FORM 
   BorderStyle     =   0  'None
   ClientHeight    =   11520
   ClientLeft      =   -1500
   ClientTop       =   -570
   ClientWidth     =   20490
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   11520
   ScaleWidth      =   20490
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   720
      TabIndex        =   4
      Top             =   8640
      Width           =   19095
      _ExtentX        =   33681
      _ExtentY        =   873
      _Version        =   327682
      Appearance      =   1
      Max             =   105
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   240
      Top             =   9840
   End
   Begin VB.Label percentage_lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   855
      Left            =   10200
      TabIndex        =   3
      Top             =   7680
      Width           =   1455
   End
   Begin VB.Label loading_lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   975
      Left            =   6360
      TabIndex        =   2
      Top             =   7560
      Width           =   3495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Version 0.0.0.1"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   735
      Left            =   16440
      TabIndex        =   1
      Top             =   10800
      Width           =   3855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "INFORMATION AND RESOURCE HANDLING CENTER"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   2175
      Left            =   3600
      TabIndex        =   0
      Top             =   480
      Width           =   13335
   End
   Begin VB.Image splash_bg 
      Height          =   11520
      Left            =   -120
      Picture         =   "SPLASH_FORM.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20685
   End
End
Attribute VB_Name = "splash_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
splash_bg.Move 0, 0, Me.Width, Me.Height
Timer1.Enabled = True
End Sub

Private Sub Form_Resize()
splash_bg.Width = Me.ScaleWidth
splash_bg.Height = Me.ScaleHeight
End Sub

Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 5
loading_lbl.Caption = "Loading..."
percentage_lbl.Caption = ProgressBar1.Value & "%"
If (ProgressBar1.Value = ProgressBar1.Max) Then
 Timer1.Enabled = False
 LOGIN_FORM.Show
 Unload Me
End If
End Sub
