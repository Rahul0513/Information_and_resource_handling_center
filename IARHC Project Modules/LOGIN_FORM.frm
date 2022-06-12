VERSION 5.00
Begin VB.Form LOGIN_FORM 
   ClientHeight    =   11280
   ClientLeft      =   -75
   ClientTop       =   315
   ClientWidth     =   20250
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   11280
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.Timer system_date_time_timer 
      Interval        =   100
      Left            =   17640
      Top             =   360
   End
   Begin VB.Frame exit_frame 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   4575
      Left            =   6720
      TabIndex        =   13
      Top             =   3480
      Visible         =   0   'False
      Width           =   7215
      Begin VB.CommandButton cmd_leave 
         BackColor       =   &H000000C0&
         Caption         =   "LEAVE"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3600
         Width           =   2175
      End
      Begin VB.CommandButton cmd_cancel 
         BackColor       =   &H0000FF00&
         Caption         =   "CANCEL"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   3600
         Width           =   2175
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Hope you have a great day ahead :)"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   1335
         Left            =   240
         TabIndex        =   15
         Top             =   1680
         Width           =   6495
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Thankyou for using this application :)"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   1335
         Left            =   240
         TabIndex        =   14
         Top             =   480
         Width           =   6495
      End
   End
   Begin VB.Frame login_Frame 
      BackColor       =   &H00FFFFC0&
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   4935
      Left            =   6600
      TabIndex        =   1
      Top             =   3240
      Width           =   7455
      Begin VB.CommandButton cmd_exit 
         BackColor       =   &H000000C0&
         Caption         =   "EXIT"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   4200
         Width           =   2175
      End
      Begin VB.CommandButton cmd_hide 
         BackColor       =   &H00FFFF80&
         Caption         =   "HIDE"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2520
         Width           =   1335
      End
      Begin VB.CommandButton cmd_see 
         BackColor       =   &H00FFFF80&
         Caption         =   "SEE"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox login_password_txt 
         BackColor       =   &H00FFFF80&
         DataSource      =   "loginado"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   735
         IMEMode         =   3  'DISABLE
         Left            =   3360
         PasswordChar    =   "-"
         TabIndex        =   3
         Top             =   2160
         Width           =   2535
      End
      Begin VB.TextBox login_user_name_txt 
         BackColor       =   &H00FFFF80&
         DataSource      =   "loginado"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   735
         Left            =   3360
         TabIndex        =   2
         Top             =   960
         Width           =   2535
      End
      Begin VB.CommandButton cmd_proceed 
         BackColor       =   &H0000FF00&
         Caption         =   "PROCEED"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   4200
         Width           =   2175
      End
      Begin VB.Label error_msg 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ERROR : INVALID USER NAME OR PASSWORD :("
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   480
         TabIndex        =   11
         Top             =   3360
         Visible         =   0   'False
         Width           =   6735
      End
      Begin VB.Label login_password_lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "PASSWORD :"
         BeginProperty Font 
            Name            =   "Forte"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   615
         Left            =   600
         TabIndex        =   10
         Top             =   2280
         Width           =   2655
      End
      Begin VB.Label login_user_name_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "USER NAME :"
         BeginProperty Font 
            Name            =   "Forte"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   615
         Left            =   720
         TabIndex        =   0
         Top             =   1080
         Width           =   2535
      End
   End
   Begin VB.Label system_time_lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   18240
      TabIndex        =   17
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label system_date_lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   18240
      TabIndex        =   16
      Top             =   120
      Width           =   1815
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
      ForeColor       =   &H00FFFFFF&
      Height          =   2295
      Left            =   4200
      TabIndex        =   12
      Top             =   120
      Width           =   12015
   End
   Begin VB.Image login_bg 
      Height          =   11715
      Left            =   0
      Picture         =   "LOGIN_FORM.frx":0000
      Stretch         =   -1  'True
      Top             =   -600
      Width           =   20730
   End
End
Attribute VB_Name = "LOGIN_FORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub cmd_cancel_Click()
exit_frame.Visible = False
login_Frame.Visible = True
End Sub

Private Sub cmd_exit_Click()
login_Frame.Visible = False
exit_frame.Visible = True
End Sub

Private Sub cmd_hide_Click()
login_password_txt.PasswordChar = "-"
End Sub

Private Sub cmd_leave_Click()
End
End Sub

Private Sub cmd_proceed_Click()
rs.Open "select * from login_table where User_Name = '" + login_user_name_txt.Text + "' and Password = '" + login_password_txt.Text + "'", con, adOpenDynamic, adLockPessimistic
If Not rs.EOF Then
 HOME_FORM.Show
 Unload Me
 Else
 login_user_name_txt.Text = ""
 login_password_txt.Text = ""
 login_user_name_txt.SetFocus
 error_msg.Visible = True
End If
rs.Close
End Sub

Private Sub cmd_see_Click()
login_password_txt.PasswordChar = ""
End Sub

Private Sub Form_Load()
login_bg.Move 0, 0, Me.Width, Me.Height
REPEAT:
On Error GoTo ERR_MSG
con.Open "Provider=Microsoft.Jet.OlEDB.4.0;Data Source =  C:\visual_basic_project\ihprodb.mdb; Persist security Info=False"
Exit Sub
ERR_MSG:
 con.Close
 GoTo REPEAT
End Sub

Private Sub Form_Resize()
login_bg.Width = Me.ScaleWidth
login_bg.Height = Me.ScaleHeight
End Sub

Private Sub system_date_time_timer_Timer()
system_date_lbl.Caption = Date
system_time_lbl.Caption = Time
End Sub

