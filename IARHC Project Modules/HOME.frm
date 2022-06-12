VERSION 5.00
Begin VB.Form HOME_FORM 
   BackColor       =   &H8000000D&
   ClientHeight    =   10935
   ClientLeft      =   480
   ClientTop       =   645
   ClientWidth     =   20250
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   Picture         =   "HOME.frx":0000
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.Timer system_date_time_timer 
      Interval        =   100
      Left            =   17400
      Top             =   240
   End
   Begin VB.CommandButton cmd_orders 
      Caption         =   " ORDERS"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   12600
      MaskColor       =   &H000000FF&
      Picture         =   "HOME.frx":943E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7440
      Width           =   3015
   End
   Begin VB.CommandButton cmd_search_records 
      Caption         =   "SEARCH RECORDS"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   8760
      MaskColor       =   &H000000FF&
      Picture         =   "HOME.frx":11EC7
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7440
      Width           =   3015
   End
   Begin VB.CommandButton cmd_logout 
      BackColor       =   &H008080FF&
      Caption         =   "Logout"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   18000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton cmd_stock_report 
      Caption         =   "STOCK REPORT"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   12360
      Picture         =   "HOME.frx":19EF7
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9240
      Width           =   3495
   End
   Begin VB.CommandButton cmd_issue_renew_return_book 
      Cancel          =   -1  'True
      Caption         =   "ISSUE / RENEW /  RETURN BOOK"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   8520
      Picture         =   "HOME.frx":23E81
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9240
      Width           =   3495
   End
   Begin VB.CommandButton cmd_update_details 
      Caption         =   "UPDATE DETAILS "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   4680
      Picture         =   "HOME.frx":2BA24
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9240
      Width           =   3495
   End
   Begin VB.CommandButton cmd_register 
      Caption         =   "REGISTER"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   4920
      MaskColor       =   &H000000FF&
      Picture         =   "HOME.frx":33616
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7440
      Width           =   3015
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
      Left            =   18120
      TabIndex        =   9
      Top             =   1080
      Width           =   1815
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
      Left            =   18120
      TabIndex        =   8
      Top             =   1560
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
      ForeColor       =   &H0000FF00&
      Height          =   2295
      Left            =   4200
      TabIndex        =   7
      Top             =   720
      Width           =   12015
   End
   Begin VB.Image home_bg 
      Height          =   11295
      Left            =   0
      Picture         =   "HOME.frx":3CA54
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20295
   End
End
Attribute VB_Name = "HOME_FORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_issue_renew_return_book_Click()
ISS_REN_RET_FORM.Show
Unload Me
End Sub

Private Sub cmd_logout_Click()
LOGIN_FORM.Show
Unload Me
End Sub

Private Sub cmd_orders_Click()
ORDERS_FORM.Show
Unload Me
End Sub

Private Sub cmd_register_Click()
register_form.Show
Unload Me
End Sub

Private Sub cmd_search_records_Click()
SEARCH_RECORDS_FORM.Show
Unload Me
End Sub

Private Sub cmd_stock_report_Click()
STOCK_REPORT_FORM.Show
Unload Me
End Sub

Private Sub cmd_update_details_Click()
UPD_DET_FORM.Show
Unload Me
End Sub

Private Sub Form_Load()
home_bg.Move 0, 0, Me.Width, Me.Height
End Sub

Private Sub Form_Resize()
home_bg.Width = Me.ScaleWidth
home_bg.Height = Me.ScaleHeight
End Sub

Private Sub system_date_time_timer_Timer()
system_date_lbl.Caption = Date
system_time_lbl.Caption = Time
End Sub
