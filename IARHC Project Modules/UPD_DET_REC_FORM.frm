VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form UPD_DET_FORM 
   ClientHeight    =   10935
   ClientLeft      =   1440
   ClientTop       =   315
   ClientWidth     =   20250
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.Frame update_password_frame 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Update User Name and Password"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   4695
      Left            =   7200
      TabIndex        =   75
      Top             =   3720
      Visible         =   0   'False
      Width           =   7695
      Begin VB.CommandButton cmd_refresh 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Refresh"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   3840
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton cmd_upass_proceed 
         BackColor       =   &H0080FF80&
         Caption         =   "Proceed"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox upass_confirm_pass_txt 
         BackColor       =   &H00FFFF80&
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
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   2880
         PasswordChar    =   "-"
         TabIndex        =   81
         Top             =   2640
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.CommandButton cmd_upass_see 
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
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton cmd_upass_hide 
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
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton cmd_upass_update 
         BackColor       =   &H0080FF80&
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   3840
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox upass_current_pass_txt 
         BackColor       =   &H00FFFF80&
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
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   2760
         PasswordChar    =   "-"
         TabIndex        =   76
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox upass_new_pass_txt 
         BackColor       =   &H00FFFF80&
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
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   2880
         PasswordChar    =   "-"
         TabIndex        =   80
         Top             =   2040
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.TextBox upass_new_user_name_txt 
         BackColor       =   &H00FFFF80&
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
         Height          =   495
         Left            =   2880
         TabIndex        =   82
         Top             =   3240
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Image upass_updated_image_pic 
         Height          =   1635
         Left            =   5880
         Picture         =   "UPD_DET_REC_FORM.frx":0000
         Stretch         =   -1  'True
         Top             =   2040
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Line Line27 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   240
         X2              =   5640
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line26 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   240
         X2              =   240
         Y1              =   480
         Y2              =   1920
      End
      Begin VB.Line Line25 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   7440
         X2              =   240
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line Line23 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   5640
         X2              =   5640
         Y1              =   480
         Y2              =   1920
      End
      Begin VB.Line Line22 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   240
         X2              =   7440
         Y1              =   4440
         Y2              =   4440
      End
      Begin VB.Line Line21 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   7440
         X2              =   7440
         Y1              =   480
         Y2              =   4440
      End
      Begin VB.Label upass_confirm_pass_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Password :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   360
         TabIndex        =   138
         Top             =   2640
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Line Line20 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   240
         X2              =   240
         Y1              =   1920
         Y2              =   4440
      End
      Begin VB.Line Line19 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   240
         X2              =   7440
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label upass_msg 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   495
         Left            =   360
         TabIndex        =   137
         Top             =   1320
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.Label upass_current_pass_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Current Password :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   360
         TabIndex        =   136
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label upass_new_pass_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " New Password :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   600
         TabIndex        =   135
         Top             =   2040
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label upass_new_user_name_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " New User Name :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   360
         TabIndex        =   134
         Top             =   3240
         Visible         =   0   'False
         Width           =   2295
      End
   End
   Begin VB.Frame update_newspaper_frame 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Update Newspaper"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   6015
      Left            =   6000
      TabIndex        =   48
      Top             =   3720
      Visible         =   0   'False
      Width           =   9975
      Begin VB.ComboBox un_language_Combo 
         BackColor       =   &H00FFFF80&
         Enabled         =   0   'False
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
         Height          =   465
         ItemData        =   "UPD_DET_REC_FORM.frx":CD5C
         Left            =   2760
         List            =   "UPD_DET_REC_FORM.frx":CD75
         TabIndex        =   53
         Text            =   "------"
         Top             =   2760
         Width           =   2895
      End
      Begin VB.TextBox un_remark_txt 
         BackColor       =   &H00FFFF80&
         Enabled         =   0   'False
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
         Height          =   855
         Left            =   6840
         MultiLine       =   -1  'True
         TabIndex        =   57
         Top             =   5040
         Width           =   2895
      End
      Begin VB.ComboBox un_status_Combo 
         BackColor       =   &H00FFFF80&
         Enabled         =   0   'False
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
         Height          =   465
         ItemData        =   "UPD_DET_REC_FORM.frx":CDB2
         Left            =   2400
         List            =   "UPD_DET_REC_FORM.frx":CDBF
         TabIndex        =   56
         Text            =   "------"
         Top             =   5040
         Width           =   2895
      End
      Begin VB.CommandButton cmd_un_search 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton cmd_un_update 
         BackColor       =   &H0080FF80&
         Caption         =   "Update"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   4200
         Width           =   2175
      End
      Begin VB.CommandButton cmd_refresh 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Refresh"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   3600
         Width           =   2175
      End
      Begin VB.CommandButton cmd_browse 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Browse"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   2880
         Width           =   2175
      End
      Begin VB.TextBox un_title_txt 
         BackColor       =   &H00FFFF80&
         Enabled         =   0   'False
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
         Height          =   495
         Left            =   2760
         TabIndex        =   51
         Top             =   1320
         Width           =   4335
      End
      Begin VB.TextBox un_reference_no_txt 
         BackColor       =   &H00FFFF80&
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
         Height          =   495
         Left            =   2760
         TabIndex        =   49
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox un_pages_txt 
         BackColor       =   &H00FFFF80&
         Enabled         =   0   'False
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
         Height          =   495
         Left            =   2760
         TabIndex        =   54
         Top             =   3480
         Width           =   2895
      End
      Begin VB.TextBox un_price_txt 
         BackColor       =   &H00FFFF80&
         Enabled         =   0   'False
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
         Height          =   495
         Left            =   2760
         TabIndex        =   55
         Top             =   4200
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker stacked_on_date_picker 
         Height          =   495
         Index           =   1
         Left            =   2760
         TabIndex        =   52
         Top             =   2040
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   873
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   16777088
         CalendarForeColor=   4210752
         CalendarTitleBackColor=   16744576
         CalendarTitleForeColor=   65535
         CalendarTrailingForeColor=   16777088
         Format          =   113704960
         CurrentDate     =   44123
      End
      Begin VB.Image upload_pic 
         BorderStyle     =   1  'Fixed Single
         Height          =   2430
         Index           =   2
         Left            =   7440
         Picture         =   "UPD_DET_REC_FORM.frx":CDDB
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2445
      End
      Begin VB.Label un_status_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Status :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   960
         TabIndex        =   133
         Top             =   5040
         Width           =   1215
      End
      Begin VB.Line Line18 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   120
         X2              =   9840
         Y1              =   4920
         Y2              =   4920
      End
      Begin VB.Label un_remark_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Remark :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   5520
         TabIndex        =   132
         Top             =   5160
         Width           =   1215
      End
      Begin VB.Label un_stacked_date_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " Stacked On :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   600
         TabIndex        =   121
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Line Line14 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   7560
         X2              =   9840
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Line Line13 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   7320
         X2              =   7320
         Y1              =   360
         Y2              =   4800
      End
      Begin VB.Label un_language_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Language :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   240
         TabIndex        =   120
         Top             =   2760
         Width           =   2295
      End
      Begin VB.Label un_title_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Title :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   1320
         TabIndex        =   119
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label un_reference_no_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Reference Number :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   120
         TabIndex        =   118
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label un_pages_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " Total Pages :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   480
         TabIndex        =   117
         Top             =   3480
         Width           =   2055
      End
      Begin VB.Label un_price_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " Price :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   480
         TabIndex        =   116
         Top             =   4200
         Width           =   2055
      End
      Begin VB.Label un_inr_lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "INR"
         Enabled         =   0   'False
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
         Height          =   495
         Left            =   4920
         TabIndex        =   115
         Top             =   4200
         Width           =   735
      End
   End
   Begin VB.Timer system_date_time_timer 
      Interval        =   100
      Left            =   17040
      Top             =   840
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
      Height          =   615
      Left            =   18000
      Style           =   1  'Graphical
      TabIndex        =   86
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton cmd_back 
      BackColor       =   &H00C0C0FF&
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   18000
      Style           =   1  'Graphical
      TabIndex        =   85
      Top             =   120
      Width           =   1935
   End
   Begin VB.Frame update_details_frame 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Update Details"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   2055
      Left            =   3480
      TabIndex        =   0
      Top             =   120
      Width           =   11535
      Begin VB.CommandButton cmd_update_pass 
         BackColor       =   &H0080FFFF&
         Caption         =   "Update Username and Password"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   7800
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   480
         Width           =   3255
      End
      Begin VB.CommandButton cmd_update_member 
         BackColor       =   &H0080FFFF&
         Caption         =   "Update Member"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   480
         Width           =   3255
      End
      Begin VB.CommandButton cmd_update_book 
         BackColor       =   &H0080FFFF&
         Caption         =   "Update Book"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1200
         Width           =   3255
      End
      Begin VB.CommandButton cmd_update_newspaper 
         BackColor       =   &H0080FFFF&
         Caption         =   "Update Newspaper"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   480
         Width           =   3255
      End
      Begin VB.CommandButton cmd_update_magazine 
         BackColor       =   &H0080FFFF&
         Caption         =   "Update Magazine"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1200
         Width           =   3255
      End
   End
   Begin MSComDlg.CommonDialog CDC 
      Left            =   17040
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame update_magazine_frame 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Update Magazine"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   6135
      Left            =   6000
      TabIndex        =   61
      Top             =   3840
      Visible         =   0   'False
      Width           =   9975
      Begin VB.ComboBox umag_language_Combo 
         BackColor       =   &H00FFFF80&
         Enabled         =   0   'False
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
         Height          =   465
         ItemData        =   "UPD_DET_REC_FORM.frx":E717
         Left            =   2760
         List            =   "UPD_DET_REC_FORM.frx":E730
         TabIndex        =   67
         Text            =   "------"
         Top             =   3000
         Width           =   2895
      End
      Begin VB.TextBox umag_remark_txt 
         BackColor       =   &H00FFFF80&
         Enabled         =   0   'False
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
         Height          =   855
         Left            =   6840
         MultiLine       =   -1  'True
         TabIndex        =   71
         Top             =   5040
         Width           =   2895
      End
      Begin VB.ComboBox umag_status_Combo 
         BackColor       =   &H00FFFF80&
         Enabled         =   0   'False
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
         Height          =   465
         ItemData        =   "UPD_DET_REC_FORM.frx":E76D
         Left            =   2400
         List            =   "UPD_DET_REC_FORM.frx":E77A
         TabIndex        =   70
         Text            =   "------"
         Top             =   5040
         Width           =   2895
      End
      Begin VB.CommandButton cmd_umag_search 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox umag_price_txt 
         BackColor       =   &H00FFFF80&
         Enabled         =   0   'False
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
         Height          =   495
         Left            =   2760
         TabIndex        =   69
         Top             =   4200
         Width           =   2175
      End
      Begin VB.TextBox umag_total_pages_txt 
         BackColor       =   &H00FFFF80&
         Enabled         =   0   'False
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
         Height          =   495
         Left            =   2760
         TabIndex        =   68
         Top             =   3600
         Width           =   2895
      End
      Begin VB.TextBox umag_reference_no_txt 
         BackColor       =   &H00FFFF80&
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
         Height          =   495
         Left            =   2760
         TabIndex        =   62
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox umag_title_txt 
         BackColor       =   &H00FFFF80&
         Enabled         =   0   'False
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
         Height          =   495
         Left            =   2760
         TabIndex        =   65
         Top             =   1800
         Width           =   4335
      End
      Begin VB.CommandButton cmd_browse 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Browse"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   2880
         Width           =   2175
      End
      Begin VB.CommandButton cmd_refresh 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Refresh"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   3600
         Width           =   2175
      End
      Begin VB.CommandButton cmd_umag_update 
         BackColor       =   &H0080FF80&
         Caption         =   "Update"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   4200
         Width           =   2175
      End
      Begin VB.TextBox umag_issn_code_txt 
         BackColor       =   &H00FFFF80&
         Enabled         =   0   'False
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
         Height          =   495
         Left            =   2760
         TabIndex        =   64
         Top             =   1200
         Width           =   2895
      End
      Begin MSComCtl2.DTPicker stacked_on_date_picker 
         Height          =   495
         Index           =   2
         Left            =   2760
         TabIndex        =   66
         Top             =   2400
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   873
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   16777088
         CalendarForeColor=   4210752
         CalendarTitleBackColor=   16744576
         CalendarTitleForeColor=   65535
         CalendarTrailingForeColor=   16777088
         Format          =   113704960
         CurrentDate     =   44123
      End
      Begin VB.Image upload_pic 
         BorderStyle     =   1  'Fixed Single
         Height          =   2430
         Index           =   3
         Left            =   7440
         Picture         =   "UPD_DET_REC_FORM.frx":E796
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2445
      End
      Begin VB.Line Line17 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   120
         X2              =   9840
         Y1              =   4920
         Y2              =   4920
      End
      Begin VB.Label umag_remark_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Remark :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   5520
         TabIndex        =   131
         Top             =   5280
         Width           =   1215
      End
      Begin VB.Label umag_status_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Status :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   960
         TabIndex        =   130
         Top             =   5040
         Width           =   1215
      End
      Begin VB.Label umag_inr_lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "INR"
         Enabled         =   0   'False
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
         Height          =   495
         Left            =   4920
         TabIndex        =   129
         Top             =   4200
         Width           =   735
      End
      Begin VB.Label umag_price_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " Price :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   480
         TabIndex        =   128
         Top             =   4200
         Width           =   2055
      End
      Begin VB.Label umag_total_pages_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " Total Pages :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   480
         TabIndex        =   127
         Top             =   3600
         Width           =   2055
      End
      Begin VB.Label umag_reference_no_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Reference Number :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   120
         TabIndex        =   126
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label umag_title_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Title :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   1320
         TabIndex        =   125
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label umag_language_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Language :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   240
         TabIndex        =   124
         Top             =   3000
         Width           =   2295
      End
      Begin VB.Line Line16 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   7320
         X2              =   7320
         Y1              =   360
         Y2              =   4800
      End
      Begin VB.Line Line15 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   7560
         X2              =   9840
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Label umag_stacked_on_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " Stacked On :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   600
         TabIndex        =   123
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label um_issn_code_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ISSN Code:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   120
         TabIndex        =   122
         Top             =   1200
         Width           =   2415
      End
   End
   Begin VB.Frame update_member_frame 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Update Member"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   5415
      Left            =   3600
      TabIndex        =   6
      Top             =   3720
      Visible         =   0   'False
      Width           =   13695
      Begin VB.CommandButton cmd_um_paid_fine 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Pay Fine"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11280
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   4800
         Width           =   2175
      End
      Begin VB.TextBox um_fine_amt_txt 
         BackColor       =   &H00FFFF80&
         Enabled         =   0   'False
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
         Height          =   495
         Left            =   7800
         TabIndex        =   21
         Top             =   3840
         Width           =   1215
      End
      Begin VB.CommandButton cmd_suspend_mem 
         BackColor       =   &H00FF8080&
         Caption         =   "Suspend"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9360
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton cmd_activate_mem 
         BackColor       =   &H0000FF00&
         Caption         =   "Re-Activate Member"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   360
         Width           =   3375
      End
      Begin VB.TextBox um_remark_txt 
         BackColor       =   &H00FFFF80&
         Enabled         =   0   'False
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
         Left            =   7800
         MultiLine       =   -1  'True
         TabIndex        =   23
         Top             =   4440
         Width           =   2895
      End
      Begin VB.CommandButton cmd_um_terminate_mem 
         BackColor       =   &H000000FF&
         Caption         =   "Terminate Member"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   960
         Width           =   3375
      End
      Begin VB.TextBox um_status_txt 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   7800
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1920
         Width           =   2895
      End
      Begin VB.CommandButton cmd_um_search 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox um_name_txt 
         BackColor       =   &H00FFFF80&
         Enabled         =   0   'False
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
         Height          =   495
         Left            =   1920
         TabIndex        =   15
         Top             =   2880
         Width           =   2895
      End
      Begin VB.CommandButton cmd_browse 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Browse"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   11280
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   2880
         Width           =   2175
      End
      Begin VB.TextBox um_reg_no_txt 
         BackColor       =   &H00FFFF80&
         Enabled         =   0   'False
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
         Height          =   495
         Left            =   7800
         TabIndex        =   20
         Top             =   3240
         Width           =   2895
      End
      Begin VB.TextBox um_card_no_txt 
         BackColor       =   &H00FFFF80&
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
         Height          =   495
         Left            =   2640
         TabIndex        =   7
         Top             =   600
         Width           =   2895
      End
      Begin VB.ComboBox um_gender_Combo 
         BackColor       =   &H00FFFF80&
         Enabled         =   0   'False
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
         Height          =   465
         ItemData        =   "UPD_DET_REC_FORM.frx":100D2
         Left            =   1920
         List            =   "UPD_DET_REC_FORM.frx":100DC
         TabIndex        =   17
         Text            =   "->Select gender<-"
         Top             =   4080
         Width           =   2895
      End
      Begin VB.ComboBox um_course_Combo 
         BackColor       =   &H00FFFF80&
         Enabled         =   0   'False
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
         Height          =   465
         ItemData        =   "UPD_DET_REC_FORM.frx":100EE
         Left            =   1920
         List            =   "UPD_DET_REC_FORM.frx":10101
         TabIndex        =   18
         Text            =   "------"
         Top             =   4680
         Width           =   2895
      End
      Begin VB.CommandButton cmd_refresh 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Refresh"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   11280
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   3600
         Width           =   2175
      End
      Begin VB.CommandButton cmd_um_update 
         BackColor       =   &H0080FF80&
         Caption         =   "Update"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11280
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   4200
         Width           =   2175
      End
      Begin VB.ComboBox um_type_Combo 
         BackColor       =   &H00FFFF80&
         Enabled         =   0   'False
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
         Height          =   465
         ItemData        =   "UPD_DET_REC_FORM.frx":10122
         Left            =   1920
         List            =   "UPD_DET_REC_FORM.frx":1012C
         TabIndex        =   16
         Text            =   "->Select type<-"
         Top             =   3480
         Width           =   2895
      End
      Begin VB.TextBox um_phone_no_txt 
         BackColor       =   &H00FFFF80&
         Enabled         =   0   'False
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
         Height          =   495
         Left            =   7800
         TabIndex        =   19
         Top             =   2640
         Width           =   2895
      End
      Begin MSComCtl2.DTPicker um_card_issue_date_picker 
         Height          =   495
         Left            =   2640
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1200
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   16777088
         CalendarForeColor=   4210752
         CalendarTitleBackColor=   16744576
         CalendarTitleForeColor=   65535
         CalendarTrailingForeColor=   16777088
         Format          =   113704961
         CurrentDate     =   44123
      End
      Begin MSComCtl2.DTPicker um_card_vailidity_date_picker 
         Height          =   495
         Left            =   2640
         TabIndex        =   10
         Top             =   1800
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   16777088
         CalendarForeColor=   4210752
         CalendarTitleBackColor=   16744576
         CalendarTitleForeColor=   65535
         CalendarTrailingForeColor=   16777088
         Format          =   113704961
         CurrentDate     =   44123
      End
      Begin VB.Label um_fine_amt_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fine Amount :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   5520
         TabIndex        =   139
         Top             =   3840
         Width           =   2055
      End
      Begin VB.Label sm_inr_bal_lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "INR Balance"
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
         Height          =   495
         Left            =   9000
         TabIndex        =   22
         Top             =   3840
         Width           =   1695
      End
      Begin VB.Image upload_pic 
         BorderStyle     =   1  'Fixed Single
         Height          =   2430
         Index           =   0
         Left            =   11160
         Picture         =   "UPD_DET_REC_FORM.frx":10140
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2445
      End
      Begin VB.Label um_remark_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Remark :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   6360
         TabIndex        =   98
         Top             =   4560
         Width           =   1215
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   5880
         X2              =   10920
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   5760
         X2              =   5760
         Y1              =   360
         Y2              =   2400
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   5040
         X2              =   5040
         Y1              =   2760
         Y2              =   5280
      End
      Begin VB.Label um_status_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Status :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   6360
         TabIndex        =   97
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label UMF_details_lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   "Details"
         BeginProperty Font 
            Name            =   "Forte"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   495
         Left            =   600
         TabIndex        =   96
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label um_name_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Name :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   480
         TabIndex        =   95
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label um_type_lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Type :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   480
         TabIndex        =   94
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label um_card_number_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Card Number :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   240
         TabIndex        =   93
         Top             =   600
         Width           =   2175
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   120
         X2              =   10920
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   11040
         X2              =   11040
         Y1              =   360
         Y2              =   5280
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   11160
         X2              =   13560
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Label um_gender_lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Gender :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   480
         TabIndex        =   92
         Top             =   4080
         Width           =   1215
      End
      Begin VB.Label um_reg_no_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Register Number :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   5280
         TabIndex        =   91
         Top             =   3240
         Width           =   2295
      End
      Begin VB.Label um_course_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Course :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   480
         TabIndex        =   90
         Top             =   4680
         Width           =   1215
      End
      Begin VB.Label um_card_issue_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " Issued On :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   360
         TabIndex        =   89
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label um_card_valid_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " Valid Till :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   600
         TabIndex        =   88
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label um_phone_no_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Phone Number :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   5400
         TabIndex        =   87
         Top             =   2640
         Width           =   2175
      End
   End
   Begin VB.Frame update_book_frame 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Update Book"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   5655
      Left            =   3600
      TabIndex        =   28
      Top             =   3720
      Visible         =   0   'False
      Width           =   14175
      Begin VB.ComboBox ub_language_Combo 
         BackColor       =   &H00FFFF80&
         Enabled         =   0   'False
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
         Height          =   465
         ItemData        =   "UPD_DET_REC_FORM.frx":11027
         Left            =   2640
         List            =   "UPD_DET_REC_FORM.frx":11040
         TabIndex        =   38
         Text            =   "------"
         Top             =   5040
         Width           =   2895
      End
      Begin VB.ComboBox ub_status_combo 
         BackColor       =   &H00FFFF80&
         Enabled         =   0   'False
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
         Height          =   465
         ItemData        =   "UPD_DET_REC_FORM.frx":1107D
         Left            =   2760
         List            =   "UPD_DET_REC_FORM.frx":1108A
         TabIndex        =   31
         Text            =   "------"
         Top             =   1080
         Width           =   2895
      End
      Begin VB.CommandButton cmd_ub_search 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11760
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   3600
         Width           =   2175
      End
      Begin VB.TextBox ub_remark_txt 
         BackColor       =   &H00FFFF80&
         Enabled         =   0   'False
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
         Height          =   975
         Left            =   8160
         MultiLine       =   -1  'True
         TabIndex        =   32
         Top             =   480
         Width           =   2895
      End
      Begin VB.CommandButton cmd_ub_update 
         BackColor       =   &H0080FF80&
         Caption         =   "Update"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11760
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   5040
         Width           =   2175
      End
      Begin VB.CommandButton cmd_refresh 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Refresh"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   11760
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   4320
         Width           =   2175
      End
      Begin VB.TextBox ub_author_name_txt 
         BackColor       =   &H00FFFF80&
         Enabled         =   0   'False
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
         Height          =   495
         Left            =   2640
         TabIndex        =   35
         Top             =   3240
         Width           =   2895
      End
      Begin VB.TextBox ub_isbn_code_txt 
         BackColor       =   &H00FFFF80&
         Enabled         =   0   'False
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
         Height          =   495
         Left            =   2640
         TabIndex        =   34
         Top             =   2640
         Width           =   2895
      End
      Begin VB.CommandButton cmd_browse 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Browse"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   11760
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   2880
         Width           =   2175
      End
      Begin VB.TextBox ub_title_txt 
         BackColor       =   &H00FFFF80&
         Enabled         =   0   'False
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
         Height          =   495
         Left            =   2640
         TabIndex        =   33
         Top             =   2040
         Width           =   7815
      End
      Begin VB.TextBox ub_reference_no_txt 
         BackColor       =   &H00FFFF80&
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
         Height          =   495
         Left            =   2760
         TabIndex        =   29
         Top             =   480
         Width           =   2895
      End
      Begin VB.TextBox ub_publisher_name_txt 
         BackColor       =   &H00FFFF80&
         Enabled         =   0   'False
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
         Height          =   495
         Left            =   2640
         TabIndex        =   36
         Top             =   3840
         Width           =   2895
      End
      Begin VB.ComboBox ub_course_Combo 
         BackColor       =   &H00FFFF80&
         Enabled         =   0   'False
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
         Height          =   465
         ItemData        =   "UPD_DET_REC_FORM.frx":110A6
         Left            =   8400
         List            =   "UPD_DET_REC_FORM.frx":110CB
         TabIndex        =   40
         Text            =   "------"
         Top             =   3240
         Width           =   2895
      End
      Begin VB.ComboBox ub_sem_Combo 
         BackColor       =   &H00FFFF80&
         Enabled         =   0   'False
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
         Height          =   465
         ItemData        =   "UPD_DET_REC_FORM.frx":1114B
         Left            =   10080
         List            =   "UPD_DET_REC_FORM.frx":11170
         TabIndex        =   42
         Text            =   "----"
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox ub_edition_txt 
         BackColor       =   &H00FFFF80&
         Enabled         =   0   'False
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
         Height          =   495
         Left            =   2640
         TabIndex        =   37
         Top             =   4440
         Width           =   2895
      End
      Begin VB.TextBox ub_pages_txt 
         BackColor       =   &H00FFFF80&
         Enabled         =   0   'False
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
         Height          =   495
         Left            =   8400
         TabIndex        =   43
         Top             =   4440
         Width           =   2895
      End
      Begin VB.TextBox ub_price_txt 
         BackColor       =   &H00FFFF80&
         Enabled         =   0   'False
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
         Height          =   495
         Left            =   8400
         TabIndex        =   44
         Top             =   5040
         Width           =   2175
      End
      Begin VB.ComboBox ub_year_Combo 
         BackColor       =   &H00FFFF80&
         Enabled         =   0   'False
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
         Height          =   465
         ItemData        =   "UPD_DET_REC_FORM.frx":111B6
         Left            =   8400
         List            =   "UPD_DET_REC_FORM.frx":111CC
         TabIndex        =   41
         Text            =   "----"
         Top             =   3840
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker stacked_on_date_picker 
         Height          =   495
         Index           =   0
         Left            =   8400
         TabIndex        =   39
         Top             =   2640
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   16777088
         CalendarForeColor=   4210752
         CalendarTitleBackColor=   16744576
         CalendarTitleForeColor=   65535
         CalendarTrailingForeColor=   16777088
         Format          =   113704961
         CurrentDate     =   44123
      End
      Begin VB.Image upload_pic 
         BorderStyle     =   1  'Fixed Single
         Height          =   2430
         Index           =   1
         Left            =   11640
         Picture         =   "UPD_DET_REC_FORM.frx":111F3
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2445
      End
      Begin VB.Line Line12 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   6360
         X2              =   6360
         Y1              =   360
         Y2              =   1680
      End
      Begin VB.Label UBF_details_lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   "Details"
         BeginProperty Font 
            Name            =   "Forte"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   495
         Left            =   360
         TabIndex        =   114
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label ub_remark_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Remark :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   6720
         TabIndex        =   113
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label ub_status_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Status :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   1320
         TabIndex        =   112
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label ub_stacked_date_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " Stacked On :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   5880
         TabIndex        =   111
         Top             =   2640
         Width           =   2295
      End
      Begin VB.Label ub_isbn_code_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ISBN Code :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   240
         TabIndex        =   110
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Line Line11 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   11640
         X2              =   14040
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   11520
         X2              =   11520
         Y1              =   360
         Y2              =   5520
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   120
         X2              =   11400
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Label ub_author_name_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Author Name :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   120
         TabIndex        =   109
         Top             =   3240
         Width           =   2295
      End
      Begin VB.Label ub_title_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Title :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   1200
         TabIndex        =   108
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label ub_reference_no_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Reference Number :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   120
         TabIndex        =   107
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label ub_publisher_name_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " Publisher :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   120
         TabIndex        =   106
         Top             =   3840
         Width           =   2295
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   5880
         X2              =   5880
         Y1              =   2640
         Y2              =   5520
      End
      Begin VB.Label ub_course_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Course :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   6960
         TabIndex        =   105
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label ub_year_sem_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Year / Semester:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   6000
         TabIndex        =   104
         Top             =   3840
         Width           =   2175
      End
      Begin VB.Label ub_edition_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " Edition :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   120
         TabIndex        =   103
         Top             =   4440
         Width           =   2295
      End
      Begin VB.Label ub_pages_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " Total Pages :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   6120
         TabIndex        =   102
         Top             =   4440
         Width           =   2055
      End
      Begin VB.Label ub_price_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " Price :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   6120
         TabIndex        =   101
         Top             =   5040
         Width           =   2055
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   9960
         X2              =   9840
         Y1              =   3960
         Y2              =   4200
      End
      Begin VB.Label ub_inr_lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "INR"
         Enabled         =   0   'False
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
         Height          =   495
         Left            =   10560
         TabIndex        =   100
         Top             =   5040
         Width           =   735
      End
      Begin VB.Label ub_language_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Language :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   360
         TabIndex        =   99
         Top             =   5040
         Width           =   2055
      End
   End
   Begin VB.Label msg_lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "d"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   1680
      TabIndex        =   142
      Top             =   2280
      Width           =   15255
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
      TabIndex        =   141
      Top             =   2040
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
      Left            =   18120
      TabIndex        =   140
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Image upd_det_rec_bg 
      Height          =   11175
      Left            =   -120
      Picture         =   "UPD_DET_REC_FORM.frx":12B2F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20730
   End
End
Attribute VB_Name = "UPD_DET_FORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim con1 As New ADODB.Connection
Dim con2 As New ADODB.Connection
Dim con3 As New ADODB.Connection
Dim con4 As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
Dim rs4 As New ADODB.Recordset
Dim flag, flag1, flag2, flag3, flag4, pass, pass1 As Boolean
Dim fine_amt As Long
Dim str, tempCardStatus As String

Private Sub cmd_activate_mem_Click()
um_status_txt.Text = "Active"
End Sub

Private Sub cmd_browse_Click(Index As Integer)
CDC.Filter = " JPG (*.jpg) | *.jpg | JPEG (*.jpeg) | *jpeg | All Files (*.*) | *.*"
CDC.ShowOpen
If CDC.FileName <> "" Then
 str = CDC.FileName
 upload_pic(Index).Picture = LoadPicture(CDC.FileName)
End If
End Sub

Private Sub cmd_refresh_Click(Index As Integer)
If Index = 4 Then
 Call refresh_fun
 Else
  Call refresh_fun
  upload_pic(Index).Picture = Nothing
  str = ""
End If
End Sub

Private Sub cmd_suspend_mem_Click()
um_status_txt.Text = "Suspended"
End Sub

Public Sub reload_data()
If (flag) Then
 rs.Close
 rs.Open "select * from member_table", con, adOpenDynamic, adLockPessimistic
 flag = 0
End If
If (flag1) Then
 rs1.Close
 rs1.Open "select * from book_table", con1, adOpenDynamic, adLockPessimistic
 flag1 = 0
End If
If (flag2) Then
 rs2.Close
 rs2.Open "select * from magazine_table", con2, adOpenDynamic, adLockPessimistic
 flag2 = 0
End If
If (flag3) Then
 rs3.Close
 rs3.Open "select * from newspaper_table", con3, adOpenDynamic, adLockPessimistic
 flag3 = 0
End If
If (flag4) Then
 rs4.Close
 rs4.Open "select *from login_table", con4, adOpenDynamic, adLockPessimistic
 falg4 = 0
End If
End Sub

Private Sub cmd_ub_search_Click()
If ub_reference_no_txt.Text = "" Then
 MsgBox "Please enter the reference number", vbInformation
 ub_reference_no_txt.SetFocus
 Exit Sub
 Else
REPEAT:
  On Error GoTo ERR_MSG
  rs1.Open "select * from book_table where Reference_Number='" + ub_reference_no_txt.Text + "'", con1, adOpenDynamic, adLockPessimistic
  If Not rs1.EOF Then
   ub_reference_no_txt.Text = rs1!Reference_Number
   ub_status_combo.Text = rs1!Book_Status
   ub_remark_txt.Text = rs1!Remark
   ub_title_txt.Text = rs1!Title
   ub_isbn_code_txt.Text = rs1!ISBN_Code
   ub_author_name_txt.Text = rs1!Author_name
   ub_publisher_name_txt.Text = rs1!Publisher
   ub_edition_txt.Text = rs1!Edition
   ub_language_Combo.Text = rs1!Language
   stacked_on_date_picker(0).Value = rs1!Stacked_On
   ub_course_Combo.Text = rs1!Course
   ub_year_Combo.Text = rs1!Course_Year
   ub_sem_Combo.Text = rs1!Course_Sem
   ub_pages_txt.Text = rs1!Total_Pages
   ub_price_txt.Text = rs1!Price
   upload_pic(1).Picture = LoadPicture(rs1!Book_Photo)
   ub_status_combo.Enabled = True
   ub_remark_txt.Enabled = True
   ub_title_txt.Enabled = True
   ub_isbn_code_txt.Enabled = True
   ub_author_name_txt.Enabled = True
   ub_publisher_name_txt.Enabled = True
   ub_edition_txt.Enabled = True
   ub_language_Combo.Enabled = True
   ub_course_Combo.Enabled = True
   ub_year_Combo.Enabled = True
   ub_sem_Combo.Enabled = True
   ub_pages_txt.Enabled = True
   ub_price_txt.Enabled = True
   cmd_browse(1).Enabled = True
   cmd_ub_update.Enabled = True
  Else
   MsgBox "Record not found... Please check the reference number!!!", vbCritical
 End If
 flag1 = 1
End If
Exit Sub
ERR_MSG:
 rs1.Close
 GoTo REPEAT
End Sub

Private Sub cmd_ub_update_Click()
If (ub_reference_no_txt.Text = "" And ub_title_txt.Text = "" And ub_isbn_code_txt.Text = "" And ub_author_name_txt.Text = "" And ub_publisher_name_txt.Text = "" And ub_edition_txt.Text = "" And ub_language_Combo.Text = "------" And ub_course_Combo.Text = "------" And ub_year_Combo.Text = "----" And ub_sem_Combo.Text = "----" And ub_pages_txt.Text = "" And ub_price_txt.Text = "") Then
   MsgBox "Please enter all the fields", vbInformation
   Exit Sub
  ElseIf ub_reference_no_txt.Text = "" Then
   MsgBox "Please enter reference number", vbInformation
   ub_reference_no_txt.SetFocus
   Exit Sub
  ElseIf ub_title_txt.Text = "" Then
   MsgBox "Please enter Title", vbInformation
   ub_title_txt.SetFocus
   Exit Sub
  ElseIf ub_isbn_code_txt.Text = "" Then
   MsgBox "Please enter ISBN code", vbInformation
   ub_isbn_code_txt.SetFocus
   Exit Sub
  ElseIf ub_author_name_txt.Text = "" Then
   MsgBox "Please enter author name", vbInformation
   ub_author_name_txt.SetFocus
   Exit Sub
  ElseIf ub_publisher_name_txt.Text = "" Then
   MsgBox "Please enter publisher name", vbInformation
   ub_publisher_name_txt.SetFocus
   Exit Sub
  ElseIf ub_edition_txt.Text = "" Then
   MsgBox "Please enter edition", vbInformation
   ub_edition_txt.SetFocus
   Exit Sub
  ElseIf ub_language_Combo.Text = "------" Then
   MsgBox "Please select language", vbInformation
   ub_language_Combo.SetFocus
   Exit Sub
  ElseIf ub_course_Combo.Text = "------" Then
   MsgBox "Please select course", vbInformation
   ub_course_Combo.SetFocus
   Exit Sub
  ElseIf ub_year_Combo.Text = "----" Then
   MsgBox "Please select year", vbInformation
   ub_year_Combo.SetFocus
   Exit Sub
  ElseIf ub_sem_Combo.Text = "----" Then
   MsgBox "Please select sem", vbInformation
   ub_sem_Combo.SetFocus
   Exit Sub
  ElseIf ub_pages_txt.Text = "" Then
   MsgBox "Please enter number of pages", vbInformation
   ub_pages_txt.SetFocus
   Exit Sub
  ElseIf ub_price_txt.Text = "" Then
   MsgBox "Please enter price", vbInformation
   ub_price_txt.SetFocus
   Exit Sub
  Else
   rs1.Fields("Reference_Number").Value = ub_reference_no_txt.Text
   rs1.Fields("Title").Value = ub_title_txt.Text
   rs1.Fields("Stacked_On").Value = stacked_on_date_picker(0).Value
   rs1.Fields("ISBN_Code").Value = ub_isbn_code_txt.Text
   rs1.Fields("Author_Name").Value = ub_author_name_txt.Text
   rs1.Fields("Publisher").Value = ub_publisher_name_txt.Text
   rs1.Fields("Edition").Value = ub_edition_txt.Text
   rs1.Fields("Language").Value = ub_language_Combo.Text
   rs1.Fields("Course").Value = ub_course_Combo.Text
   rs1.Fields("Course_year").Value = ub_year_Combo.Text
   rs1.Fields("Course_Sem").Value = ub_sem_Combo.Text
   rs1.Fields("Total_pages").Value = ub_pages_txt.Text
   rs1.Fields("Price").Value = ub_price_txt.Text
   If str = "" Then
    str = "C:\visual_basic_project\Default_Pics\not_available.jpeg"
   End If
   rs1.Fields("Book_Photo").Value = str
   rs1.Fields("Book_Status").Value = ub_status_combo.Text
   rs1.Fields("Remark").Value = ub_remark_txt.Text
   rs1.Update
   MsgBox "Process completed successfully", vbInformation
   Call cmd_refresh_Click(1)
   ub_reference_no_txt.SetFocus
   disable
   flag1 = 1
End If
End Sub

Private Sub cmd_um_paid_fine_Click()
fine_amt = Val(InputBox("Please enter the amount", "Pay fine amount", vbOK))
 If Val(fine_amt) <= Val(um_fine_amt_txt) Then
  um_fine_amt_txt.Text = Val(um_fine_amt_txt) - Val(fine_amt)
 Else
  MsgBox "Please enter the correct amount...", vbCritical
 End If
End Sub

Private Sub cmd_um_search_Click()
If um_card_no_txt.Text = "" Then
 MsgBox "Please enter the Card number...", vbInformation
 um_card_no_txt.SetFocus
 Exit Sub
 Else
REPEAT:
  On Error GoTo ERR_MSG
  rs.Open "select * from member_table where Card_Number= '" + um_card_no_txt.Text + "' ", con, adOpenDynamic, adLockPessimistic
  If Not rs.EOF Then
   um_card_no_txt.Text = rs!Card_Number
   um_name_txt.Text = rs!Name
   um_status_txt.Text = rs!Card_Status
   tempCardStatus = rs!Card_Status
   um_type_Combo.Text = rs!Type
   um_gender_Combo.Text = rs!Gender
   um_course_Combo.Text = rs!Course
   um_phone_no_txt.Text = rs!Phone_Number
   um_reg_no_txt.Text = rs!Register_Number
   um_fine_amt_txt.Text = rs!Fine_Balance
   um_remark_txt.Text = rs!Remark
   um_card_vailidity_date_picker.Value = rs!Card_Valid_Till
   um_card_issue_date_picker.Value = rs!Card_Issued_On
   upload_pic(0).Picture = LoadPicture(rs!Member_Photo)
   um_name_txt.Enabled = True
   um_status_txt.Enabled = True
   um_type_Combo.Enabled = True
   um_gender_Combo.Enabled = True
   um_course_Combo.Enabled = True
   um_phone_no_txt.Enabled = True
   um_reg_no_txt.Enabled = True
   um_fine_amt_txt.Enabled = True
   um_remark_txt.Enabled = True
   um_card_vailidity_date_picker.Enabled = True
   cmd_activate_mem.Enabled = True
   cmd_um_terminate_mem.Enabled = True
   cmd_suspend_mem.Enabled = True
   cmd_browse(0).Enabled = True
   cmd_um_update.Enabled = True
   cmd_um_paid_fine.Enabled = True
   Else
   MsgBox "Record not found... Please check the card number!!!", vbCritical
 End If
 flag = 1
End If
Exit Sub
ERR_MSG:
 rs.Close
 GoTo REPEAT
End Sub

Private Sub cmd_um_terminate_mem_Click()
um_status_txt.Text = "Terminated"
End Sub

Private Sub cmd_um_update_Click()
If (um_name_txt.Text = "" And um_type_Combo.Text = "->Select type<-" And um_gender_Combo.Text = "->Select gender<-" And um_phone_no_txt.Text = "" And um_course_Combo.Text = "------" And um_reg_no_txt.Text = "" And um_card_no_txt.Text = "") Then
   MsgBox "Please enter all the fields", vbInformation
   Exit Sub
 ElseIf um_name_txt.Text = "" Then
   MsgBox "Please enter the name", vbInformation
   um_name_txt.SetFocus
   Exit Sub
 ElseIf um_type_Combo.Text = "->Select type<-" Then
   MsgBox "Please select type", vbInformation
   um_type_Combo.SetFocus
   Exit Sub
 ElseIf um_gender_Combo.Text = "->Select gender<-" Then
   MsgBox "Please select gender", vbInformation
   um_gender_Combo.SetFocus
   Exit Sub
 ElseIf um_phone_no_txt.Text = "" Then
   MsgBox "Please enter phone number", vbInformation
   um_phone_no_txt.SetFocus
   Exit Sub
  ElseIf Len(um_phone_no_txt.Text) <> 10 Then
   MsgBox "Please enter a valid Indian phone number", vbInformation
   um_phone_no_txt.SetFocus
   Exit Sub
 ElseIf um_course_Combo.Text = "------" Then
   MsgBox "Please select course", vbInformation
   um_course_Combo.SetFocus
   Exit Sub
 ElseIf um_reg_no_txt.Text = "" Then
   MsgBox "Please enter the register number", vbInformation
   um_reg_no_txt.SetFocus
   Exit Sub
 ElseIf um_card_no_txt.Text = "" Then
   MsgBox "Please enter the card number", vbInformation
   um_card_no_txt.SetFocus
   Exit Sub
 ElseIf um_card_vailidity_date_picker <= um_card_issue_date_picker.Value Then
   MsgBox "Please check the card validity date", vbInformation
   um_card_vailidity_date_picker.SetFocus
   Exit Sub
 Else
   rs.Fields("Name").Value = um_name_txt.Text
   rs.Fields("Type").Value = um_type_Combo.Text
   rs.Fields("Gender").Value = um_gender_Combo.Text
   rs.Fields("Phone_Number").Value = um_phone_no_txt.Text
   rs.Fields("Course").Value = um_course_Combo.Text
   rs.Fields("Register_Number").Value = um_reg_no_txt.Text
   rs.Fields("Card_Number").Value = um_card_no_txt.Text
   rs.Fields("Card_Issued_On").Value = um_card_issue_date_picker.Value
   rs.Fields("Card_Valid_Till").Value = um_card_vailidity_date_picker.Value
   rs.Fields("Card_Status").Value = um_status_txt.Text
   If tempCardStatus <> rs.Fields("Card_Status").Value Then
    If um_status_txt.Text = "Terminated" Then
     rs.Fields("Number_Of_Times_Terminated").Value = rs.Fields("Number_Of_Times_Terminated").Value + 1
     ElseIf um_status_txt.Text = "Suspended" Then
     rs.Fields("Number_Of_Times_Suspended").Value = rs.Fields("Number_Of_Times_Suspended").Value + 1
    End If
   End If
   If str = "" And um_gender_Combo.Text = "Male" Then
    str = "C:\visual_basic_project\Default_Pics\male.jpeg"
    ElseIf str = "" And um_gender_Combo.Text = "Female" Then
     str = "C:\visual_basic_project\Default_Pics\female.jpeg"
    ElseIf str = "" Then
     str = "C:\visual_basic_project\Default_Pics\not_available.jpeg"
   End If
   rs.Fields("Member_Photo").Value = str
   rs.Fields("Total_Fine_Paid").Value = rs.Fields("Total_Fine_Paid") + Val(fine_amt)
   rs.Fields("Fine_Balance").Value = um_fine_amt_txt.Text
   rs.Update
   MsgBox "Process completed successfully", vbInformation
   Call cmd_refresh_Click(0)
   um_card_no_txt.SetFocus
   disable
   flag = 1
End If
End Sub

Private Sub cmd_umag_search_Click()
If umag_reference_no_txt.Text = "" Then
  MsgBox "Please enter the reference number", vbInformation
  umag_reference_no_txt.SetFocus
  Exit Sub
  Else
REPEAT:
   On Error GoTo ERR_MSG
   rs3.Open "select * from magazine_table where Reference_Number='" + umag_reference_no_txt.Text + "'", con3, adOpenDynamic, adLockPessimistic
   If Not rs3.EOF Then
    umag_reference_no_txt.Text = rs3!Reference_Number
    umag_issn_code_txt.Text = rs3!ISSN_Code
    umag_title_txt.Text = rs3!Title
    stacked_on_date_picker(2).Value = rs3!Stacked_On
    umag_language_Combo.Text = rs3!Language
    umag_total_pages_txt.Text = rs3!Total_Pages
    umag_price_txt.Text = rs3!Price
    umag_status_Combo.Text = rs3!Magazine_Status
    umag_remark_txt.Text = rs3!Remark
    upload_pic(3).Picture = LoadPicture(rs3!Magazine_Photo)
    umag_issn_code_txt.Enabled = True
    umag_title_txt.Enabled = True
    umag_language_Combo.Enabled = True
    umag_total_pages_txt.Enabled = True
    umag_price_txt.Enabled = True
    umag_status_Combo.Enabled = True
    umag_remark_txt.Enabled = True
    cmd_browse(3).Enabled = True
    cmd_umag_update.Enabled = True
   Else
   MsgBox "Record not found... Please check the reference number!!!", vbCritical
 End If
 flag3 = 1
End If
Exit Sub
ERR_MSG:
 rs3.Close
 GoTo REPEAT
End Sub

Private Sub cmd_umag_update_Click()
If (umag_issn_code_txt.Text = "" And umag_reference_no_txt.Text = "" And umag_title_txt.Text = "" And umag_language_Combo.Text = "------" And umag_total_pages_txt.Text = "" And umag_price_txt.Text = "") Then
   MsgBox "Please enter all the fields", vbInformation
   Exit Sub
  ElseIf umag_issn_code_txt.Text = "" Then
   MsgBox "Please enter ISSN code", vbInformation
   umag_issn_code_txt.SetFocus
   Exit Sub
  ElseIf umag_reference_no_txt.Text = "" Then
   MsgBox "Please enter reference number", vbInformation
   umag_reference_no_txt.SetFocus
   Exit Sub
  ElseIf umag_title_txt.Text = "" Then
   MsgBox "Please enter title", vbInformation
   umag_title_txt.SetFocus
   Exit Sub
  ElseIf umag_language_Combo.Text = "------" Then
   MsgBox "Please select language", vbInformation
   umag_language_Combo.SetFocus
   Exit Sub
  ElseIf umag_total_pages_txt.Text = "" Then
   MsgBox "Please enter total number pages", vbInformation
   umag_total_pages_txt.SetFocus
   Exit Sub
  ElseIf umag_price_txt.Text = "" Then
   MsgBox "Please enter price", vbInformation
   umag_price_txt.SetFocus
   Exit Sub
  Else
   rs3.Fields("Reference_Number").Value = umag_reference_no_txt.Text
   rs3.Fields("ISSN_Code").Value = umag_issn_code_txt.Text
   rs3.Fields("Title").Value = umag_title_txt.Text
   rs3.Fields("Stacked_On").Value = stacked_on_date_picker(2).Value
   rs3.Fields("Language").Value = umag_language_Combo.Text
   rs3.Fields("Total_Pages").Value = umag_total_pages_txt.Text
   rs3.Fields("Price").Value = umag_price_txt.Text
   If str = "" Then
    str = "C:\visual_basic_project\Default_Pics\not_available.jpeg"
   End If
   rs3.Fields("Magazine_Photo").Value = str
   rs3.Fields("Magazine_Status") = umag_status_Combo.Text
   rs3.Fields("Remark") = umag_remark_txt.Text
   rs3.Update
   MsgBox "Process completed successfully", vbInformation
   Call cmd_refresh_Click(3)
   umag_reference_no_txt.SetFocus
   disable
   flag = 1
End If
End Sub

Private Sub cmd_un_search_Click()
If un_reference_no_txt.Text = "" Then
  MsgBox "Please enter the reference number", vbInformation
  un_reference_no_txt.SetFocus
  Exit Sub
  Else
REPEAT:
   On Error GoTo ERR_MSG
   rs2.Open "select * from newspaper_table where Reference_Number='" + un_reference_no_txt.Text + "'", con2, adOpenDynamic, adLockPessimistic
   If Not rs2.EOF Then
    un_reference_no_txt.Text = rs2!Reference_Number
    un_title_txt.Text = rs2!Title
    stacked_on_date_picker(1).Value = rs2!Stacked_On
    un_language_Combo.Text = rs2!Language
    un_pages_txt.Text = rs2!Total_Pages
    un_price_txt.Text = rs2!Price
    un_status_Combo.Text = rs2!Newspaper_Status
    un_remark_txt.Text = rs2!Remark
    upload_pic(2).Picture = LoadPicture(rs2!Newspaper_Photo)
    un_title_txt.Enabled = True
    un_language_Combo.Enabled = True
    un_pages_txt.Enabled = True
    un_price_txt.Enabled = True
    un_status_Combo.Enabled = True
    un_remark_txt.Enabled = True
    cmd_browse(2).Enabled = True
    cmd_un_update.Enabled = True
   Else
   MsgBox "Record not found... Please check the reference number!!!", vbCritical
 End If
 flag2 = 1
End If
Exit Sub
ERR_MSG:
 rs2.Close
 GoTo REPEAT
End Sub

Private Sub cmd_un_update_Click()
If (un_reference_no_txt.Text = "" And un_title_txt.Text = "" And un_language_Combo.Text = "------" And un_pages_txt.Text = "" And un_price_txt.Text = "") Then
   MsgBox "Please enter all the fields", vbInformation
   Exit Sub
  ElseIf un_reference_no_txt.Text = "" Then
   MsgBox "Please enter reference number", vbInformation
   un_reference_no_txt.SetFocus
   Exit Sub
  ElseIf un_title_txt.Text = "" Then
   MsgBox "Please enter title", vbInformation
   un_title_txt.SetFocus
   Exit Sub
  ElseIf un_language_Combo.Text = "------" Then
   MsgBox "Please select language", vbInformation
   un_language_Combo.SetFocus
   Exit Sub
  ElseIf un_pages_txt.Text = "" Then
   MsgBox "Please enter number of pages", vbInformation
   un_pages_txt.SetFocus
   Exit Sub
  ElseIf un_price_txt.Text = "" Then
   MsgBox "Please enter price", vbInformation
   un_price_txt.SetFocus
   Exit Sub
  Else
   rs2.Fields("Reference_Number").Value = un_reference_no_txt.Text
   rs2.Fields("Title").Value = un_title_txt.Text
   rs2.Fields("Stacked_On").Value = stacked_on_date_picker(1).Value
   rs2.Fields("Language").Value = un_language_Combo.Text
   rs2.Fields("Total_Pages").Value = un_pages_txt.Text
   rs2.Fields("Price").Value = un_price_txt.Text
   If str = "" Then
    str = "C:\visual_basic_project\Default_Pics\not_available.jpeg"
   End If
   rs2.Fields("Newspaper_Photo").Value = str
   rs2.Fields("Newspaper_Status").Value = un_status_Combo.Text
   rs2.Fields("Remark").Value = un_remark_txt.Text
   rs2.Update
   MsgBox "Process completed successfully", vbInformation
   Call cmd_refresh_Click(2)
   un_reference_no_txt.SetFocus
   disable
   flag2 = 1
End If
End Sub

Private Sub cmd_upass_hide_Click()
upass_current_pass_txt.PasswordChar = "-"
upass_new_pass_txt.PasswordChar = "-"
upass_confirm_pass_txt.PasswordChar = "-"
End Sub

Private Sub cmd_upass_proceed_Click()
If upass_current_pass_txt.Text = "" Then
 MsgBox "Please enter the current password", vbInformation
 upass_current_pass_txt.SetFocus
 Exit Sub
 Else
REPEAT:
  On Error GoTo ERR_MSG
  rs4.Open "select * from login_table where Password='" + upass_current_pass_txt.Text + "'", con4, adOpenDynamic, adLockPessimistic
  If Not rs4.EOF Then
   upass_new_pass_lbl.Visible = True
   upass_new_pass_txt.Visible = True
   upass_confirm_pass_lbl.Visible = True
   upass_confirm_pass_txt.Visible = True
   upass_new_user_name_lbl.Visible = True
   upass_new_user_name_txt.Visible = True
   cmd_refresh(4).Visible = True
   cmd_upass_update.Visible = True
   upass_current_pass_txt.Locked = True
   Else
    MsgBox "Please enter the right password", vbCritical
    upass_current_pass_txt.SetFocus
  End If
  flag4 = 1
End If
Exit Sub
ERR_MSG:
 rs4.Close
 GoTo REPEAT:
End Sub

Private Sub cmd_upass_see_Click()
upass_current_pass_txt.PasswordChar = ""
upass_new_pass_txt.PasswordChar = ""
upass_confirm_pass_txt.PasswordChar = ""
End Sub

Private Sub cmd_upass_update_Click()
If upass_new_pass_txt.Text = "" And upass_confirm_pass_txt.Text = "" And upass_new_user_name_txt.Text = "" Then
 MsgBox "Please enter the fields", vbInformation
 upass_new_pass_txt.SetFocus
 Exit Sub
 ElseIf upass_new_pass_txt.Text = "" Then
  MsgBox "Please enter the New password", vbInformation
  upass_new_pass_txt.SetFocus
  Exit Sub
 ElseIf upass_confirm_pass_txt.Text = "" Then
  MsgBox "Please re-enter the password for confirmation...", vbInformation
  upass_confirm_pass_txt.SetFocus
  Exit Sub
 ElseIf upass_new_user_name_txt.Text = "" Then
  MsgBox "Please enter the New User name", vbInformation
  upass_new_user_name_txt.SetFocus
  Exit Sub
 ElseIf upass_confirm_pass_txt.Text <> upass_new_pass_txt.Text Then
  MsgBox "New password and confirmed password does not match", vbCritical
  upass_new_user_name_txt.SetFocus
  Exit Sub
 Else
  If rs4.Fields("Password").Value = upass_new_pass_txt.Text Then
   MsgBox "New password cannot be same as old password", vbCritical
   upass_new_pass_txt.SetFocus
   Exit Sub
   Else
    pass = 1
  End If
  If rs4.Fields("User_Name").Value = upass_new_user_name_txt.Text Then
   MsgBox "New User Name is same as the old one... please change", vbCritical
   upass_new_user_name_txt.SetFocus
   Exit Sub
   Else
    pass1 = 1
  End If
  If pass And pass1 Then
   rs4.Fields("Password").Value = upass_new_pass_txt.Text
   rs4.Fields("User_Name").Value = upass_new_user_name_txt.Text
   rs4.Update
   upass_updated_image_pic.Visible = True
   MsgBox "Password Updated...", vbInformation
   MsgBox "Redirecting to login page...", vbOKOnly + vbInformation
   flag4 = 1
   Call cmd_logout_Click
  End If
End If
End Sub

Private Sub cmd_update_book_Click()
Call refresh_fun
msg_lbl.Visible = True
update_member_frame.Visible = False
update_book_frame.Visible = True
update_newspaper_frame.Visible = False
update_magazine_frame.Visible = False
update_password_frame.Visible = False
If (flag) Then
 con.Close
 flag = 0
End If
If (flag2) Then
 con2.Close
 flag2 = 0
End If
If (flag3) Then
 con3.Close
 flag3 = 0
End If
If (flag4) Then
 con4.Close
 flag4 = 0
End If
REPEAT:
 On Error GoTo ERR_MSG
 con1.Open "Provider=Microsoft.Jet.OlEDB.4.0;Data Source =  C:\visual_basic_project\ihprodb.mdb; Persist security Info=False"
 flag1 = 1
 Exit Sub
ERR_MSG:
 con1.Close
 GoTo REPEAT
End Sub

Private Sub cmd_update_magazine_Click()
Call refresh_fun
msg_lbl.Visible = True
update_member_frame.Visible = False
update_book_frame.Visible = False
update_newspaper_frame.Visible = False
update_magazine_frame.Visible = True
update_password_frame.Visible = False
If (flag1) Then
 con1.Close
 flag1 = 0
End If
If (flag2) Then
 con2.Close
 flag2 = 0
End If
If (flag) Then
 con.Close
 flag = 0
End If
If (flag4) Then
 con4.Close
 flag4 = 0
End If
REPEAT:
 On Error GoTo ERR_MSG
 con3.Open "Provider=Microsoft.Jet.OlEDB.4.0;Data Source =  C:\visual_basic_project\ihprodb.mdb; Persist security Info=False"
 flag3 = 1
 Exit Sub
ERR_MSG:
 con3.Close
 GoTo REPEAT:
End Sub

Private Sub cmd_update_member_Click()
Call refresh_fun
msg_lbl.Visible = True
update_member_frame.Visible = True
update_book_frame.Visible = False
update_newspaper_frame.Visible = False
update_magazine_frame.Visible = False
update_password_frame.Visible = False
If (flag1) Then
 con1.Close
 flag1 = 0
End If
If (flag2) Then
 con2.Close
 flag2 = 0
End If
If (flag3) Then
 con3.Close
 flag3 = 0
End If
If (flag4) Then
 con4.Close
 flag4 = 0
End If
REPEAT:
 On Error GoTo ERR_MSG
 con.Open "Provider=Microsoft.Jet.OlEDB.4.0;Data Source =  C:\visual_basic_project\ihprodb.mdb; Persist security Info=False"
 flag = 1
 Exit Sub
ERR_MSG:
 con.Close
 GoTo REPEAT
End Sub

Private Sub cmd_update_newspaper_Click()
Call refresh_fun
msg_lbl.Visible = True
update_member_frame.Visible = False
update_book_frame.Visible = False
update_newspaper_frame.Visible = True
update_magazine_frame.Visible = False
update_password_frame.Visible = False
If (flag1) Then
 con1.Close
 flag1 = 0
End If
If (flag2) Then
 con2.Close
 flag2 = 0
End If
If (flag3) Then
 con3.Close
 flag3 = 0
End If
If (flag4) Then
 con4.Close
 flag4 = 0
End If
REPEAT:
 On Error GoTo ERR_MSG
 con2.Open "Provider=Microsoft.Jet.OlEDB.4.0;Data Source =  C:\visual_basic_project\ihprodb.mdb; Persist security Info=False"
 flag2 = 1
 Exit Sub
ERR_MSG:
  con2.Close
  GoTo REPEAT
End Sub
Private Function refresh_fun()
um_card_no_txt.Text = ""
um_name_txt.Text = ""
um_status_txt.Text = ""
um_type_Combo.Text = "->Select type<-"
um_gender_Combo.Text = "->Select gender<-"
um_course_Combo.Text = "------"
um_phone_no_txt.Text = ""
um_reg_no_txt.Text = ""
um_fine_amt_txt.Text = ""
um_remark_txt.Text = ""
um_card_vailidity_date_picker.Value = Now
um_card_issue_date_picker.Value = Now
um_status_txt.BackColor = &HFFFF80
ub_reference_no_txt.Text = ""
ub_status_combo.Text = "------"
ub_remark_txt.Text = ""
ub_title_txt.Text = ""
ub_isbn_code_txt.Text = ""
ub_author_name_txt.Text = ""
ub_publisher_name_txt.Text = ""
ub_edition_txt.Text = ""
ub_language_Combo.Text = "------"
ub_course_Combo.Text = "------"
ub_year_Combo.Text = "----"
ub_sem_Combo.Text = "----"
ub_pages_txt.Text = ""
ub_price_txt.Text = ""
un_reference_no_txt.Text = ""
un_title_txt.Text = ""
un_language_Combo.Text = "------"
un_pages_txt.Text = ""
un_price_txt.Text = ""
un_status_Combo.Text = "------"
un_remark_txt.Text = ""
umag_reference_no_txt.Text = ""
umag_issn_code_txt.Text = ""
umag_title_txt.Text = ""
umag_language_Combo.Text = "------"
umag_total_pages_txt.Text = ""
umag_price_txt.Text = ""
umag_status_Combo.Text = "------"
umag_remark_txt.Text = ""
upass_new_pass_txt.Text = ""
upass_confirm_pass_txt.Text = ""
upass_new_user_name_txt.Text = ""
End Function

Private Sub cmd_update_pass_Click()
Call refresh_fun
msg_lbl.Visible = False
update_member_frame.Visible = False
update_book_frame.Visible = False
update_newspaper_frame.Visible = False
update_magazine_frame.Visible = False
update_password_frame.Visible = True
If (flag) Then
 con.Close
 flag = 0
End If
If (flag2) Then
 con2.Close
 flag2 = 0
End If
If (flag3) Then
 con3.Close
 flag3 = 0
End If
If (flag1) Then
 con1.Close
 flag1 = 0
End If
REPEAT:
 On Error GoTo ERR_MSG
 con4.Open "Provider=Microsoft.Jet.OlEDB.4.0;Data Source =  C:\visual_basic_project\ihprodb.mdb; Persist security Info=False"
 flag4 = 1
 Exit Sub
ERR_MSG:
 con4.Close
 GoTo REPEAT
End Sub

Private Sub Form_Load()
upd_det_rec_bg.Move 0, 0, Me.Width, Me.Height
msg_lbl.Visible = True
msg_lbl.Caption = "Please enter the Card Number / Reference Number and click on search button"
End Sub

Private Sub Form_Resize()
upd_det_rec_bg.Width = Me.ScaleWidth
upd_det_rec_bg.Height = Me.ScaleHeight
End Sub

Private Sub cmd_back_Click()
HOME_FORM.Show
If (flag) Then
 con.Close
 flag = 0
End If
If (flag1) Then
 con1.Close
 flag1 = 0
End If
If (flag2) Then
 con2.Close
 flag2 = 0
End If
If (flag3) Then
 con3.Close
 flag3 = 0
End If
If (flag4) Then
 con4.Close
 flag4 = 0
End If
Unload Me
End Sub

Private Sub cmd_logout_Click()
LOGIN_FORM.Show
If (flag) Then
 con.Close
 flag = 0
End If
If (flag1) Then
 con1.Close
 flag1 = 0
End If
If (flag2) Then
 con2.Close
 flag2 = 0
End If
If (flag3) Then
 con3.Close
 flag3 = 0
End If
If (flag4) Then
 con4.Close
 flag4 = 0
End If
Unload Me
End Sub

Private Sub Option_see_records_Click()
update_details_frame.Visible = False
records_frame.Visible = True
End Sub

Private Sub Option_update_details_Click()
update_details_frame.Visible = True
records_frame.Visible = False
End Sub

Private Sub system_date_time_timer_Timer()
system_date_lbl.Caption = Date
system_time_lbl.Caption = Time
End Sub

Private Sub ub_author_name_txt_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub ub_course_Combo_Change()
If ub_course_Combo.Text = "Reference Book" Then
 ub_year_Combo.Text = "N/A"
 ub_year_Combo.Enabled = False
 ub_sem_Combo.Text = "N/A"
 ub_sem_Combo.Enabled = False
 Else
  ub_year_Combo.Text = "----"
  ub_year_Combo.Enabled = True
  ub_sem_Combo.Text = "----"
  ub_sem_Combo.Enabled = True
End If
End Sub

Private Sub ub_edition_txt_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub ub_isbn_code_txt_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or KeyAscii = 45 Then
 Else
  KeyAscii = 0
  MsgBox "Enter numbers only", vbInformation
End If
End Sub

Private Sub ub_pages_txt_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
 Else
  KeyAscii = 0
  MsgBox "Enter numbers only", vbInformation
End If
End Sub

Private Sub ub_price_txt_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or KeyAscii = 46 Then
 Else
  KeyAscii = 0
  MsgBox "Enter numbers only", vbInformation
End If
End Sub

Private Sub ub_publisher_name_txt_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub ub_reference_no_txt_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub ub_title_txt_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub um_card_no_txt_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub um_fine_amt_txt_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or KeyAscii = 46 Then
 Else
  KeyAscii = 0
  MsgBox "Enter numbers only", vbInformation
End If
End Sub

Private Sub um_name_txt_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub um_phone_no_txt_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
 Else
  KeyAscii = 0
  MsgBox "Enter numbers only", vbInformation
End If
End Sub

Private Sub um_status_txt_Change()
If um_status_txt.Text = "Active" Then
 um_status_txt.BackColor = &HFF00&
 ElseIf um_status_txt.Text = "Terminated" Then
  um_status_txt.BackColor = &HFF&
 ElseIf um_status_txt.Text = "Suspended" Then
  um_status_txt.BackColor = &HFF8080
End If
End Sub

Private Sub umag_price_txt_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or KeyAscii = 46 Then
 Else
  KeyAscii = 0
  MsgBox "Enter numbers only", vbInformation
End If
End Sub

Private Sub umag_reference_no_txt_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub umag_title_txt_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub umag_total_pages_txt_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
 Else
  KeyAscii = 0
  MsgBox "Enter numbers only", vbInformation
End If
End Sub

Private Sub un_pages_txt_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
 Else
  KeyAscii = 0
  MsgBox "Enter numbers only", vbInformation
End If
End Sub

Private Sub un_price_txt_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or KeyAscii = 46 Then
 Else
  KeyAscii = 0
  MsgBox "Enter numbers only", vbInformation
End If
End Sub

Private Sub un_reference_no_txt_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub un_title_txt_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub upass_confirm_pass_txt_Click()
upass_msg.Caption = "Re-Enter the password"
End Sub

Private Sub upass_current_pass_txt_Click()
upass_msg.Visible = True
upass_msg.Caption = "Enter the current password"
End Sub

Private Sub upass_new_pass_txt_Click()
upass_msg.Caption = "Enter the new password"
End Sub

Private Sub upass_new_user_name_txt_Click()
upass_msg.Caption = "Enter the new User Name"
End Sub

Sub disable()
ub_status_combo.Enabled = False
ub_remark_txt.Enabled = False
ub_title_txt.Enabled = False
ub_isbn_code_txt.Enabled = False
ub_author_name_txt.Enabled = False
ub_publisher_name_txt.Enabled = False
ub_edition_txt.Enabled = False
ub_language_Combo.Enabled = False
ub_course_Combo.Enabled = False
ub_year_Combo.Enabled = False
ub_sem_Combo.Enabled = False
ub_pages_txt.Enabled = False
ub_price_txt.Enabled = False
cmd_browse(1).Enabled = False
cmd_ub_update.Enabled = False
um_name_txt.Enabled = False
um_status_txt.Enabled = False
um_type_Combo.Enabled = False
um_gender_Combo.Enabled = False
um_course_Combo.Enabled = False
um_phone_no_txt.Enabled = False
um_reg_no_txt.Enabled = False
um_fine_amt_txt.Enabled = False
um_remark_txt.Enabled = False
um_card_vailidity_date_picker.Enabled = False
cmd_activate_mem.Enabled = False
cmd_um_terminate_mem.Enabled = False
cmd_suspend_mem.Enabled = False
cmd_browse(0).Enabled = False
cmd_um_update.Enabled = False
cmd_um_paid_fine.Enabled = False
umag_issn_code_txt.Enabled = False
umag_title_txt.Enabled = False
umag_language_Combo.Enabled = False
umag_total_pages_txt.Enabled = False
umag_price_txt.Enabled = False
umag_status_Combo.Enabled = False
umag_remark_txt.Enabled = False
cmd_browse(3).Enabled = False
cmd_umag_update.Enabled = False
un_title_txt.Enabled = False
un_language_Combo.Enabled = False
un_pages_txt.Enabled = False
un_price_txt.Enabled = False
un_status_Combo.Enabled = False
un_remark_txt.Enabled = False
cmd_browse(2).Enabled = False
cmd_un_update.Enabled = False
End Sub
