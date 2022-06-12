VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form REGISTER_FORM 
   ClientHeight    =   10935
   ClientLeft      =   1815
   ClientTop       =   -450
   ClientWidth     =   20250
   ControlBox      =   0   'False
   FillColor       =   &H00404040&
   ForeColor       =   &H00404040&
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.Frame add_member_frame 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Add Member"
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
      Left            =   2760
      TabIndex        =   5
      Top             =   3120
      Visible         =   0   'False
      Width           =   13695
      Begin VB.TextBox m_card_issued_on_txt 
         BackColor       =   &H00FFFF80&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "d MMMM yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
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
         Left            =   8280
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   3720
         Width           =   2055
      End
      Begin VB.ComboBox m_course_Combo 
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
         ItemData        =   "REGISTER_FORM.frx":0000
         Left            =   7680
         List            =   "REGISTER_FORM.frx":0022
         TabIndex        =   10
         Text            =   "------"
         Top             =   1440
         Width           =   2895
      End
      Begin VB.TextBox m_card_status_txt 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         DataField       =   "Card_Status"
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
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   "Active"
         Top             =   4440
         Width           =   2895
      End
      Begin VB.TextBox m_phone_no_txt 
         BackColor       =   &H00FFFF80&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
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
         Left            =   7680
         TabIndex        =   9
         Top             =   720
         Width           =   2895
      End
      Begin VB.ComboBox m_type_Combo 
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
         Height          =   465
         ItemData        =   "REGISTER_FORM.frx":0089
         Left            =   1800
         List            =   "REGISTER_FORM.frx":0093
         TabIndex        =   7
         Text            =   "->Select type<-"
         Top             =   1440
         Width           =   2895
      End
      Begin VB.CommandButton cmd_m_add 
         BackColor       =   &H0080FF80&
         Caption         =   "Add"
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
         TabIndex        =   18
         Top             =   4440
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
         Index           =   0
         Left            =   11280
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   3720
         Width           =   2175
      End
      Begin VB.ComboBox m_gender_Combo 
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
         Height          =   465
         ItemData        =   "REGISTER_FORM.frx":00A7
         Left            =   1800
         List            =   "REGISTER_FORM.frx":00B1
         TabIndex        =   8
         Text            =   "->Select gender<-"
         Top             =   2160
         Width           =   2895
      End
      Begin VB.TextBox m_card_no_txt 
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
         Left            =   3000
         TabIndex        =   12
         Top             =   3720
         Width           =   2895
      End
      Begin VB.TextBox m_reg_no_txt 
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
         Left            =   7680
         TabIndex        =   11
         Top             =   2160
         Width           =   2895
      End
      Begin VB.CommandButton cmd_browse 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Browse"
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
         TabIndex        =   16
         Top             =   2880
         Width           =   2175
      End
      Begin VB.TextBox m_name_txt 
         BackColor       =   &H00FFFF80&
         DataField       =   "Name"
         DataSource      =   "mem_ado"
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
         Left            =   1800
         TabIndex        =   6
         Top             =   720
         Width           =   2895
      End
      Begin MSComCtl2.DTPicker m_card_vailidity_date_picker 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "dd-MM-yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
         Height          =   495
         Left            =   8280
         TabIndex        =   15
         Top             =   4440
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         _Version        =   393216
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
         Format          =   104333313
         CurrentDate     =   44123
      End
      Begin VB.Image upload_pic 
         BorderStyle     =   1  'Fixed Single
         DataField       =   "Member_Photo"
         DataSource      =   "mem_ado"
         Height          =   2430
         Index           =   0
         Left            =   11160
         Picture         =   "REGISTER_FORM.frx":00C3
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2445
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   6240
         X2              =   6240
         Y1              =   3360
         Y2              =   5280
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   5040
         X2              =   5040
         Y1              =   360
         Y2              =   3000
      End
      Begin VB.Label m_card_status_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Card Status :"
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
         TabIndex        =   90
         Top             =   4440
         Width           =   1815
      End
      Begin VB.Label m_phone_no_lbl 
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
         Left            =   5160
         TabIndex        =   80
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label AMF_card_details_lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   "Card Details"
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
         Left            =   480
         TabIndex        =   67
         Top             =   3000
         Width           =   2535
      End
      Begin VB.Label m_card_valid_lbl 
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
         Left            =   6480
         TabIndex        =   66
         Top             =   4440
         Width           =   1575
      End
      Begin VB.Label m_card_issue_lbl 
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
         Left            =   6360
         TabIndex        =   65
         Top             =   3720
         Width           =   1695
      End
      Begin VB.Label m_course_lbl 
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
         Left            =   6240
         TabIndex        =   64
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label m_reg_no_lbl 
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
         Left            =   5160
         TabIndex        =   63
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label m_gender_lbl 
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
         Left            =   360
         TabIndex        =   62
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   11160
         X2              =   13560
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   11040
         X2              =   11040
         Y1              =   360
         Y2              =   5160
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   120
         X2              =   10920
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Label m_card_number_lbl 
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
         Left            =   480
         TabIndex        =   61
         Top             =   3720
         Width           =   2295
      End
      Begin VB.Label m_type_lbl 
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
         Left            =   360
         TabIndex        =   60
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label m_name_lbl 
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
         Left            =   360
         TabIndex        =   59
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.Frame add_newspaper_frame 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Add Newspaper"
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
      Height          =   4935
      Left            =   4800
      TabIndex        =   36
      Top             =   3360
      Visible         =   0   'False
      Width           =   9975
      Begin VB.TextBox n_stacked_on_txt 
         BackColor       =   &H00FFFF80&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "d MMMM yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
         DataSource      =   "registerado"
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
         Locked          =   -1  'True
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   2040
         Width           =   2895
      End
      Begin VB.ComboBox n_language_Combo 
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
         Height          =   465
         ItemData        =   "REGISTER_FORM.frx":0FAA
         Left            =   2760
         List            =   "REGISTER_FORM.frx":0FC3
         TabIndex        =   40
         Text            =   "------"
         Top             =   2760
         Width           =   2895
      End
      Begin VB.TextBox n_price_txt 
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
         TabIndex        =   42
         Top             =   4200
         Width           =   2175
      End
      Begin VB.TextBox n_pages_txt 
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
         TabIndex        =   41
         Top             =   3480
         Width           =   2895
      End
      Begin VB.TextBox n_reference_no_txt 
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
         TabIndex        =   37
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox n_title_txt 
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
         TabIndex        =   38
         Top             =   1320
         Width           =   4335
      End
      Begin VB.CommandButton cmd_browse 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Browse"
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
         TabIndex        =   43
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
         Index           =   2
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   3600
         Width           =   2175
      End
      Begin VB.CommandButton cmd_n_add 
         BackColor       =   &H0080FF80&
         Caption         =   "Add"
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
         TabIndex        =   45
         Top             =   4200
         Width           =   2175
      End
      Begin VB.Image upload_pic 
         BorderStyle     =   1  'Fixed Single
         Height          =   2430
         Index           =   2
         Left            =   7440
         Picture         =   "REGISTER_FORM.frx":1000
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2445
      End
      Begin VB.Label n_inr_lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "INR"
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
         TabIndex        =   88
         Top             =   4200
         Width           =   735
      End
      Begin VB.Label n_price_lbl 
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
         TabIndex        =   86
         Top             =   4200
         Width           =   2055
      End
      Begin VB.Label n_pages_lbl 
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
         TabIndex        =   85
         Top             =   3480
         Width           =   2055
      End
      Begin VB.Label n_reference_no_lbl 
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
         TabIndex        =   84
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label n_title_lbl 
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
         TabIndex        =   83
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label n_language_lbl 
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
         TabIndex        =   82
         Top             =   2760
         Width           =   2295
      End
      Begin VB.Line Line12 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   7320
         X2              =   7320
         Y1              =   360
         Y2              =   4800
      End
      Begin VB.Line Line13 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   7440
         X2              =   9840
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Label n_stacked_date_lbl 
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
         TabIndex        =   81
         Top             =   2040
         Width           =   1935
      End
   End
   Begin VB.Timer system_date_time_timer 
      Interval        =   100
      Left            =   17160
      Top             =   720
   End
   Begin MSComDlg.CommonDialog CDC 
      Left            =   17160
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Left            =   17880
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   120
      Width           =   1935
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
      Left            =   17880
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   840
      Width           =   1935
   End
   Begin VB.Frame register_menu 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Register"
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
      Height          =   1335
      Left            =   2640
      TabIndex        =   0
      Top             =   120
      Width           =   12855
      Begin VB.CommandButton cmd_add_magazine 
         BackColor       =   &H0080FFFF&
         Caption         =   "Add Magazine"
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
         Left            =   9600
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   480
         Width           =   3015
      End
      Begin VB.CommandButton cmd_add_newspaper 
         BackColor       =   &H0080FFFF&
         Caption         =   "Add Newspaper"
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
         Left            =   6480
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   480
         Width           =   3015
      End
      Begin VB.CommandButton cmd_add_book 
         BackColor       =   &H0080FFFF&
         Caption         =   "Add Book"
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
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   480
         Width           =   3015
      End
      Begin VB.CommandButton cmd_add_member 
         BackColor       =   &H0080FFFF&
         Caption         =   "Add Member"
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   480
         Width           =   3015
      End
   End
   Begin VB.Frame add_magazine_frame 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Add Magazine"
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
      Height          =   4935
      Left            =   4800
      TabIndex        =   46
      Top             =   3360
      Visible         =   0   'False
      Width           =   9975
      Begin VB.TextBox mag_stacked_on_txt 
         BackColor       =   &H00FFFF80&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "d MMMM yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
         DataSource      =   "registerado"
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
         Locked          =   -1  'True
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   2400
         Width           =   2895
      End
      Begin VB.ComboBox mag_language_Combo 
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
         Height          =   465
         ItemData        =   "REGISTER_FORM.frx":293C
         Left            =   2760
         List            =   "REGISTER_FORM.frx":2955
         TabIndex        =   51
         Text            =   "------"
         Top             =   3000
         Width           =   2895
      End
      Begin VB.TextBox mag_issn_code_txt 
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
         TabIndex        =   47
         Top             =   600
         Width           =   2895
      End
      Begin VB.CommandButton cmd_mag_add 
         BackColor       =   &H0080FF80&
         Caption         =   "Add"
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
         TabIndex        =   56
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
         Index           =   3
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   3600
         Width           =   2175
      End
      Begin VB.CommandButton cmd_browse 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Browse"
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
         TabIndex        =   54
         Top             =   2880
         Width           =   2175
      End
      Begin VB.TextBox mag_title_txt 
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
         Top             =   1800
         Width           =   4335
      End
      Begin VB.TextBox mag_reference_no_txt 
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
         TabIndex        =   48
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox mag_total_pages_txt 
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
         TabIndex        =   52
         Top             =   3600
         Width           =   2895
      End
      Begin VB.TextBox mag_price_txt 
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
         TabIndex        =   53
         Top             =   4200
         Width           =   2175
      End
      Begin VB.Image upload_pic 
         BorderStyle     =   1  'Fixed Single
         Height          =   2430
         Index           =   3
         Left            =   7440
         Picture         =   "REGISTER_FORM.frx":2992
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2445
      End
      Begin VB.Label mag_issn_code_lbl 
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
         TabIndex        =   98
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label mag_stacked_on_lbl 
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
         TabIndex        =   97
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Line Line14 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   7440
         X2              =   9840
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Line Line11 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   7320
         X2              =   7320
         Y1              =   360
         Y2              =   4800
      End
      Begin VB.Label mag_language_lbl 
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
         TabIndex        =   96
         Top             =   3000
         Width           =   2295
      End
      Begin VB.Label mag_title_lbl 
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
         TabIndex        =   95
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label mag_reference_no_lbl 
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
         TabIndex        =   94
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label mag_total_pages_lbl 
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
         TabIndex        =   93
         Top             =   3600
         Width           =   2055
      End
      Begin VB.Label mag_price_lbl 
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
         TabIndex        =   92
         Top             =   4200
         Width           =   2055
      End
      Begin VB.Label mag_inr_lbl 
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
         TabIndex        =   91
         Top             =   4200
         Width           =   735
      End
   End
   Begin VB.Frame add_book_frame 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Add Book"
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
      Height          =   5175
      Left            =   2760
      TabIndex        =   19
      Top             =   3120
      Visible         =   0   'False
      Width           =   14175
      Begin VB.TextBox b_stacked_on_txt 
         BackColor       =   &H00FFFF80&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "dd-MM-yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
         DataSource      =   "registerado"
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
         Left            =   8280
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   600
         Width           =   2055
      End
      Begin VB.ComboBox b_language_Combo 
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
         Height          =   465
         ItemData        =   "REGISTER_FORM.frx":42CE
         Left            =   8400
         List            =   "REGISTER_FORM.frx":42E7
         TabIndex        =   27
         Text            =   "------"
         Top             =   2160
         Width           =   2895
      End
      Begin VB.ComboBox b_year_Combo 
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
         Height          =   465
         ItemData        =   "REGISTER_FORM.frx":4324
         Left            =   8400
         List            =   "REGISTER_FORM.frx":433A
         TabIndex        =   29
         Top             =   3360
         Width           =   1335
      End
      Begin VB.TextBox b_price_txt 
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
         Left            =   8400
         TabIndex        =   32
         Top             =   4560
         Width           =   2175
      End
      Begin VB.TextBox b_pages_txt 
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
         Left            =   8400
         TabIndex        =   31
         Top             =   3960
         Width           =   2895
      End
      Begin VB.TextBox b_edition_txt 
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
         TabIndex        =   26
         Top             =   4320
         Width           =   2895
      End
      Begin VB.ComboBox b_sem_Combo 
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
         Height          =   465
         ItemData        =   "REGISTER_FORM.frx":4361
         Left            =   10080
         List            =   "REGISTER_FORM.frx":4386
         TabIndex        =   30
         Top             =   3360
         Width           =   1215
      End
      Begin VB.ComboBox b_course_Combo 
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
         Height          =   465
         ItemData        =   "REGISTER_FORM.frx":43CC
         Left            =   8400
         List            =   "REGISTER_FORM.frx":43F1
         TabIndex        =   28
         Text            =   "------"
         Top             =   2760
         Width           =   2895
      End
      Begin VB.TextBox b_publisher_name_txt 
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
         TabIndex        =   25
         Top             =   3720
         Width           =   2895
      End
      Begin VB.TextBox b_reference_no_txt 
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
         TabIndex        =   20
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox b_title_txt 
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
         TabIndex        =   21
         Top             =   1200
         Width           =   7815
      End
      Begin VB.CommandButton cmd_browse 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Browse"
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
         TabIndex        =   33
         Top             =   2880
         Width           =   2175
      End
      Begin VB.TextBox b_isbn_code_txt 
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
         TabIndex        =   23
         Top             =   2520
         Width           =   2895
      End
      Begin VB.TextBox b_author_name_txt 
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
         TabIndex        =   24
         Top             =   3120
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
         Index           =   1
         Left            =   11760
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   3720
         Width           =   2175
      End
      Begin VB.CommandButton cmd_b_add 
         BackColor       =   &H0080FF80&
         Caption         =   "Add"
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
         TabIndex        =   35
         Top             =   4440
         Width           =   2175
      End
      Begin VB.Image upload_pic 
         BorderStyle     =   1  'Fixed Single
         Height          =   2430
         Index           =   1
         Left            =   11640
         Picture         =   "REGISTER_FORM.frx":4471
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2445
      End
      Begin VB.Label b_language_lbl 
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
         Left            =   6120
         TabIndex        =   89
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label b_inr_lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "INR"
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
         TabIndex        =   87
         Top             =   4560
         Width           =   735
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   9960
         X2              =   9840
         Y1              =   3480
         Y2              =   3720
      End
      Begin VB.Label b_price_lbl 
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
         TabIndex        =   79
         Top             =   4560
         Width           =   2055
      End
      Begin VB.Label b_pages_lbl 
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
         TabIndex        =   78
         Top             =   3960
         Width           =   2055
      End
      Begin VB.Label b_edition_lbl 
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
         TabIndex        =   77
         Top             =   4320
         Width           =   2295
      End
      Begin VB.Label b_year_sem_lbl 
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
         TabIndex        =   76
         Top             =   3360
         Width           =   2175
      End
      Begin VB.Label b_course_lbl 
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
         TabIndex        =   75
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   5880
         X2              =   5880
         Y1              =   2160
         Y2              =   5040
      End
      Begin VB.Label b_publisher_name_lbl 
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
         TabIndex        =   74
         Top             =   3720
         Width           =   2295
      End
      Begin VB.Label ABF_book_details_lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   "Book Details"
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
         TabIndex        =   73
         Top             =   1800
         Width           =   2655
      End
      Begin VB.Label b_reference_no_lbl 
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
         TabIndex        =   72
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label b_title_lbl 
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
         TabIndex        =   71
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label b_author_name_lbl 
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
         TabIndex        =   70
         Top             =   3120
         Width           =   2295
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   120
         X2              =   11400
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   11520
         X2              =   11520
         Y1              =   360
         Y2              =   5040
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   11640
         X2              =   14040
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Label b_isbn_code_lbl 
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
         TabIndex        =   69
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Label b_stacked_date_lbl 
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
         TabIndex        =   68
         Top             =   600
         Width           =   2295
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
      Left            =   18000
      TabIndex        =   100
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label system_date_lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "."
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "d MMMM yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16393
         SubFormatType   =   3
      EndProperty
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
      Left            =   18000
      TabIndex        =   99
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Image register_bg 
      Height          =   10905
      Left            =   0
      Picture         =   "REGISTER_FORM.frx":5DAD
      Stretch         =   -1  'True
      Top             =   0
      WhatsThisHelpID =   1
      Width           =   20295
   End
End
Attribute VB_Name = "register_form"
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
Dim flag, flag1, flag2, flag3, flag4 As Boolean
Dim str As String

Private Sub b_author_name_txt_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub b_edition_txt_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub b_publisher_name_txt_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub b_reference_no_txt_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub b_title_txt_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cmd_b_add_Click()
If (b_reference_no_txt.Text = "" And b_title_txt.Text = "" And b_isbn_code_txt.Text = "" And b_author_name_txt.Text = "" And b_publisher_name_txt.Text = "" And b_edition_txt.Text = "" And b_language_Combo.Text = "------" And b_course_Combo.Text = "------" And b_year_Combo.Text = "----" And b_sem_Combo.Text = "----" And b_pages_txt.Text = "" And b_price_txt.Text = "") Then
   MsgBox "Please enter all the fields", vbInformation
   Exit Sub
  ElseIf b_reference_no_txt.Text = "" Then
   MsgBox "Please enter reference number", vbInformation
   b_reference_no_txt.SetFocus
   Exit Sub
  ElseIf b_title_txt.Text = "" Then
   MsgBox "Please enter Title", vbInformation
   b_title_txt.SetFocus
   Exit Sub
  ElseIf b_isbn_code_txt.Text = "" Then
   MsgBox "Please enter ISBN code", vbInformation
   b_isbn_code_txt.SetFocus
   Exit Sub
  ElseIf b_author_name_txt.Text = "" Then
   MsgBox "Please enter author name", vbInformation
   b_author_name_txt.SetFocus
   Exit Sub
  ElseIf b_publisher_name_txt.Text = "" Then
   MsgBox "Please enter publisher name", vbInformation
   b_publisher_name_txt.SetFocus
   Exit Sub
  ElseIf b_edition_txt.Text = "" Then
   MsgBox "Please enter edition", vbInformation
   b_edition_txt.SetFocus
   Exit Sub
  ElseIf b_language_Combo.Text = "------" Then
   MsgBox "Please select language", vbInformation
   b_language_Combo.SetFocus
   Exit Sub
  ElseIf b_course_Combo.Text = "------" Then
   MsgBox "Please select course", vbInformation
   b_course_Combo.SetFocus
   Exit Sub
  ElseIf b_year_Combo.Text = "----" Then
   MsgBox "Please select year", vbInformation
   b_year_Combo.SetFocus
   Exit Sub
  ElseIf b_sem_Combo.Text = "----" Then
   MsgBox "Please select sem", vbInformation
   b_sem_Combo.SetFocus
   Exit Sub
  ElseIf b_pages_txt.Text = "" Then
   MsgBox "Please enter number of pages", vbInformation
   b_pages_txt.SetFocus
   Exit Sub
  ElseIf b_price_txt.Text = "" Then
   MsgBox "Please enter price", vbInformation
   b_price_txt.SetFocus
   Exit Sub
  Else
   rs1.Fields("Reference_Number").Value = b_reference_no_txt.Text
   rs1.Fields("Title").Value = b_title_txt.Text
   rs1.Fields("Stacked_On").Value = b_stacked_on_txt.Text
   rs1.Fields("ISBN_Code").Value = b_isbn_code_txt.Text
   rs1.Fields("Author_Name").Value = b_author_name_txt.Text
   rs1.Fields("Publisher").Value = b_publisher_name_txt.Text
   rs1.Fields("Edition").Value = b_edition_txt.Text
   rs1.Fields("Language").Value = b_language_Combo.Text
   rs1.Fields("Course").Value = b_course_Combo.Text
   rs1.Fields("Course_year").Value = b_year_Combo.Text
   rs1.Fields("Course_Sem").Value = b_sem_Combo.Text
   rs1.Fields("Total_pages").Value = b_pages_txt.Text
   rs1.Fields("Price").Value = b_price_txt.Text
   If str = "" Then
    str = "C:\visual_basic_project\Default_Pics\not_available.jpeg"
   End If
   rs1.Fields("Book_Photo").Value = str
   On Error GoTo ERR_MSG
   rs1.AddNew
   MsgBox "Process completed successfully", vbInformation
   b_reference_no_txt.SetFocus
   Call cmd_refresh_Click(1)
End If
Exit Sub
ERR_MSG:
 MsgBox "Book Reference Number : " & b_reference_no_txt.Text & "already exist !!! please change the Reference Number", vbExclamation
End Sub

Private Sub b_price_txt_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or KeyAscii = 46 Then
 Else
  KeyAscii = 0
  MsgBox "Enter numbers only", vbInformation
End If
End Sub

Private Sub b_isbn_code_txt_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or KeyAscii = 45 Then
 Else
  KeyAscii = 0
  MsgBox "Enter numbers only", vbInformation
End If
End Sub

Private Sub b_pages_txt_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
 Else
  KeyAscii = 0
  MsgBox "Enter numbers only", vbInformation
End If
End Sub

Private Sub cmd_m_add_GotFocus()
If m_reg_no_txt.Text = "" Then
 Exit Sub
 ElseIf m_reg_no_txt.Text = "N/A" Then
  Exit Sub
 ElseIf m_reg_no_txt.Text <> "N/A" Then
   If flag4 Then
    con4.Close
    flag4 = 0
   End If
   con4.Open "Provider=Microsoft.Jet.OlEDB.4.0;Data Source =  C:\visual_basic_project\ihprodb.mdb; Persist security Info=False"
   rs4.Open "select * from member_table where Register_Number= '" + m_reg_no_txt.Text + "'", con4, adOpenDynamic, adLockPessimistic
   If Not rs4.EOF Then
    MsgBox "Register Number already exist!!! please check the Register Number", vbExclamation
    flag4 = 1
    Exit Sub
    Else
     flag4 = 1
   End If
End If
End Sub

Private Sub cmd_mag_add_Click()
If (mag_issn_code_txt.Text = "" And mag_reference_no_txt.Text = "" And mag_title_txt.Text = "" And mag_language_Combo.Text = "------" And mag_total_pages_txt.Text = "" And mag_price_txt.Text = "") Then
   MsgBox "Please enter all the fields", vbInformation
   Exit Sub
  ElseIf mag_issn_code_txt.Text = "" Then
   MsgBox "Please enter ISSN code", vbInformation
   mag_issn_code_txt.SetFocus
   Exit Sub
  ElseIf mag_reference_no_txt.Text = "" Then
   MsgBox "Please enter reference number", vbInformation
   mag_reference_no_txt.SetFocus
   Exit Sub
  ElseIf mag_title_txt.Text = "" Then
   MsgBox "Please enter title", vbInformation
   mag_title_txt.SetFocus
   Exit Sub
  ElseIf mag_language_Combo.Text = "------" Then
   MsgBox "Please select language", vbInformation
   mag_language_Combo.SetFocus
   Exit Sub
  ElseIf mag_total_pages_txt.Text = "" Then
   MsgBox "Please enter total number pages", vbInformation
   mag_total_pages_txt.SetFocus
   Exit Sub
  ElseIf mag_price_txt.Text = "" Then
   MsgBox "Please enter price", vbInformation
   mag_price_txt.SetFocus
   Exit Sub
  Else
   rs3.Fields("Reference_Number").Value = mag_reference_no_txt.Text
   rs3.Fields("ISSN_Code").Value = mag_issn_code_txt.Text
   rs3.Fields("Title").Value = mag_title_txt.Text
   rs3.Fields("Stacked_On").Value = mag_stacked_on_txt.Text
   rs3.Fields("Language").Value = mag_language_Combo.Text
   rs3.Fields("Total_Pages").Value = mag_total_pages_txt.Text
   rs3.Fields("Price").Value = mag_price_txt.Text
   If str = "" Then
    str = "C:\visual_basic_project\Default_Pics\not_available.jpeg"
   End If
   rs3.Fields("Magazine_Photo").Value = str
   On Error GoTo ERR_MSG
   rs3.AddNew
   MsgBox "Process completed successfully", vbInformation
   mag_reference_no_txt.SetFocus
   Call cmd_refresh_Click(3)
End If
Exit Sub
ERR_MSG:
 MsgBox "Magazine Reference Number : " & mag_reference_no_txt.Text & "already exist !!! please change the Reference Number", vbExclamation
End Sub

Private Sub cmd_n_add_Click()
If (n_reference_no_txt.Text = "" And n_title_txt.Text = "" And n_language_Combo.Text = "------" And n_pages_txt.Text = "" And n_price_txt.Text = "") Then
   MsgBox "Please enter all the fields", vbInformation
   Exit Sub
  ElseIf n_reference_no_txt.Text = "" Then
   MsgBox "Please enter reference number", vbInformation
   n_reference_no_txt.SetFocus
   Exit Sub
  ElseIf n_title_txt.Text = "" Then
   MsgBox "Please enter title", vbInformation
   n_title_txt.SetFocus
   Exit Sub
  ElseIf n_language_Combo.Text = "------" Then
   MsgBox "Please select language", vbInformation
   n_language_Combo.SetFocus
   Exit Sub
  ElseIf n_pages_txt.Text = "" Then
   MsgBox "Please enter number of pages", vbInformation
   n_pages_txt.SetFocus
   Exit Sub
  ElseIf n_price_txt.Text = "" Then
   MsgBox "Please enter price", vbInformation
   n_price_txt.SetFocus
   Exit Sub
  Else
   rs2.Fields("Reference_Number").Value = n_reference_no_txt.Text
   rs2.Fields("Title").Value = n_title_txt.Text
   rs2.Fields("Stacked_On").Value = n_stacked_on_txt.Text
   rs2.Fields("Language").Value = n_language_Combo.Text
   rs2.Fields("Total_Pages").Value = n_pages_txt.Text
   rs2.Fields("Price").Value = n_price_txt.Text
   If str = "" Then
    str = "C:\visual_basic_project\Default_Pics\not_available.jpeg"
   End If
   rs2.Fields("Newspaper_Photo").Value = str
   On Error GoTo ERR_MSG
   rs2.AddNew
   MsgBox "Process completed successfully", vbInformation
   n_reference_no_txt.SetFocus
   Call cmd_refresh_Click(2)
End If
Exit Sub
ERR_MSG:
 MsgBox "Newspaper Reference Number : " & n_reference_no_txt.Text & "already exist !!! please change the Reference Number", vbExclamation
End Sub

Private Sub m_card_no_txt_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub m_name_txt_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub m_phone_no_txt_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
 Else
  KeyAscii = 0
  MsgBox "Enter numbers only", vbInformation
End If
End Sub

Private Sub cmd_m_add_Click()
If (m_name_txt.Text = "" And m_type_Combo.Text = "->Select type<-" And m_gender_Combo.Text = "->Select gender<-" And m_phone_no_txt.Text = "" And m_course_Combo.Text = "------" And m_reg_no_txt.Text = "" And m_card_no_txt.Text = "") Then
   MsgBox "Please enter all the fields", vbInformation
   Exit Sub
 ElseIf m_name_txt.Text = "" Then
   MsgBox "Please enter the name", vbInformation
   m_name_txt.SetFocus
   Exit Sub
 ElseIf m_type_Combo.Text = "->Select type<-" Then
   MsgBox "Please select type", vbInformation
   m_type_Combo.SetFocus
   Exit Sub
 ElseIf m_gender_Combo.Text = "->Select gender<-" Then
   MsgBox "Please select gender", vbInformation
   m_gender_Combo.SetFocus
   Exit Sub
 ElseIf m_phone_no_txt.Text = "" Then
   MsgBox "Please enter phone number", vbInformation
   m_phone_no_txt.SetFocus
   Exit Sub
  ElseIf Len(m_phone_no_txt.Text) <> 10 Then
   MsgBox "Please enter a valid Indian phone number", vbInformation
   m_phone_no_txt.SetFocus
   Exit Sub
 ElseIf m_course_Combo.Text = "------" Then
   MsgBox "Please select course", vbInformation
   m_course_Combo.SetFocus
   Exit Sub
 ElseIf m_reg_no_txt.Text = "" Then
   MsgBox "Please enter the register number", vbInformation
   m_reg_no_txt.SetFocus
   Exit Sub
 ElseIf m_card_no_txt.Text = "" Then
   MsgBox "Please enter the card number", vbInformation
   m_card_no_txt.SetFocus
   Exit Sub
 ElseIf m_card_vailidity_date_picker <= Now Then
   MsgBox "Please check the card validity date", vbInformation
   m_card_vailidity_date_picker.SetFocus
   Exit Sub
 Else
   rs.Fields("Name").Value = m_name_txt.Text
   rs.Fields("Type").Value = m_type_Combo.Text
   rs.Fields("Gender").Value = m_gender_Combo.Text
   rs.Fields("Phone_Number").Value = m_phone_no_txt.Text
   rs.Fields("Course").Value = m_course_Combo.Text
   rs.Fields("Register_Number").Value = m_reg_no_txt.Text
   rs.Fields("Card_Number").Value = m_card_no_txt.Text
   rs.Fields("Card_Issued_On").Value = m_card_issued_on_txt.Text
   rs.Fields("Card_Valid_Till").Value = m_card_vailidity_date_picker.Value
   rs.Fields("Card_Status").Value = m_card_status_txt.Text
   If str = "" And m_gender_Combo.Text = "Male" Then
    str = "C:\visual_basic_project\Default_Pics\male.jpeg"
    ElseIf str = "" And m_gender_Combo.Text = "Female" Then
     str = "C:\visual_basic_project\Default_Pics\female.jpeg"
    ElseIf str = "" Then
     str = "C:\visual_basic_project\Default_Pics\not_available.jpeg"
   End If
   rs.Fields("Member_Photo").Value = str
   On Error GoTo ERR_MSG
   rs.AddNew
   MsgBox "Process completed successfully", vbInformation
   m_name_txt.SetFocus
   Call cmd_refresh_Click(0)
End If
Exit Sub
ERR_MSG:
 MsgBox "Card Number : " & m_card_no_txt.Text & " holder already exist !!! please change the Card Number", vbExclamation
End Sub

Private Sub b_course_Combo_Click()
If b_course_Combo.Text = "Reference Book" Then
 b_year_Combo.Text = "N/A"
 b_year_Combo.Enabled = False
 b_sem_Combo.Text = "N/A"
 b_sem_Combo.Enabled = False
 Else
  b_year_Combo.Text = "----"
  b_year_Combo.Enabled = True
  b_sem_Combo.Text = "----"
  b_sem_Combo.Enabled = True
End If
End Sub

Private Sub cmd_add_book_Click()
Call refresh_fun
add_member_frame.Visible = False
add_book_frame.Visible = True
add_newspaper_frame.Visible = False
add_magazine_frame.Visible = False
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
 On Error GoTo ERR_MSG
 con1.Open "Provider=Microsoft.Jet.OlEDB.4.0;Data Source =  C:\visual_basic_project\ihprodb.mdb; Persist security Info=False"
 rs1.Open "select * from book_table", con1, adOpenDynamic, adLockPessimistic
 flag1 = 1
 rs1.AddNew
 Exit Sub
ERR_MSG:
 MsgBox "Please enter all the details and then click on Add button", vbInformation
End Sub

Private Sub cmd_add_magazine_Click()
Call refresh_fun
add_member_frame.Visible = False
add_book_frame.Visible = False
add_newspaper_frame.Visible = False
add_magazine_frame.Visible = True
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
 On Error GoTo ERR_MSG
 con3.Open "Provider=Microsoft.Jet.OlEDB.4.0;Data Source =  C:\visual_basic_project\ihprodb.mdb; Persist security Info=False"
 rs3.Open "select * from magazine_table", con3, adOpenDynamic, adLockPessimistic
 flag3 = 1
 rs3.AddNew
 Exit Sub
ERR_MSG:
 MsgBox "Please enter all the details and then click on Add button", vbInformation
End Sub

Private Sub cmd_add_member_Click()
Call refresh_fun
add_member_frame.Visible = True
add_book_frame.Visible = False
add_newspaper_frame.Visible = False
add_magazine_frame.Visible = False
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
 On Error GoTo ERR_MSG
 con.Open "Provider=Microsoft.Jet.OlEDB.4.0;Data Source =  C:\visual_basic_project\ihprodb.mdb; Persist security Info=False"
 rs.Open "select * from member_table", con, adOpenDynamic, adLockPessimistic
 flag = 1
 rs.AddNew
 Exit Sub
ERR_MSG:
 MsgBox "Please enter all the details and then click on Add button", vbInformation
End Sub

Private Sub cmd_add_newspaper_Click()
Call refresh_fun
add_member_frame.Visible = False
add_book_frame.Visible = False
add_newspaper_frame.Visible = True
add_magazine_frame.Visible = False
If (flag) Then
con.Close
flag = 0
End If
If (flag1) Then
 con1.Close
 flag1 = 0
End If
If (flag3) Then
 con3.Close
 flag3 = 0
End If
 On Error GoTo ERR_MSG
 con2.Open "Provider=Microsoft.Jet.OlEDB.4.0;Data Source =  C:\visual_basic_project\ihprodb.mdb; Persist security Info=False"
 rs2.Open "select * from newspaper_table", con2, adOpenDynamic, adLockPessimistic
 flag2 = 1
 rs2.AddNew
 Exit Sub
ERR_MSG:
 MsgBox "Please enter all the details and then click on Add button", vbInformation
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
 flag = 0
End If
If (flag3) Then
 con3.Close
 flag3 = 0
End If
Unload Me
End Sub

Private Sub cmd_browse_Click(Index As Integer)
CDC.Filter = " JPG (*.jpg) | *.jpg | JPEG (*.jpeg) | *jpeg | All Files (*.*) | *.*"
CDC.ShowOpen
If CDC.FileName <> "" Then
 str = CDC.FileName
 upload_pic(Index).Picture = LoadPicture(CDC.FileName)
End If
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
 flag = 0
End If
If (flag3) Then
 con3.Close
 flag3 = 0
End If
Unload Me
End Sub

Private Sub cmd_refresh_Click(Index As Integer)
Call refresh_fun
upload_pic(Index).Picture = Nothing
End Sub

Private Sub Form_Load()
register_bg.Move 0, 0, Me.Width, Me.Height
End Sub

Private Sub Form_Resize()
register_bg.Width = Me.ScaleWidth
register_bg.Height = Me.ScaleHeight
End Sub

Private Sub m_reg_no_txt_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub m_type_Combo_Click()
If m_type_Combo.Text = "Student" Then
 m_reg_no_lbl.Enabled = True
 m_reg_no_txt.Enabled = True
 m_course_lbl.Enabled = True
 m_course_Combo.Enabled = True
 m_reg_no_txt.Text = ""
 m_course_Combo.Text = "------"
 Else
  m_reg_no_txt.Enabled = False
  m_course_Combo.Enabled = False
  m_reg_no_txt.Text = "N/A"
  m_course_Combo.Text = "N/A"
End If
End Sub

Private Function refresh_fun()
m_name_txt.Text = ""
m_type_Combo.Text = "->Select type<-"
m_gender_Combo.Text = "->Select gender<-"
m_phone_no_txt.Text = ""
m_course_Combo.Text = "------"
m_reg_no_txt.Text = ""
m_card_no_txt.Text = ""
m_card_issued_on_txt.Text = system_date_lbl.Caption
m_card_vailidity_date_picker.Value = Now
m_card_status_txt.Text = "Active"
b_reference_no_txt.Text = ""
b_title_txt.Text = ""
b_stacked_on_txt.Text = system_date_lbl.Caption
b_isbn_code_txt.Text = ""
b_author_name_txt.Text = ""
b_publisher_name_txt.Text = ""
b_edition_txt.Text = ""
b_language_Combo.Text = "------"
b_course_Combo.Text = "------"
b_year_Combo.Text = "----"
b_sem_Combo.Text = "----"
b_pages_txt.Text = ""
b_price_txt.Text = ""
n_reference_no_txt.Text = ""
n_title_txt.Text = ""
n_language_Combo.Text = "------"
n_pages_txt.Text = ""
n_price_txt.Text = ""
n_stacked_on_txt.Text = system_date_lbl.Caption
mag_issn_code_txt.Text = ""
mag_reference_no_txt.Text = ""
mag_title_txt.Text = ""
mag_language_Combo.Text = "------"
mag_total_pages_txt.Text = ""
mag_price_txt.Text = ""
mag_stacked_on_txt.Text = system_date_lbl.Caption
str = ""
End Function

Private Sub mag_issn_code_txt_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or KeyAscii = 45 Then
 Else
  KeyAscii = 0
  MsgBox "Enter numbers only", vbInformation
End If
End Sub

Private Sub mag_reference_no_txt_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub mag_title_txt_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub n_reference_no_txt_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub n_title_txt_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub system_date_time_timer_Timer()
system_date_lbl.Caption = Date
system_time_lbl.Caption = Time
End Sub

