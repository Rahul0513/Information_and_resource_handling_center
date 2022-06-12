VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ISS_REN_RET_FORM 
   ClientHeight    =   10935
   ClientLeft      =   4890
   ClientTop       =   1710
   ClientWidth     =   20250
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid IB_DataGrid 
      Height          =   1455
      Left            =   480
      TabIndex        =   58
      Top             =   8640
      Visible         =   0   'False
      Width           =   19455
      _ExtentX        =   34316
      _ExtentY        =   2566
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16776960
      ForeColor       =   16711680
      HeadLines       =   1
      RowHeight       =   23
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MV Boli"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid M_DataGrid 
      Height          =   1335
      Left            =   480
      TabIndex        =   56
      Top             =   7320
      Visible         =   0   'False
      Width           =   19455
      _ExtentX        =   34316
      _ExtentY        =   2355
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16776960
      ForeColor       =   16711680
      HeadLines       =   1
      RowHeight       =   23
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MV Boli"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Member Details"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Timer system_date_time_timer 
      Interval        =   100
      Left            =   15480
      Top             =   120
   End
   Begin VB.Frame upd_det_rec_option_frame 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Select"
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
      Height          =   2775
      Left            =   15960
      TabIndex        =   0
      Top             =   1440
      Width           =   4215
      Begin VB.CommandButton cmd_issue_book 
         BackColor       =   &H0080FFFF&
         Caption         =   "Issue Book"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   480
         Width           =   3495
      End
      Begin VB.CommandButton cmd_renew_book 
         BackColor       =   &H0080FFFF&
         Caption         =   "Renew Book"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1200
         Width           =   3495
      End
      Begin VB.CommandButton cmd_return_book 
         BackColor       =   &H0080FFFF&
         Caption         =   "Return Book"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1920
         Width           =   3495
      End
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
      Left            =   18120
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton cmd_back 
      BackColor       =   &H00C0E0FF&
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
      Left            =   16080
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   120
      Width           =   1935
   End
   Begin VB.Frame iss_ren_ret_frame 
      BackColor       =   &H00FFFFC0&
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
      Height          =   7095
      Left            =   1200
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   14175
      Begin VB.TextBox irr_issued_date_txt 
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
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   5760
         Width           =   2055
      End
      Begin VB.TextBox irr_fine_amt_txt 
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
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   2400
         Width           =   1935
      End
      Begin VB.CommandButton cmd_irr_return 
         BackColor       =   &H0080FF80&
         Caption         =   "Return"
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
         Left            =   9720
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   480
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton cmd_irr_renew 
         BackColor       =   &H0080FF80&
         Caption         =   "Renew"
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
         Left            =   9720
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   480
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox irr_book_status_txt 
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
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   5160
         Width           =   2895
      End
      Begin VB.TextBox irr_sem_txt 
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
         Left            =   10080
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox irr_year_txt 
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
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox irr_course_txt 
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
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   6360
         Width           =   2895
      End
      Begin VB.CommandButton cmd_irr_check 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Check"
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
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox irr_card_status_txt 
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
         ForeColor       =   &H00404040&
         Height          =   495
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1440
         Width           =   2895
      End
      Begin VB.TextBox irr_reference_no_txt 
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
         TabIndex        =   6
         Top             =   1200
         Width           =   2895
      End
      Begin VB.CommandButton cmd_irr_issue 
         BackColor       =   &H0080FF80&
         Caption         =   "Issue"
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
         Left            =   9720
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   480
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton cmd_irr_refresh 
         BackColor       =   &H00C0C0FF&
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
         Left            =   8040
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox irr_author_name_txt 
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
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   3960
         Width           =   2895
      End
      Begin VB.TextBox irr_isbn_no_txt 
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
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   3360
         Width           =   2895
      End
      Begin VB.TextBox irr_title_txt 
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
         Height          =   855
         Left            =   1680
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   2280
         Width           =   4215
      End
      Begin VB.TextBox irr_card_no_txt 
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
         TabIndex        =   5
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox irr_publisher_name_txt 
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
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   4560
         Width           =   2895
      End
      Begin VB.TextBox irr_edition_txt 
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
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   5160
         Width           =   2895
      End
      Begin VB.TextBox irr_pages_txt 
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
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   3960
         Width           =   2895
      End
      Begin VB.TextBox irr_price_txt 
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
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   4560
         Width           =   2175
      End
      Begin VB.TextBox irr_language_txt 
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
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   5760
         Width           =   2895
      End
      Begin MSComCtl2.DTPicker irr_due_date_date_picker 
         Height          =   495
         Left            =   8400
         TabIndex        =   27
         Top             =   6360
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
         Format          =   114360321
         CurrentDate     =   44123
      End
      Begin VB.Label irr_fine_amt_lbl 
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
         Left            =   6480
         TabIndex        =   52
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label irr_inr_bal_lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "INR "
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
         Left            =   10320
         TabIndex        =   15
         Top             =   2400
         Width           =   975
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   6240
         X2              =   6240
         Y1              =   2160
         Y2              =   3120
      End
      Begin VB.Image irr_member_pic 
         BorderStyle     =   1  'Fixed Single
         Height          =   2295
         Left            =   11640
         Picture         =   "ISS_REN_RET_FORM.frx":0000
         Stretch         =   -1  'True
         Top             =   720
         Width           =   2445
      End
      Begin VB.Image irr_book_pic 
         BorderStyle     =   1  'Fixed Single
         Height          =   2430
         Left            =   11640
         Picture         =   "ISS_REN_RET_FORM.frx":0EE7
         Stretch         =   -1  'True
         Top             =   4560
         Width           =   2445
      End
      Begin VB.Label irr_mem_phone_no_lbl 
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   11640
         TabIndex        =   30
         Top             =   3600
         Width           =   2415
      End
      Begin VB.Label irr_book_status_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Book Status :"
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
         TabIndex        =   51
         Top             =   5160
         Width           =   1815
      End
      Begin VB.Label irr_due_date_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         TabIndex        =   50
         Top             =   6360
         Width           =   2175
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   6000
         X2              =   6000
         Y1              =   360
         Y2              =   1920
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   6120
         X2              =   11400
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label irr_card_status_lbl 
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
         Left            =   6120
         TabIndex        =   49
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label irr_mem_name_lbl 
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   11640
         TabIndex        =   29
         Top             =   3120
         Width           =   2415
      End
      Begin VB.Label irr_book_lbl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Book"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   495
         Left            =   11760
         TabIndex        =   31
         Top             =   4080
         Width           =   2055
      End
      Begin VB.Label irr_member_lbl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Member"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   495
         Left            =   11760
         TabIndex        =   28
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label irr_reference_no_lbl 
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
         Left            =   240
         TabIndex        =   48
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   720
         X2              =   11040
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Label IRR_book_details_lbl 
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
         TabIndex        =   47
         Top             =   1800
         Width           =   2655
      End
      Begin VB.Label irr_issued_date_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Issued On :"
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
         TabIndex        =   46
         Top             =   5760
         Width           =   2175
      End
      Begin VB.Label irr_isbn_no_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ISBN Number :"
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
         TabIndex        =   45
         Top             =   3360
         Width           =   2175
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   11640
         X2              =   14040
         Y1              =   4080
         Y2              =   4080
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   11520
         X2              =   11520
         Y1              =   360
         Y2              =   6960
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   120
         X2              =   11400
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Label irr_author_name_lbl 
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
         TabIndex        =   44
         Top             =   3960
         Width           =   2295
      End
      Begin VB.Label irr_title_lbl 
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
         Left            =   240
         TabIndex        =   43
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label irr_card_no_lbl 
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
         TabIndex        =   42
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label irr_publisher_name_lbl 
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
         TabIndex        =   41
         Top             =   4560
         Width           =   2295
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   5880
         X2              =   5880
         Y1              =   3360
         Y2              =   6960
      End
      Begin VB.Label irr_course_lbl 
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
         Left            =   1200
         TabIndex        =   40
         Top             =   6360
         Width           =   1215
      End
      Begin VB.Label irr_year_sem_lbl 
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
         TabIndex        =   39
         Top             =   3360
         Width           =   2175
      End
      Begin VB.Label irr_edition_lbl 
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
         TabIndex        =   38
         Top             =   5160
         Width           =   2295
      End
      Begin VB.Label irr_pages_lbl 
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
         TabIndex        =   37
         Top             =   3960
         Width           =   2055
      End
      Begin VB.Label irr_price_lbl 
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
         TabIndex        =   36
         Top             =   4560
         Width           =   2055
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   9960
         X2              =   9840
         Y1              =   3480
         Y2              =   3720
      End
      Begin VB.Label irr_inr_lbl 
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
         TabIndex        =   35
         Top             =   4560
         Width           =   735
      End
      Begin VB.Label irr_language_lbl 
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
         TabIndex        =   34
         Top             =   5760
         Width           =   2055
      End
   End
   Begin MSDataGridLib.DataGrid B_DataGrid 
      Height          =   1335
      Left            =   480
      TabIndex        =   57
      Top             =   10080
      Visible         =   0   'False
      Width           =   19455
      _ExtentX        =   34316
      _ExtentY        =   2355
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16776960
      ForeColor       =   16711680
      HeadLines       =   1
      RowHeight       =   23
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MV Boli"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Book Details"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
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
      TabIndex        =   54
      Top             =   840
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
      Left            =   16200
      TabIndex        =   53
      Top             =   840
      Width           =   1815
   End
   Begin VB.Image iss_ren_ret_bg 
      Height          =   11040
      Left            =   0
      Picture         =   "ISS_REN_RET_FORM.frx":2823
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20235
   End
End
Attribute VB_Name = "ISS_REN_RET_FORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim con1 As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim flag, flag1, cfalg, cflag1, pass, pass1 As Boolean
Dim str, str1, str2, issue_status As String

Private Sub cmd_irr_check_Click()
If irr_card_no_txt.Text = "" And irr_reference_no_txt.Text = "" Then
 MsgBox "Please enter the card number and reference number of the book", vbInformation
 Exit Sub
 ElseIf irr_card_no_txt.Text = "" Then
  MsgBox "Please enter the card number", vbInformation
  irr_card_no_txt.SetFocus
  Exit Sub
 ElseIf irr_reference_no_txt.Text = "" Then
  MsgBox "Please enter the reference number of the book", vbInformation
  irr_reference_no_txt.SetFocus
  Exit Sub
 Else
  If flag Then
   rs.Close
   flag = 0
  End If
  rs.CursorLocation = adUseClient
REPEAT:
  On Error GoTo ERR_MSG
  rs.Open "select * from member_table where Card_Number='" + irr_card_no_txt.Text + "' ", con, adOpenDynamic, adLockPessimistic
  Set M_DataGrid.DataSource = rs
  If Not rs.EOF Then
   irr_card_no_txt.Text = rs!Card_Number
   issue_status = rs!Card_Number
   irr_card_status_txt.Text = rs!Card_Status
   irr_mem_name_lbl.Caption = rs!Name
   irr_mem_phone_no_lbl.Caption = "Ph: " & rs!Phone_Number
   irr_fine_amt_txt.Text = rs!Fine_Balance
   irr_member_pic.Picture = LoadPicture(rs!Member_Photo)
   pass = 1
   flag = 1
   Else
    MsgBox "Record not found... Please check the card number!!!", vbCritical
    flag = 1
    Exit Sub
  End If
  If flag1 Then
   rs1.Close
   flag1 = 0
  End If
  rs1.CursorLocation = adUseClient
REPEAT1:
  On Error GoTo ERR_MSG1
  rs1.Open "select * from book_table where Reference_Number='" + irr_reference_no_txt.Text + "'", con1, adOpenDynamic, adLockPessimistic
  Set B_DataGrid.DataSource = rs1
  If Not rs1.EOF Then
   irr_reference_no_txt.Text = rs1!Reference_Number
   irr_book_status_txt.Text = rs1!Book_Status
   irr_title_txt.Text = rs1!Title
   irr_isbn_no_txt.Text = rs1!ISBN_Code
   irr_author_name_txt.Text = rs1!Author_name
   irr_publisher_name_txt.Text = rs1!Publisher
   irr_edition_txt.Text = rs1!Edition
   irr_language_txt.Text = rs1!Language
   irr_course_txt.Text = rs1!Course
   irr_year_txt.Text = rs1!Course_Year
   irr_sem_txt.Text = rs1!Course_Sem
   irr_pages_txt.Text = rs1!Total_Pages
   irr_price_txt.Text = rs1!Price
   irr_book_pic.Picture = LoadPicture(rs1!Book_Photo)
   str1 = rs1!Issued_To
   str2 = rs1!Availability
   pass1 = 1
   flag1 = 1
  Else
   MsgBox "Record not found... Please check the reference number!!!", vbCritical
   flag1 = 1
   Exit Sub
 End If
 If pass And pass1 Then
   cmd_irr_return.Enabled = True
   cmd_irr_issue.Enabled = True
   cmd_irr_renew.Enabled = True
 End If
REPEAT2:
  On Error GoTo ERR_MSG2
  rs2.CursorLocation = adUseClient
  rs2.Open "select * from book_table where Issued_To='" + irr_card_no_txt.Text + "'", con1, adOpenDynamic, adLockPessimistic
  Set IB_DataGrid.DataSource = rs2
  IB_DataGrid.Caption = "Books Issued to card number : " + irr_card_no_txt.Text
End If
Exit Sub
ERR_MSG:
 rs.Close
 GoTo REPEAT
 Exit Sub
ERR_MSG1:
 rs1.Close
 GoTo REPEAT1
 Exit Sub
ERR_MSG2:
 rs2.Close
 GoTo REPEAT2
End Sub

Private Sub cmd_irr_issue_Click()
If irr_card_no_txt.Text = "" And irr_reference_no_txt.Text = "" Then
 MsgBox "Please enter the card number and reference number of the book", vbInformation
 Exit Sub
 ElseIf irr_card_no_txt.Text = "" Then
  MsgBox "Please enter the card number", vbInformation
  irr_card_no_txt.SetFocus
  Exit Sub
 ElseIf irr_reference_no_txt.Text = "" Then
  MsgBox "Please enter the reference number of the book", vbInformation
  irr_reference_no_txt.SetFocus
  Exit Sub
 ElseIf irr_book_status_txt.Text = "Missing" Then
 MsgBox "Book is missing !!! Book cannot be issued...", vbInformation
 Exit Sub
 ElseIf irr_card_status_txt.Text <> "Active" Then
  MsgBox "Card holder does not have permission to borrow books !!!", vbExclamation
  Exit Sub
 ElseIf str2 <> "Available" Then
  MsgBox "Book not available!!! Book is issued to card number : " + str1, vbInformation
  Exit Sub
 ElseIf rs2.RecordCount = 2 Then
  MsgBox "Process canceled !!! " & irr_card_no_txt.Text & " member has already borrowed two books... (maximum borrowing limit = 2 books)", vbInformation
  Exit Sub
 Else
  str = "Issued"
  issue_status = irr_card_no_txt.Text
  update_data
End If
End Sub

Private Sub cmd_irr_refresh_Click()
irr_issued_date_txt.Text = system_date_lbl.Caption
irr_due_date_date_picker.Value = Now
irr_card_no_txt.Text = ""
irr_reference_no_txt.Text = ""
irr_card_status_txt.Text = ""
irr_title_txt.Text = ""
irr_fine_amt_txt.Text = ""
irr_isbn_no_txt.Text = ""
irr_author_name_txt.Text = ""
irr_publisher_name_txt.Text = ""
irr_edition_txt.Text = ""
irr_language_txt.Text = ""
irr_course_txt.Text = ""
irr_year_txt.Text = ""
irr_sem_txt.Text = ""
irr_pages_txt.Text = ""
irr_price_txt.Text = ""
irr_book_status_txt.Text = ""
irr_mem_name_lbl.Caption = ""
irr_mem_phone_no_lbl.Caption = ""
irr_card_status_txt.BackColor = &HFFFF80
irr_member_pic.Picture = Nothing
irr_book_pic.Picture = Nothing
End Sub

Private Sub cmd_irr_renew_Click()
If irr_card_no_txt.Text = "" And irr_reference_no_txt.Text = "" Then
 MsgBox "Please enter the card number and reference number of the book", vbInformation
 Exit Sub
 ElseIf irr_card_no_txt.Text = "" Then
  MsgBox "Please enter the card number", vbInformation
  irr_card_no_txt.SetFocus
  Exit Sub
 ElseIf irr_reference_no_txt.Text = "" Then
  MsgBox "Please enter the reference number of the book", vbInformation
  irr_reference_no_txt.SetFocus
  Exit Sub
 ElseIf str2 = "Available" Then
  MsgBox "Book number : " + irr_reference_no_txt.Text + " is not yet issued to Card number : " + irr_card_no_txt.Text, vbInformation
  Exit Sub
 ElseIf rs2.RecordCount = 0 Then
  MsgBox "Process canceled !!! No books are issued to card number : " & irr_card_no_txt.Text, vbInformation
  Exit Sub
 Else
  str = "Renew/Issued"
  issue_status = irr_card_no_txt.Text
  update_data
End If
End Sub

Private Sub cmd_irr_return_Click()
If irr_card_no_txt.Text = "" And irr_reference_no_txt.Text = "" Then
 MsgBox "Please enter the card number and reference number of the book", vbInformation
 Exit Sub
 ElseIf irr_card_no_txt.Text = "" Then
  MsgBox "Please enter the card number", vbInformation
  irr_card_no_txt.SetFocus
  Exit Sub
 ElseIf irr_reference_no_txt.Text = "" Then
  MsgBox "Please enter the reference number of the book", vbInformation
  irr_reference_no_txt.SetFocus
  Exit Sub
 ElseIf str2 = "Available" Then
  MsgBox "Book number : " + irr_reference_no_txt.Text + " is not yet issued to Card number : " + irr_card_no_txt.Text, vbInformation
  Exit Sub
 ElseIf rs2.RecordCount = 0 Then
  MsgBox "Process canceled !!! No books are issued to card number : " & irr_card_no_txt.Text, vbInformation
  Exit Sub
 Else
  str = "Available"
  issue_status = "------"
  update_data
End If
End Sub

Private Sub Form_Load()
iss_ren_ret_bg.Move 0, 0, Me.Width, Me.Height
REPEAT:
 On Error GoTo ERR_MSG
 con.Open "Provider=Microsoft.Jet.OlEDB.4.0;Data Source =  C:\visual_basic_project\ihprodb.mdb; Persist security Info=False"
 cflag = 1
REPEAT1:
 On Error GoTo ERR_MSG1
 con1.Open "Provider=Microsoft.Jet.OlEDB.4.0;Data Source =  C:\visual_basic_project\ihprodb.mdb; Persist security Info=False"
 cflag1 = 1
 Exit Sub
ERR_MSG:
 con.Close
 GoTo REPEAT
 Exit Sub
ERR_MSG1:
 con1.Close
 GoTo REPEAT1
End Sub

Private Sub Form_Resize()
iss_ren_ret_bg.Width = Me.ScaleWidth
iss_ren_ret_bg.Height = Me.ScaleHeight
End Sub

Private Sub cmd_back_Click()
HOME_FORM.Show
If (flag Or cflag) Then
con.Close
cflag = 0
flag = 0
End If
If (flag1 Or cflag1) Then
 con1.Close
 cflag1 = 0
 flag1 = 0
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
Unload Me
End Sub

Private Sub cmd_issue_book_Click()
Call cmd_irr_refresh_Click
iss_ren_ret_frame.Caption = "Issue Book"
iss_ren_ret_frame.Visible = True
cmd_irr_return.Visible = False
cmd_irr_return.Enabled = False
cmd_irr_renew.Visible = False
cmd_irr_renew.Enabled = False
cmd_irr_issue.Visible = True
irr_due_date_lbl.Caption = "Due Date :"
M_DataGrid.Visible = True
B_DataGrid.Visible = True
IB_DataGrid.Visible = True
disable
End Sub

Private Sub cmd_renew_book_Click()
Call cmd_irr_refresh_Click
iss_ren_ret_frame.Caption = "Renew Book"
iss_ren_ret_frame.Visible = True
cmd_irr_return.Visible = False
cmd_irr_return.Enabled = False
cmd_irr_renew.Visible = True
cmd_irr_issue.Visible = False
cmd_irr_issue.Enabled = False
irr_due_date_lbl.Caption = "Renew Due Date :"
M_DataGrid.Visible = True
B_DataGrid.Visible = True
IB_DataGrid.Visible = True
disable
End Sub

Private Sub cmd_return_book_Click()
Call cmd_irr_refresh_Click
iss_ren_ret_frame.Caption = "Return Book"
iss_ren_ret_frame.Visible = True
cmd_irr_return.Visible = True
cmd_irr_renew.Visible = False
cmd_irr_renew.Enabled = False
cmd_irr_issue.Visible = False
cmd_irr_issue.Enabled = False
irr_due_date_lbl.Caption = "Returned On :"
M_DataGrid.Visible = True
B_DataGrid.Visible = True
IB_DataGrid.Visible = True
disable
End Sub

Private Sub irr_card_no_txt_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub irr_card_status_txt_Change()
If irr_card_status_txt.Text = "Active" Then
 irr_card_status_txt.BackColor = &HFF00&
 ElseIf irr_card_status_txt.Text = "Terminated" Then
  irr_card_status_txt.BackColor = &HFF&
 ElseIf irr_card_status_txt.Text = "Suspended" Then
  irr_card_status_txt.BackColor = &HFF8080
End If
End Sub

Private Sub irr_fine_amt_txt_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or KeyAscii = 46 Then
 Else
  KeyAscii = 0
  MsgBox "Enter numbers only", vbInformation
End If
End Sub

Private Sub irr_reference_no_txt_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub system_date_time_timer_Timer()
system_date_lbl.Caption = Date
system_time_lbl.Caption = Time
End Sub

Public Sub update_data()
   rs1.Fields("Reference_Number").Value = irr_reference_no_txt.Text
   rs1.Fields("Title").Value = irr_title_txt.Text
   rs1.Fields("ISBN_Code").Value = irr_isbn_no_txt.Text
   rs1.Fields("Author_Name").Value = irr_author_name_txt.Text
   rs1.Fields("Publisher").Value = irr_publisher_name_txt.Text
   rs1.Fields("Edition").Value = irr_edition_txt.Text
   rs1.Fields("Language").Value = irr_language_txt.Text
   rs1.Fields("Course").Value = irr_course_txt.Text
   rs1.Fields("Course_year").Value = irr_year_txt.Text
   rs1.Fields("Course_Sem").Value = irr_sem_txt.Text
   rs1.Fields("Total_pages").Value = irr_pages_txt.Text
   rs1.Fields("Price").Value = irr_price_txt.Text
   rs1.Fields("Book_Status").Value = irr_book_status_txt.Text
   rs1.Fields("Issued_On").Value = irr_issued_date_txt.Text
   rs1.Fields("Last_Borrower").Value = irr_card_no_txt.Text
   rs1.Fields("Due_Or_Return_Date").Value = irr_due_date_date_picker.Value
   rs1.Fields("Issued_To").Value = issue_status
   rs1.Fields("Availability").Value = str
   rs1.Update
   MsgBox "Process completed successfully", vbInformation
   disable
   flag1 = 1
End Sub

Sub disable()
cmd_irr_refresh_Click
cmd_irr_return.Enabled = False
cmd_irr_issue.Enabled = False
cmd_irr_renew.Enabled = False
End Sub

