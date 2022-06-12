VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form SEARCH_RECORDS_FORM 
   ClientHeight    =   10935
   ClientLeft      =   2745
   ClientTop       =   510
   ClientWidth     =   20250
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc m_ado 
      Height          =   330
      Left            =   15480
      Top             =   1560
      Visible         =   0   'False
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\visual_basic_project\ihprodb.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\visual_basic_project\ihprodb.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "member_table"
      Caption         =   "m_ado"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Timer system_date_time_timer 
      Interval        =   100
      Left            =   19680
      Top             =   2640
   End
   Begin VB.Frame records_frame 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Records"
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
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   17175
      Begin VB.CommandButton cmd_search_member 
         BackColor       =   &H0080FFFF&
         Caption         =   "Search Member"
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
         Width           =   3255
      End
      Begin VB.CommandButton cmd_search_book 
         BackColor       =   &H0080FFFF&
         Caption         =   "Search Book"
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
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   480
         Width           =   3255
      End
      Begin VB.CommandButton cmd_search_newspaper 
         BackColor       =   &H0080FFFF&
         Caption         =   "Search Newspaper"
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
         Left            =   6960
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   480
         Width           =   3255
      End
      Begin VB.CommandButton cmd_search_magazine 
         BackColor       =   &H0080FFFF&
         Caption         =   "Search Magazine"
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
         Left            =   10320
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   480
         Width           =   3255
      End
      Begin VB.CommandButton cmd_members_list 
         BackColor       =   &H0080FFFF&
         Caption         =   "Members List"
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
         Left            =   13680
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   480
         Width           =   3255
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
      TabIndex        =   71
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
      Left            =   18120
      Style           =   1  'Graphical
      TabIndex        =   70
      Top             =   120
      Width           =   1935
   End
   Begin VB.Frame search_book_frame 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Search Book"
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
      Left            =   2880
      TabIndex        =   26
      Top             =   3240
      Visible         =   0   'False
      Width           =   14175
      Begin VB.TextBox sb_issued_to_txt 
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
         Left            =   11640
         TabIndex        =   125
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox sb_sem_txt 
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
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox sb_stacked_on_date_txt 
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
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   2640
         Width           =   2895
      End
      Begin VB.TextBox sb_course_txt 
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
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   3240
         Width           =   2895
      End
      Begin VB.TextBox sb_year_txt 
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
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   3840
         Width           =   1335
      End
      Begin VB.TextBox sb_status_txt 
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
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox sb_language_txt 
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
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   5040
         Width           =   2895
      End
      Begin VB.TextBox sb_price_txt 
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
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   5040
         Width           =   2175
      End
      Begin VB.TextBox sb_pages_txt 
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
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   4440
         Width           =   2895
      End
      Begin VB.TextBox sb_edition_txt 
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
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   4440
         Width           =   2895
      End
      Begin VB.TextBox sb_publisher_name_txt 
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
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   3840
         Width           =   2895
      End
      Begin VB.TextBox sb_reference_no_txt 
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
         TabIndex        =   27
         Top             =   480
         Width           =   2895
      End
      Begin VB.TextBox sb_title_txt 
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
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   2040
         Width           =   8655
      End
      Begin VB.TextBox sb_isbn_code_txt 
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
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   2640
         Width           =   2895
      End
      Begin VB.TextBox sb_author_name_txt 
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
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   3240
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
         TabIndex        =   43
         Top             =   4320
         Width           =   2175
      End
      Begin VB.TextBox sb_remark_txt 
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
         Height          =   975
         Left            =   8160
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   480
         Width           =   2895
      End
      Begin VB.CommandButton cmd_sb_search 
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
         TabIndex        =   44
         Top             =   4920
         Width           =   2175
      End
      Begin VB.Label sb_issued_to_lbl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Issued To "
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
         Left            =   12000
         TabIndex        =   124
         Top             =   360
         Width           =   1695
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   11640
         X2              =   14040
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label SBF_details_lbl 
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
         TabIndex        =   103
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label sb_language_lbl 
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
         TabIndex        =   102
         Top             =   5040
         Width           =   2055
      End
      Begin VB.Label sb_inr_lbl 
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
         TabIndex        =   42
         Top             =   5040
         Width           =   735
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   9960
         X2              =   9840
         Y1              =   3960
         Y2              =   4200
      End
      Begin VB.Label sb_price_lbl 
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
      Begin VB.Label sb_pages_lbl 
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
         TabIndex        =   100
         Top             =   4440
         Width           =   2055
      End
      Begin VB.Label sb_edition_lbl 
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
         TabIndex        =   99
         Top             =   4440
         Width           =   2295
      End
      Begin VB.Label sb_year_sem_lbl 
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
         TabIndex        =   98
         Top             =   3840
         Width           =   2175
      End
      Begin VB.Label sb_course_lbl 
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
         TabIndex        =   97
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   5880
         X2              =   5880
         Y1              =   2640
         Y2              =   5520
      End
      Begin VB.Label sb_publisher_name_lbl 
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
         TabIndex        =   96
         Top             =   3840
         Width           =   2295
      End
      Begin VB.Label sb_reference_no_lbl 
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
         TabIndex        =   95
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label sb_title_lbl 
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
         TabIndex        =   94
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label sb_author_name_lbl 
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
         TabIndex        =   93
         Top             =   3240
         Width           =   2295
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   120
         X2              =   11400
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   11520
         X2              =   11520
         Y1              =   360
         Y2              =   5520
      End
      Begin VB.Line Line11 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   11640
         X2              =   14040
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Label sb_isbn_code_lbl 
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
         TabIndex        =   92
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Label sb_stacked_date_lbl 
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
         TabIndex        =   91
         Top             =   2640
         Width           =   2295
      End
      Begin VB.Label sb_status_lbl 
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
         TabIndex        =   90
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label sb_remark_lbl 
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
         TabIndex        =   89
         Top             =   720
         Width           =   1215
      End
      Begin VB.Line Line12 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   6360
         X2              =   6360
         Y1              =   360
         Y2              =   1680
      End
      Begin VB.Image sb_upload_pic 
         BorderStyle     =   1  'Fixed Single
         Height          =   2430
         Left            =   11640
         Picture         =   "SEARCH_RECORDS_FORM.frx":0000
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   2445
      End
   End
   Begin VB.Frame search_newspaper_frame 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Search Newspaper"
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
      Left            =   5040
      TabIndex        =   45
      Top             =   3240
      Visible         =   0   'False
      Width           =   9975
      Begin VB.TextBox sn_status_txt 
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   5040
         Width           =   2895
      End
      Begin VB.TextBox sn_stacked_on_date_txt 
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
         Locked          =   -1  'True
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   2040
         Width           =   2895
      End
      Begin VB.TextBox sn_price_txt 
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
         Locked          =   -1  'True
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   4200
         Width           =   2175
      End
      Begin VB.TextBox sn_pages_txt 
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
         Locked          =   -1  'True
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   3480
         Width           =   2895
      End
      Begin VB.TextBox sn_reference_no_txt 
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
         TabIndex        =   46
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox sn_title_txt 
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
         Locked          =   -1  'True
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   1320
         Width           =   4335
      End
      Begin VB.TextBox sn_language_txt 
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
         Locked          =   -1  'True
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   2760
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
         Index           =   2
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   3480
         Width           =   2175
      End
      Begin VB.CommandButton cmd_sn_search 
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
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   4200
         Width           =   2175
      End
      Begin VB.TextBox sn_remark_txt 
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
         Left            =   6840
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   5040
         Width           =   2895
      End
      Begin VB.Label sn_inr_lbl 
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
         TabIndex        =   52
         Top             =   4200
         Width           =   735
      End
      Begin VB.Label sn_price_lbl 
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
         TabIndex        =   111
         Top             =   4200
         Width           =   2055
      End
      Begin VB.Label sn_pages_lbl 
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
         TabIndex        =   110
         Top             =   3480
         Width           =   2055
      End
      Begin VB.Label sn_reference_no_lbl 
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
         TabIndex        =   109
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label sn_title_lbl 
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
         TabIndex        =   108
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label sn_language_lbl 
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
         TabIndex        =   107
         Top             =   2760
         Width           =   2295
      End
      Begin VB.Line Line13 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   7320
         X2              =   7320
         Y1              =   360
         Y2              =   4800
      End
      Begin VB.Line Line14 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   7440
         X2              =   9720
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Label sn_stacked_date_lbl 
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
         TabIndex        =   106
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label sn_remark_lbl 
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
         TabIndex        =   105
         Top             =   5160
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
      Begin VB.Label sn_status_lbl 
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
         TabIndex        =   104
         Top             =   5040
         Width           =   1215
      End
      Begin VB.Image sn_upload_pic 
         BorderStyle     =   1  'Fixed Single
         Height          =   2430
         Left            =   7440
         Picture         =   "SEARCH_RECORDS_FORM.frx":193C
         Stretch         =   -1  'True
         Top             =   480
         Width           =   2445
      End
   End
   Begin VB.Frame search_magazine_frame 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Search Magazine"
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
      Left            =   5040
      TabIndex        =   57
      Top             =   3240
      Visible         =   0   'False
      Width           =   9975
      Begin VB.TextBox smag_status_txt 
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   5040
         Width           =   2895
      End
      Begin VB.TextBox smag_stacked_on_txt 
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
         Locked          =   -1  'True
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   2400
         Width           =   2895
      End
      Begin VB.TextBox smag_language_txt 
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
         Locked          =   -1  'True
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   3000
         Width           =   2895
      End
      Begin VB.TextBox smag_issn_code_txt 
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
         Locked          =   -1  'True
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   1200
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
         Index           =   3
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   3480
         Width           =   2175
      End
      Begin VB.TextBox smag_title_txt 
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
         Locked          =   -1  'True
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   1800
         Width           =   4335
      End
      Begin VB.TextBox smag_reference_no_txt 
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
         TabIndex        =   58
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox smag_total_pages_txt 
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
         Locked          =   -1  'True
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   3600
         Width           =   2895
      End
      Begin VB.TextBox smag_price_txt 
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
         Locked          =   -1  'True
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   4200
         Width           =   2175
      End
      Begin VB.CommandButton cmd_smag_search 
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
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   4200
         Width           =   2175
      End
      Begin VB.TextBox smag_remark_txt 
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
         Left            =   6840
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   5040
         Width           =   2895
      End
      Begin VB.Label smag_issn_code_lbl 
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
         TabIndex        =   120
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label smag_stacked_on_lbl 
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
         TabIndex        =   119
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Line Line15 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   7440
         X2              =   9840
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Line Line16 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   7320
         X2              =   7320
         Y1              =   360
         Y2              =   4800
      End
      Begin VB.Label smag_language_lbl 
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
         TabIndex        =   118
         Top             =   3000
         Width           =   2295
      End
      Begin VB.Label smag_title_lbl 
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
         TabIndex        =   117
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label smag_reference_no_lbl 
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
         TabIndex        =   116
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label smag_total_pages_lbl 
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
         TabIndex        =   115
         Top             =   3600
         Width           =   2055
      End
      Begin VB.Label smag_price_lbl 
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
         TabIndex        =   114
         Top             =   4200
         Width           =   2055
      End
      Begin VB.Label umag_inr_lbl 
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
         TabIndex        =   65
         Top             =   4200
         Width           =   735
      End
      Begin VB.Label smag_status_lbl 
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
         TabIndex        =   113
         Top             =   5040
         Width           =   1215
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
         TabIndex        =   112
         Top             =   5280
         Width           =   1215
      End
      Begin VB.Line Line17 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   120
         X2              =   9840
         Y1              =   4920
         Y2              =   4920
      End
      Begin VB.Image smag_upload_pic 
         BorderStyle     =   1  'Fixed Single
         Height          =   2430
         Left            =   7440
         Picture         =   "SEARCH_RECORDS_FORM.frx":3278
         Stretch         =   -1  'True
         Top             =   480
         Width           =   2445
      End
   End
   Begin MSDataGridLib.DataGrid M_DataGrid 
      Bindings        =   "SEARCH_RECORDS_FORM.frx":4BB4
      Height          =   7095
      Left            =   840
      TabIndex        =   123
      Top             =   2640
      Visible         =   0   'False
      Width           =   17175
      _ExtentX        =   30295
      _ExtentY        =   12515
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
      Caption         =   "Members List/Details"
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
   Begin VB.Frame search_member_frame 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Search Member"
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
      Height          =   5535
      Left            =   2880
      TabIndex        =   6
      Top             =   3240
      Visible         =   0   'False
      Width           =   13695
      Begin VB.TextBox sm_gender_txt 
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   4320
         Width           =   2895
      End
      Begin VB.TextBox sm_course_txt 
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   4920
         Width           =   2895
      End
      Begin VB.TextBox sm_type_txt 
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   3720
         Width           =   2895
      End
      Begin VB.TextBox sm_card_valid_till_date_txt 
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1800
         Width           =   2775
      End
      Begin VB.TextBox sm_card_issue_date_txt 
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox sm_tot_fine_amt_paid_txt 
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
         Left            =   8880
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox sm_no_of_times_suspended_txt 
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
         Left            =   9600
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox sm_no_of_times_terminated_txt 
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
         Left            =   9600
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox sm_fine_amt_txt 
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
         Left            =   7800
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   4080
         Width           =   1215
      End
      Begin VB.TextBox sm_remark_txt 
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
         Height          =   735
         Left            =   7800
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   4680
         Width           =   2895
      End
      Begin VB.TextBox sm_status_txt 
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
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   360
         Width           =   2895
      End
      Begin VB.CommandButton cmd_sm_search 
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
         Left            =   11280
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   4560
         Width           =   2175
      End
      Begin VB.TextBox sm_name_txt 
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   3120
         Width           =   2895
      End
      Begin VB.TextBox sm_reg_no_txt 
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
         Left            =   7800
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   3480
         Width           =   2895
      End
      Begin VB.TextBox sm_card_no_txt 
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
         Left            =   2400
         TabIndex        =   7
         Top             =   600
         Width           =   2775
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
         TabIndex        =   24
         Top             =   3720
         Width           =   2175
      End
      Begin VB.TextBox sm_phone_no_txt 
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
         Left            =   7800
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   2880
         Width           =   2895
      End
      Begin VB.Label sm_inr_lbl 
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
         Left            =   10080
         TabIndex        =   12
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label4 
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
         Left            =   2880
         TabIndex        =   88
         Top             =   600
         Width           =   735
      End
      Begin VB.Label sm_tot_fine_amt_paid_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Fine Amount Paid :"
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
         Left            =   5640
         TabIndex        =   87
         Top             =   960
         Width           =   3015
      End
      Begin VB.Label sm_no_of_times_suspended_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Number Of Times Suspended :"
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
         Left            =   5640
         TabIndex        =   86
         Top             =   2160
         Width           =   3735
      End
      Begin VB.Label sm_no_of_times_terminated_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Number Of Times Terminated :"
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
         Left            =   5640
         TabIndex        =   85
         Top             =   1560
         Width           =   3735
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
         Top             =   4080
         Width           =   1695
      End
      Begin VB.Label sm_fine_amt_lbl 
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
         TabIndex        =   84
         Top             =   4080
         Width           =   2055
      End
      Begin VB.Image sm_upload_pic 
         BorderStyle     =   1  'Fixed Single
         Height          =   2655
         Left            =   11160
         Picture         =   "SEARCH_RECORDS_FORM.frx":4BC8
         Stretch         =   -1  'True
         Top             =   480
         Width           =   2445
      End
      Begin VB.Label sm_remark_lbl 
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
         TabIndex        =   83
         Top             =   4800
         Width           =   1215
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   5400
         X2              =   5400
         Y1              =   360
         Y2              =   2640
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   5040
         X2              =   5040
         Y1              =   2880
         Y2              =   5400
      End
      Begin VB.Label sm_status_lbl 
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
         Left            =   5760
         TabIndex        =   82
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label SMF_details_lbl 
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
         TabIndex        =   81
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label sm_name_lbl 
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
         TabIndex        =   80
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label sm_type_lbl 
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
         TabIndex        =   79
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label sm_card_number_lbl 
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
         TabIndex        =   78
         Top             =   600
         Width           =   1935
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   120
         X2              =   10920
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   11040
         X2              =   11040
         Y1              =   360
         Y2              =   5400
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   11160
         X2              =   13560
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Label sm_gender_lbl 
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
         TabIndex        =   77
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Label sm_reg_no_lbl 
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
         TabIndex        =   76
         Top             =   3480
         Width           =   2295
      End
      Begin VB.Label sm_course_lbl 
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
         TabIndex        =   75
         Top             =   4920
         Width           =   1215
      End
      Begin VB.Label sm_card_issue_lbl 
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
         Left            =   480
         TabIndex        =   74
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label sm_card_valid_lbl 
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
         TabIndex        =   73
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label sm_phone_no_lbl 
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
         TabIndex        =   72
         Top             =   2880
         Width           =   2175
      End
   End
   Begin VB.Label msg_lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
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
      Height          =   855
      Left            =   600
      TabIndex        =   126
      Top             =   1560
      Width           =   16815
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
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   18240
      TabIndex        =   122
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
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   18240
      TabIndex        =   121
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Image search_records_bg 
      Height          =   10905
      Left            =   0
      Picture         =   "SEARCH_RECORDS_FORM.frx":5AAF
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20325
   End
End
Attribute VB_Name = "SEARCH_RECORDS_FORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim con1 As New ADODB.Connection
Dim con2 As New ADODB.Connection
Dim con3 As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
Dim flag, flag1, falg2, flag3 As Boolean

Private Sub cmd_back_Click()
HOME_FORM.Show
If flag Then
 con.Close
 flag = 0
End If
If flag1 Then
 con1.Close
 flag = 0
End If
If flag2 Then
 con2.Close
 flag = 0
End If
If flag3 Then
 con3.Close
 flag = 0
End If
Unload Me
End Sub

Private Sub cmd_logout_Click()
LOGIN_FORM.Show
If flag Then
 con.Close
 flag = 0
End If
If flag1 Then
 con1.Close
 flag = 0
End If
If flag2 Then
 con2.Close
 flag = 0
End If
If flag3 Then
 con3.Close
 flag = 0
End If
Unload Me
End Sub

Private Sub cmd_members_list_Click()
m_ado.Refresh
search_member_frame.Visible = False
search_book_frame.Visible = False
search_newspaper_frame.Visible = False
search_magazine_frame.Visible = False
M_DataGrid.Visible = True
msg_lbl.Visible = False
End Sub

Private Sub cmd_refresh_Click(Index As Integer)
Call refresh_fun
End Sub

Private Sub cmd_sb_search_Click()
If sb_reference_no_txt.Text = "" Then
 MsgBox "Please enter the reference number", vbInformation
 sb_reference_no_txt.SetFocus
 Exit Sub
 Else
  con1.Open "Provider=Microsoft.Jet.OlEDB.4.0;Data Source =  C:\visual_basic_project\ihprodb.mdb; Persist security Info=False"
  rs1.Open "select * from book_table where Reference_Number='" + sb_reference_no_txt.Text + "'", con1, adOpenDynamic, adLockPessimistic
  If Not rs1.EOF Then
   sb_reference_no_txt.Text = rs1!Reference_Number
   sb_status_txt.Text = rs1!Book_Status
   sb_remark_txt.Text = rs1!Remark
   sb_issued_to_txt.Text = rs1!Issued_To
   sb_title_txt.Text = rs1!Title
   sb_isbn_code_txt.Text = rs1!ISBN_Code
   sb_author_name_txt.Text = rs1!Author_name
   sb_publisher_name_txt.Text = rs1!Publisher
   sb_edition_txt.Text = rs1!Edition
   sb_language_txt.Text = rs1!Language
   sb_stacked_on_date_txt.Text = rs1!Stacked_On
   sb_course_txt.Text = rs1!Course
   sb_year_txt.Text = rs1!Course_Year
   sb_sem_txt.Text = rs1!Course_Sem
   sb_pages_txt.Text = rs1!Total_Pages
   sb_price_txt.Text = rs1!Price
   sb_upload_pic.Picture = LoadPicture(rs1!Book_Photo)
  Else
   MsgBox "Record not found... Please check the reference number!!!", vbCritical
 End If
 flag1 = 1
 reload_data
 con1.Close
End If
End Sub

Private Sub cmd_sm_search_Click()
If sm_card_no_txt.Text = "" Then
 MsgBox "Please enter the card number", vbInformation
 sm_card_no_txt.SetFocus
 Exit Sub
 Else
 con.Open "Provider=Microsoft.Jet.OlEDB.4.0;Data Source =  C:\visual_basic_project\ihprodb.mdb; Persist security Info=False"
 rs.Open "select * from member_table where Card_Number='" + sm_card_no_txt.Text + "' ", con, adOpenDynamic, adLockPessimistic
 If Not rs.EOF Then
  sm_card_no_txt.Text = rs!Card_Number
  sm_card_issue_date_txt.Text = rs!Card_Issued_On
  sm_card_valid_till_date_txt.Text = rs!Card_Valid_Till
  sm_status_txt.Text = rs!Card_Status
  sm_tot_fine_amt_paid_txt.Text = rs!Total_Fine_Paid
  sm_no_of_times_terminated_txt.Text = rs!Number_Of_Times_Terminated
  sm_no_of_times_suspended_txt.Text = rs!Number_Of_Times_Suspended
  sm_name_txt.Text = rs!Name
  sm_type_txt.Text = rs!Type
  sm_gender_txt.Text = rs!Gender
  sm_course_txt.Text = rs!Course
  sm_phone_no_txt.Text = rs!Phone_Number
  sm_reg_no_txt.Text = rs!Register_Number
  sm_fine_amt_txt.Text = rs!Fine_Balance
  sm_remark_txt.Text = rs!Remark
  sm_upload_pic.Picture = LoadPicture(rs!Member_Photo)
  Else
   MsgBox "Record not found... Please check the card number!!!", vbCritical
 End If
 flag = 1
 reload_data
 con.Close
End If
End Sub

Private Sub cmd_smag_search_Click()
If smag_reference_no_txt.Text = "" Then
  MsgBox "Please enter the reference number", vbInformation
  smag_reference_no_txt.SetFocus
  Exit Sub
  Else
   con2.Open "Provider=Microsoft.Jet.OlEDB.4.0;Data Source =  C:\visual_basic_project\ihprodb.mdb; Persist security Info=False"
   rs2.Open "select * from magazine_table where Reference_Number='" + smag_reference_no_txt.Text + "'", con2, adOpenDynamic, adLockPessimistic
   If Not rs2.EOF Then
    smag_reference_no_txt.Text = rs2!Reference_Number
    smag_issn_code_txt.Text = rs2!ISSN_Code
    smag_title_txt.Text = rs2!Title
    smag_stacked_on_txt.Text = rs2!Stacked_On
    smag_language_txt.Text = rs2!Language
    smag_total_pages_txt.Text = rs2!Total_Pages
    smag_price_txt.Text = rs2!Price
    smag_status_txt.Text = rs2!Magazine_Status
    smag_remark_txt.Text = rs2!Remark
    smag_upload_pic.Picture = LoadPicture(rs2!Magazine_Photo)
   Else
   MsgBox "Record not found... Please check the reference number!!!", vbCritical
 End If
 flag2 = 1
 reload_data
 con2.Close
End If
End Sub

Private Sub cmd_sn_search_Click()
If sn_reference_no_txt.Text = "" Then
  MsgBox "Please enter the reference number", vbInformation
  sn_reference_no_txt.SetFocus
  Exit Sub
  Else
   con3.Open "Provider=Microsoft.Jet.OlEDB.4.0;Data Source =  C:\visual_basic_project\ihprodb.mdb; Persist security Info=False"
   rs3.Open "select * from newspaper_table where Reference_Number='" + sn_reference_no_txt.Text + "'", con3, adOpenDynamic, adLockPessimistic
   If Not rs3.EOF Then
    sn_reference_no_txt.Text = rs3!Reference_Number
    sn_title_txt.Text = rs3!Title
    sn_stacked_on_date_txt.Text = rs3!Stacked_On
    sn_language_txt.Text = rs3!Language
    sn_pages_txt.Text = rs3!Total_Pages
    sn_price_txt.Text = rs3!Price
    sn_status_txt.Text = rs3!Newspaper_Status
    sn_remark_txt.Text = rs3!Remark
    sn_upload_pic.Picture = LoadPicture(rs3!Newspaper_Photo)
   Else
   MsgBox "Record not found... Please check the reference number!!!", vbCritical
 End If
 flag3 = 1
 reload_data
 con3.Close
End If
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
End Sub

Private Sub sb_reference_no_txt_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub sm_card_no_txt_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub sm_status_txt_Change()
If sm_status_txt.Text = "Active" Then
 sm_status_txt.BackColor = &HFF00&
 ElseIf sm_status_txt.Text = "Terminated" Then
  sm_status_txt.BackColor = &HFF&
 ElseIf sm_status_txt.Text = "Suspended" Then
  sm_status_txt.BackColor = &HFF8080
End If
End Sub

Private Sub smag_reference_no_txt_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub sn_reference_no_txt_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub system_date_time_timer_Timer()
system_date_lbl.Caption = Date
system_time_lbl.Caption = Time
End Sub

Private Sub cmd_search_book_Click()
Call refresh_fun
search_member_frame.Visible = False
search_book_frame.Visible = True
search_newspaper_frame.Visible = False
search_magazine_frame.Visible = False
M_DataGrid.Visible = False
msg_lbl.Visible = True
End Sub

Private Sub cmd_search_magazine_Click()
Call refresh_fun
search_member_frame.Visible = False
search_book_frame.Visible = False
search_newspaper_frame.Visible = False
search_magazine_frame.Visible = True
M_DataGrid.Visible = False
msg_lbl.Visible = True
End Sub

Private Sub cmd_search_member_Click()
Call refresh_fun
search_member_frame.Visible = True
search_book_frame.Visible = False
search_newspaper_frame.Visible = False
search_magazine_frame.Visible = False
M_DataGrid.Visible = False
msg_lbl.Visible = True
End Sub

Private Sub cmd_search_newspaper_Click()
Call refresh_fun
search_member_frame.Visible = False
search_book_frame.Visible = False
search_newspaper_frame.Visible = True
search_magazine_frame.Visible = False
M_DataGrid.Visible = False
msg_lbl.Visible = True
End Sub

Private Sub Form_Load()
search_records_bg.Move 0, 0, Me.Width, Me.Height
msg_lbl.Visible = True
msg_lbl.Caption = "Please enter the Card Number / Reference Number and click on search button"
End Sub

Private Sub Form_Resize()
search_records_bg.Width = Me.ScaleWidth
search_records_bg.Height = Me.ScaleHeight
End Sub

Private Function refresh_fun()
sm_card_no_txt.Text = ""
sm_card_issue_date_txt.Text = ""
sm_card_valid_till_date_txt.Text = ""
sm_status_txt.Text = ""
sm_tot_fine_amt_paid_txt.Text = ""
sm_no_of_times_terminated_txt.Text = ""
sm_no_of_times_suspended_txt.Text = ""
sm_name_txt.Text = ""
sm_type_txt.Text = ""
sm_gender_txt.Text = ""
sm_course_txt.Text = ""
sm_phone_no_txt.Text = ""
sm_reg_no_txt.Text = ""
sm_fine_amt_txt.Text = ""
sm_remark_txt.Text = ""
sm_upload_pic.Picture = Nothing
sm_status_txt.BackColor = &HFFFF80
sb_reference_no_txt.Text = ""
sb_status_txt.Text = ""
sb_remark_txt.Text = ""
sb_title_txt.Text = ""
sb_isbn_code_txt.Text = ""
sb_author_name_txt.Text = ""
sb_publisher_name_txt.Text = ""
sb_edition_txt.Text = ""
sb_language_txt.Text = ""
sb_stacked_on_date_txt.Text = ""
sb_course_txt.Text = ""
sb_year_txt.Text = ""
sb_sem_txt.Text = ""
sb_pages_txt.Text = ""
sb_price_txt.Text = ""
sb_issued_to_txt.Text = ""
sb_upload_pic.Picture = Nothing
sn_reference_no_txt.Text = ""
sn_title_txt.Text = ""
sn_stacked_on_date_txt.Text = ""
sn_language_txt.Text = ""
sn_pages_txt.Text = ""
sn_price_txt.Text = ""
sn_status_txt.Text = ""
sn_remark_txt.Text = ""
sn_upload_pic.Picture = Nothing
smag_reference_no_txt.Text = ""
smag_issn_code_txt.Text = ""
smag_title_txt.Text = ""
smag_stacked_on_txt.Text = ""
smag_language_txt.Text = ""
smag_total_pages_txt.Text = ""
smag_price_txt.Text = ""
smag_status_txt.Text = ""
smag_remark_txt.Text = ""
smag_upload_pic.Picture = Nothing
End Function
