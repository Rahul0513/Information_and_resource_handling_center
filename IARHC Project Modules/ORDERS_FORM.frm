VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ORDERS_FORM 
   ClientHeight    =   10935
   ClientLeft      =   4605
   ClientTop       =   2610
   ClientWidth     =   20250
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   600
      Top             =   5760
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      RecordSource    =   "orders_table"
      Caption         =   "Adodc1"
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
   Begin MSComDlg.CommonDialog CDC 
      Left            =   15120
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer system_date_time_timer 
      Interval        =   100
      Left            =   15120
      Top             =   120
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
      Left            =   3840
      TabIndex        =   0
      Top             =   120
      Width           =   11175
      Begin VB.CommandButton cmd_view_list 
         BackColor       =   &H0080FFFF&
         Caption         =   "View List"
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
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   480
         Width           =   3255
      End
      Begin VB.CommandButton cmd_update_order_list 
         BackColor       =   &H0080FFFF&
         Caption         =   "Update Order List"
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
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   480
         Width           =   3255
      End
      Begin VB.CommandButton cmd_create_order_list 
         BackColor       =   &H0080FFFF&
         Caption         =   "Create Order List"
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
         Width           =   3015
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
      TabIndex        =   26
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
      TabIndex        =   25
      Top             =   120
      Width           =   1935
   End
   Begin VB.Frame create_order_frame 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Create Order"
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
      Left            =   5040
      TabIndex        =   3
      Top             =   3240
      Visible         =   0   'False
      Width           =   9975
      Begin VB.ComboBox cr_o_language_Combo 
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
         ItemData        =   "ORDERS_FORM.frx":0000
         Left            =   2760
         List            =   "ORDERS_FORM.frx":0019
         TabIndex        =   8
         Text            =   "------"
         Top             =   3480
         Width           =   2895
      End
      Begin VB.TextBox cr_o_isbn_issn_code_txt 
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
         TabIndex        =   7
         Top             =   2760
         Width           =   2895
      End
      Begin VB.CommandButton cmd_cr_o_create 
         BackColor       =   &H0080FF80&
         Caption         =   "Create"
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
         TabIndex        =   12
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
         Index           =   0
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   11
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
         Index           =   0
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2880
         Width           =   2175
      End
      Begin VB.TextBox cr_o_title_txt 
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
         Top             =   2040
         Width           =   4335
      End
      Begin VB.TextBox cr_o_order_no_txt 
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
         TabIndex        =   4
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox cr_o_quantity_txt 
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
         TabIndex        =   9
         Top             =   4200
         Width           =   2895
      End
      Begin VB.ComboBox cr_o_order_type_Combo 
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
         ItemData        =   "ORDERS_FORM.frx":0056
         Left            =   2760
         List            =   "ORDERS_FORM.frx":0066
         TabIndex        =   5
         Text            =   "------"
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Image upload_pic 
         BorderStyle     =   1  'Fixed Single
         Height          =   2430
         Index           =   0
         Left            =   7560
         Picture         =   "ORDERS_FORM.frx":0097
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2205
      End
      Begin VB.Label cr_o_isbn_issn_code_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ISBN / ISSN Code:"
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
         TabIndex        =   32
         Top             =   2760
         Width           =   2415
      End
      Begin VB.Line Line15 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   7440
         X2              =   9840
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Line Line16 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   7320
         X2              =   7320
         Y1              =   360
         Y2              =   4800
      End
      Begin VB.Label cr_o_language_lbl 
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
         TabIndex        =   31
         Top             =   3480
         Width           =   2295
      End
      Begin VB.Label cr_o_title_lbl 
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
         TabIndex        =   30
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label cr_o_order_no_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Order Number :"
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
         TabIndex        =   29
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label cr_o_quantity_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity :"
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
         TabIndex        =   28
         Top             =   4200
         Width           =   2055
      End
      Begin VB.Label cr_o_order_type_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Order Type :"
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
         TabIndex        =   27
         Top             =   1320
         Width           =   2055
      End
   End
   Begin VB.Frame update_order_frame 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Update Order"
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
      Left            =   5040
      TabIndex        =   13
      Top             =   3240
      Visible         =   0   'False
      Width           =   9975
      Begin VB.ComboBox up_o_language_Combo 
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
         ItemData        =   "ORDERS_FORM.frx":19D3
         Left            =   2760
         List            =   "ORDERS_FORM.frx":19EC
         TabIndex        =   20
         Text            =   "------"
         Top             =   3600
         Width           =   2895
      End
      Begin VB.CommandButton cmd_o_update 
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
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   4200
         Width           =   2175
      End
      Begin VB.CommandButton cmd_up_o_check 
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
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   600
         Width           =   1695
      End
      Begin VB.ComboBox up_o_status_Combo 
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
         ItemData        =   "ORDERS_FORM.frx":1A29
         Left            =   2760
         List            =   "ORDERS_FORM.frx":1A36
         TabIndex        =   16
         Text            =   "------"
         Top             =   1200
         Width           =   2895
      End
      Begin VB.ComboBox up_o_order_type_Combo 
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
         ItemData        =   "ORDERS_FORM.frx":1A68
         Left            =   2760
         List            =   "ORDERS_FORM.frx":1A78
         TabIndex        =   17
         Text            =   "------"
         Top             =   1800
         Width           =   2895
      End
      Begin VB.TextBox up_o_quantity_txt 
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
         TabIndex        =   21
         Top             =   4200
         Width           =   2895
      End
      Begin VB.TextBox up_o_order_no_txt 
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
         TabIndex        =   14
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox up_o_title_txt 
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
         TabIndex        =   18
         Top             =   2400
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
         Index           =   1
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   22
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
         Index           =   1
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   3600
         Width           =   2175
      End
      Begin VB.TextBox up_o_isbn_issn_code_txt 
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
         TabIndex        =   19
         Top             =   3000
         Width           =   2895
      End
      Begin VB.Image upload_pic 
         BorderStyle     =   1  'Fixed Single
         Height          =   2430
         Index           =   1
         Left            =   7680
         Picture         =   "ORDERS_FORM.frx":1AA9
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2205
      End
      Begin VB.Label up_o_status_lbl 
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
         Left            =   480
         TabIndex        =   39
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label up_o_order_type_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Order Type :"
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
         TabIndex        =   38
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label up_o_quantity_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity :"
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
         TabIndex        =   37
         Top             =   4200
         Width           =   2055
      End
      Begin VB.Label up_o_order_no_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Order Number :"
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
         TabIndex        =   36
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label up_o_title_lbl 
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
         TabIndex        =   35
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label up_o_language_lbl 
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
         TabIndex        =   34
         Top             =   3600
         Width           =   2295
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   7560
         X2              =   7560
         Y1              =   360
         Y2              =   4800
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   7680
         X2              =   9840
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Label up_o_isbn_issn_code_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ISBN / ISSN Code:"
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
         TabIndex        =   33
         Top             =   3000
         Width           =   2415
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "ORDERS_FORM.frx":33E5
      Height          =   6615
      Left            =   840
      TabIndex        =   42
      Top             =   2280
      Visible         =   0   'False
      Width           =   18135
      _ExtentX        =   31988
      _ExtentY        =   11668
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
      Caption         =   "Orders List"
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
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   16080
      TabIndex        =   41
      Top             =   720
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
      Left            =   16080
      TabIndex        =   40
      Top             =   240
      Width           =   1815
   End
   Begin VB.Image orders_bg 
      Height          =   10935
      Left            =   0
      Picture         =   "ORDERS_FORM.frx":33FA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20295
   End
End
Attribute VB_Name = "ORDERS_FORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim flag, flag1 As Boolean
Dim str As String

Private Sub cmd_back_Click()
HOME_FORM.Show
On Error Resume Next
con.Close
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

Private Sub cmd_cr_o_create_Click()
If cr_o_order_no_txt.Text = "" And cr_o_order_type_Combo.Text = "------" And cr_o_title_txt.Text = "" And cr_o_isbn_issn_code_txt.Text = "" And cr_o_language_Combo.Text = "------" And cr_o_quantity_txt.Text = "" Then
   MsgBox "Please enter all the fields", vbInformation
   Exit Sub
  ElseIf cr_o_order_no_txt.Text = "" Then
   MsgBox "Please enter Order number", vbInformation
   cr_o_order_no_txt.SetFocus
   Exit Sub
  ElseIf cr_o_order_type_Combo.Text = "------" Then
   MsgBox "Please select type", vbInformation
   cr_o_order_type_Combo.SetFocus
   Exit Sub
  ElseIf cr_o_title_txt.Text = "" Then
   MsgBox "Please metion title", vbInformation
   cr_o_title_txt.SetFocus
   Exit Sub
  ElseIf cr_o_isbn_issn_code_txt.Text = "" Then
   MsgBox "Please mention ISBN/ISSN code", vbInformation
   cr_o_isbn_issn_code_txt.SetFocus
   Exit Sub
  ElseIf cr_o_language_Combo.Text = "------" Then
   MsgBox "Please select language", vbInformation
   cr_o_language_Combo.SetFocus
   Exit Sub
  ElseIf cr_o_quantity_txt.Text = "" Then
   MsgBox "Please mention quantity", vbInformation
   cr_o_quantity_txt.SetFocus
   Exit Sub
  Else
REPEAT:
   On Error GoTo ERR_MSG
   rs.Fields("Order_Number").Value = cr_o_order_no_txt.Text
   rs.Fields("Order_Type").Value = cr_o_order_type_Combo.Text
   rs.Fields("Title").Value = cr_o_title_txt.Text
   rs.Fields("ISSN_Or_ISBN_Code").Value = cr_o_isbn_issn_code_txt.Text
   rs.Fields("Language").Value = cr_o_language_Combo.Text
   rs.Fields("Quantity").Value = cr_o_quantity_txt.Text
   If str = "" Then
     str = "C:\visual_basic_project\Default_Pics\not_available.jpeg"
   End If
   rs.Fields("Photo").Value = str
   MsgBox "Order successfully created", vbInformation
   On Error GoTo ERR_MSG1
   rs.AddNew
   Call cmd_refresh_Click(0)
   flag = 1
End If
Exit Sub
ERR_MSG:
 rs.AddNew
 GoTo REPEAT
ERR_MSG1:
 MsgBox "Card Number : " & m_card_no_txt.Text & " holder already exist !!! please change the Card Number", vbExclamation
End Sub

Private Sub cmd_create_order_list_Click()
Call refresh_fun
create_order_frame.Visible = True
update_order_frame.Visible = False
DataGrid1.Visible = False
If flag1 Then
 con.Close
 flag1 = 0
End If
REPEAT:
On Error GoTo ERR_MSG
con.Open "Provider=Microsoft.Jet.OlEDB.4.0;Data Source =  C:\visual_basic_project\ihprodb.mdb; Persist security Info=False"
rs.Open "select * from orders_table", con, adOpenDynamic, adLockPessimistic
flag = 1
On Error GoTo ERR_MSG
rs.AddNew
Exit Sub
ERR_MSG:
 con.Close
 GoTo REPEAT
End Sub

Private Sub cmd_logout_Click()
LOGIN_FORM.Show
On Error Resume Next
con.Close
Unload Me
End Sub

Private Sub cmd_refresh_Click(Index As Integer)
upload_pic(Index).Picture = Nothing
Call refresh_fun
End Sub

Private Sub cmd_up_o_check_Click()
If up_o_order_no_txt = "" Then
   MsgBox "Please enter the order number", vbInformation
   Exit Sub
 Else
REPEAT:
 On Error GoTo ERR_MSG
 rs1.Open "select * from orders_table where Order_Number='" + up_o_order_no_txt.Text + "'", con, adOpenDynamic, adLockPessimistic
 If Not rs1.EOF Then
    up_o_order_no_txt.Text = rs1!Order_Number
    up_o_order_type_Combo.Text = rs1!Order_Type
    up_o_title_txt.Text = rs1!Title
    up_o_isbn_issn_code_txt.Text = rs1!ISSN_Or_ISBN_Code
    up_o_language_Combo.Text = rs1!Language
    up_o_quantity_txt.Text = rs1!Quantity
    up_o_status_Combo.Text = rs1!Status
    upload_pic(1) = LoadPicture(rs1!Photo)
    up_o_status_Combo.Enabled = True
    up_o_order_type_Combo.Enabled = True
    up_o_title_txt.Enabled = True
    up_o_isbn_issn_code_txt.Enabled = True
    up_o_language_Combo.Enabled = True
    up_o_quantity_txt.Enabled = True
    cmd_browse(1).Enabled = True
    cmd_o_update.Enabled = True
  Else
   MsgBox "Record not found ... Please check the order number", vbInformation
 End If
End If
Exit Sub
ERR_MSG:
 rs1.Close
 GoTo REPEAT
End Sub

Private Sub cmd_o_update_Click()
If up_o_order_no_txt.Text = "" Then
   MsgBox "Please enter Order number", vbInformation
   up_o_order_no_txt.SetFocus
   Exit Sub
  ElseIf up_o_status_Combo.Text = "------" Then
   MsgBox "Please select order status", vbInformation
   up_o_status_Combo.SetFocus
   Exit Sub
  ElseIf up_o_order_type_Combo.Text = "------" Then
   MsgBox "Please select order type", vbInformation
   up_o_order_type_Combo.SetFocus
   Exit Sub
  ElseIf up_o_title_txt.Text = "" Then
   MsgBox "Please select order type", vbInformation
   up_o_title_txt.SetFocus
   Exit Sub
  ElseIf up_o_isbn_issn_code_txt.Text = "------" Then
   MsgBox "Please mention the ISBN / ISSN code", vbInformation
   up_o_isbn_issn_code_txt.SetFocus
   Exit Sub
  ElseIf up_o_language_Combo.Text = "" Then
   MsgBox "Please select language", vbInformation
   up_o_language_Combo.SetFocus
   Exit Sub
  Else
   rs1.Fields("Order_Number") = up_o_order_no_txt.Text
   rs1.Fields("Order_Type").Value = up_o_order_type_Combo.Text
   rs1.Fields("Title").Value = up_o_title_txt.Text
   rs1.Fields("ISSN_Or_ISBN_Code").Value = up_o_isbn_issn_code_txt.Text
   rs1.Fields("Language").Value = up_o_language_Combo.Text
   rs1.Fields("Quantity").Value = up_o_quantity_txt.Text
   rs1.Fields("Status").Value = up_o_status_Combo.Text
   If str = "" Then
     str = "C:\visual_basic_project\Default_Pics\not_available.jpeg"
   End If
   rs1.Fields("Photo").Value = str
   rs1.Update
   disable
   MsgBox "Information Updated successfully", vbInformation
   Call cmd_refresh_Click(1)
   flag1 = 1
End If
End Sub

Private Sub cmd_update_order_list_Click()
Call refresh_fun
create_order_frame.Visible = False
update_order_frame.Visible = True
DataGrid1.Visible = False
If flag Then
 con.Close
 flag = 0
End If
On Error GoTo ERR_MSG
con.Open "Provider=Microsoft.Jet.OlEDB.4.0;Data Source =  C:\visual_basic_project\ihprodb.mdb; Persist security Info=False"
flag1 = 1
Exit Sub
ERR_MSG:
 MsgBox "Please enter the values and click on Check button then on Update button", vbInformation
End Sub

Private Sub cmd_view_list_Click()
Call refresh_fun
create_order_frame.Visible = False
update_order_frame.Visible = False
Adodc1.Refresh
DataGrid1.Refresh
DataGrid1.Visible = True
End Sub

Private Sub cr_o_isbn_issn_code_txt_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or KeyAscii = 45 Then
 Else
  KeyAscii = 0
  MsgBox "Enter numbers only", vbInformation
End If
End Sub

Private Sub cr_o_order_no_txt_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cr_o_order_type_Combo_Click()
If cr_o_order_type_Combo = "Book" Or cr_o_order_type_Combo = "Magazine" Then
 cr_o_isbn_issn_code_txt.Enabled = True
 cr_o_isbn_issn_code_txt.Text = ""
 Else
  cr_o_isbn_issn_code_txt.Enabled = False
  cr_o_isbn_issn_code_txt.Text = "N/A"
 End If
 If cr_o_order_type_Combo = "Book" Or cr_o_order_type_Combo = "Magazine" Or cr_o_order_type_Combo = "Newspaper" Then
  cr_o_language_Combo.Enabled = True
  cr_o_language_Combo.Text = "------  "
  Else
   cr_o_language_Combo.Enabled = False
   cr_o_language_Combo.Text = "N/A"
End If
End Sub

Private Sub cr_o_quantity_txt_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
 Else
  KeyAscii = 0
  MsgBox "Enter numbers only", vbInformation
End If
End Sub

Private Sub cr_o_title_txt_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Form_Load()
orders_bg.Move 0, 0, Me.Width, Me.Height
End Sub

Private Sub Form_Resize()
orders_bg.Width = Me.ScaleWidth
orders_bg.Height = Me.ScaleHeight
End Sub

Private Function refresh_fun()
up_o_order_no_txt.Text = ""
up_o_status_Combo.Text = "------"
up_o_order_type_Combo.Text = "------"
up_o_title_txt.Text = ""
up_o_isbn_issn_code_txt.Text = "------"
up_o_language_Combo.Text = ""
up_o_quantity_txt.Text = ""
cr_o_order_no_txt.Text = ""
cr_o_order_type_Combo.Text = "------"
cr_o_title_txt.Text = ""
cr_o_isbn_issn_code_txt.Text = ""
cr_o_language_Combo.Text = "------"
cr_o_quantity_txt.Text = ""
str = ""
End Function

Private Sub system_date_time_timer_Timer()
system_date_lbl.Caption = Date
system_time_lbl.Caption = Time
End Sub

Private Sub up_o_isbn_issn_code_txt_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or KeyAscii = 45 Then
 Else
  KeyAscii = 0
  MsgBox "Enter numbers only", vbInformation
End If
End Sub

Private Sub up_o_order_no_txt_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub up_o_quantity_txt_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
 Else
  KeyAscii = 0
  MsgBox "Enter numbers only", vbInformation
End If
End Sub

Private Sub up_o_title_txt_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Sub disable()
up_o_status_Combo.Enabled = False
up_o_order_type_Combo.Enabled = False
up_o_title_txt.Enabled = False
up_o_isbn_issn_code_txt.Enabled = False
up_o_language_Combo.Enabled = False
up_o_quantity_txt.Enabled = False
cmd_browse(1).Enabled = False
cmd_o_update.Enabled = False
End Sub
