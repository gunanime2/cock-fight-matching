VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form form_new_event 
   Caption         =   "Events"
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20250
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.Frame frame_edit_entry 
      Height          =   4215
      Left            =   10440
      TabIndex        =   97
      Top             =   1920
      Visible         =   0   'False
      Width           =   4575
      Begin VB.CommandButton command_edit_entry_save 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   104
         Top             =   3600
         Width           =   1575
      End
      Begin MSDataListLib.DataCombo datacombo_edit_entry_schedule_type 
         Bindings        =   "form_new_event.frx":0000
         Height          =   420
         Left            =   1680
         TabIndex        =   103
         Top             =   3000
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   741
         _Version        =   393216
         ListField       =   "entry_schedule_type"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox text_edit_entry_weight 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         TabIndex        =   102
         Top             =   2280
         Width           =   2775
      End
      Begin VB.TextBox text_edit_entry_wing_band 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         TabIndex        =   101
         Top             =   1560
         Width           =   2775
      End
      Begin VB.TextBox text_edit_entry_leg_band 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         TabIndex        =   100
         Top             =   840
         Width           =   2775
      End
      Begin VB.CommandButton command_edit_entry_close 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   99
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         Caption         =   "Schedule:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   108
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
         Caption         =   "Weight:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   107
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         Caption         =   "Wing Band:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   106
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         Caption         =   "Leg Band:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   105
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label31 
         BackColor       =   &H00404040&
         Caption         =   " Edit Entry"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   98
         Top             =   240
         Width           =   2775
      End
   End
   Begin MSAdodcLib.Adodc adodc_getter 
      Height          =   375
      Left            =   12840
      Top             =   9960
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data\fight_db.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data\fight_db.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from events"
      Caption         =   "getter"
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
   Begin VB.Frame frame_edit_participant 
      Height          =   5175
      Left            =   15600
      TabIndex        =   83
      Top             =   5640
      Visible         =   0   'False
      Width           =   6135
      Begin VB.CommandButton command_edit_participant_save 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Save"
         Height          =   495
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   91
         Top             =   4440
         Width           =   2055
      End
      Begin MSDataListLib.DataCombo datacombo_edit_participant_category 
         Bindings        =   "form_new_event.frx":0028
         Height          =   420
         Left            =   1920
         TabIndex        =   90
         Top             =   3840
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   741
         _Version        =   393216
         ListField       =   "participant_category"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox text_edit_participant_company 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   89
         Top             =   3120
         Width           =   4095
      End
      Begin VB.TextBox text_edit_participant_address 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   88
         Top             =   2400
         Width           =   4095
      End
      Begin VB.TextBox text_edit_participant_bet 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   87
         Top             =   1680
         Width           =   4095
      End
      Begin VB.TextBox text_edit_participant_name 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   86
         Top             =   960
         Width           =   4095
      End
      Begin VB.CommandButton command_edit_participant_close 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Close"
         Height          =   375
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   85
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         Caption         =   "Category:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   96
         Top             =   3840
         Width           =   1575
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         Caption         =   "Company:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   95
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         Caption         =   "Address:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   94
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         Caption         =   "Bet:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   93
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   92
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label25 
         BackColor       =   &H00404040&
         Caption         =   " Edit Participant"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   84
         Top             =   240
         Width           =   4455
      End
   End
   Begin VB.Frame frame_edit_event 
      Height          =   3975
      Left            =   12960
      TabIndex        =   71
      Top             =   600
      Visible         =   0   'False
      Width           =   5535
      Begin VB.CommandButton command_edit_event_close 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Close"
         Height          =   375
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton command_edit_event_save 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Save"
         Height          =   495
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   3360
         Width           =   2415
      End
      Begin VB.TextBox text_edit_event_minimum_bet 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   76
         Top             =   2760
         Width           =   3495
      End
      Begin MSComCtl2.DTPicker date_edit_event_schedule 
         Height          =   495
         Left            =   1920
         TabIndex        =   75
         Top             =   2040
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   78970881
         CurrentDate     =   42112
      End
      Begin MSDataListLib.DataCombo datacombo_edit_event_type 
         Bindings        =   "form_new_event.frx":0051
         Height          =   420
         Left            =   1920
         TabIndex        =   74
         Top             =   1440
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   741
         _Version        =   393216
         ListField       =   "event_type"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox text_edit_event_name 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   73
         Top             =   720
         Width           =   3495
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         Caption         =   "Min. Bet:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   80
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         Caption         =   "Date:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   79
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Caption         =   "Event Type:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   78
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   77
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label17 
         BackColor       =   &H00404040&
         Caption         =   " Edit Event"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   72
         Top             =   240
         Width           =   3855
      End
   End
   Begin MSAdodcLib.Adodc adodc_participant_no_match 
      Height          =   375
      Left            =   11040
      Top             =   10440
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data\fight_db.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data\fight_db.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select participant_name as no_match from participant_no_match_query order by participant_no_match_id desc"
      Caption         =   "participant_no_match"
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
   Begin VB.Frame frame_set_no_match 
      BackColor       =   &H00404040&
      Height          =   8415
      Left            =   18480
      TabIndex        =   58
      Top             =   360
      Visible         =   0   'False
      Width           =   615
      Begin MSAdodcLib.Adodc adodc_participant_no_match_union 
         Height          =   375
         Left            =   360
         Top             =   7800
         Visible         =   0   'False
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data\fight_db.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data\fight_db.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from participant_no_match_union_view_query order by id"
         Caption         =   "participant_no_match_union"
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
      Begin VB.CommandButton command_delete_no_match 
         Caption         =   "Delete"
         Height          =   420
         Left            =   3360
         TabIndex        =   66
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton command_add_no_match 
         Caption         =   "Add"
         Height          =   420
         Left            =   4800
         TabIndex        =   64
         Top             =   1680
         Width           =   1455
      End
      Begin MSDataListLib.DataCombo datacombo_frame_set_no_match 
         Bindings        =   "form_new_event.frx":0070
         Height          =   420
         Left            =   240
         TabIndex        =   63
         Top             =   1200
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   741
         _Version        =   393216
         ListField       =   "participant_name"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "form_new_event.frx":0091
         Height          =   5775
         Left            =   240
         TabIndex        =   62
         Top             =   2520
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   10186
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   24
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
               LCID            =   1033
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
               LCID            =   1033
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
      Begin VB.CommandButton command_frame_set_no_match_close 
         Caption         =   "Close"
         Height          =   375
         Left            =   5160
         TabIndex        =   61
         Top             =   240
         Width           =   1095
      End
      Begin VB.Timer timer_frame_set_no_match 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   4560
         Top             =   240
      End
      Begin VB.Label Label18 
         BackColor       =   &H00404040&
         Caption         =   "This participant won't be matched with the following:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   65
         Top             =   2160
         Width           =   5775
      End
      Begin VB.Label label_participant_no_match 
         BackColor       =   &H00404040&
         DataField       =   "participant_name"
         DataSource      =   "adodc_participants"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1920
         TabIndex        =   60
         Top             =   720
         Width           =   4335
      End
      Begin VB.Label Label16 
         BackColor       =   &H00404040&
         Caption         =   "Setting no match for:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   59
         Top             =   720
         Width           =   1815
      End
   End
   Begin MSAdodcLib.Adodc adodc_entry_schedule_type 
      Height          =   375
      Left            =   10200
      Top             =   9960
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data\fight_db.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data\fight_db.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from entry_schedule_type"
      Caption         =   "entry_schedule_type"
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
   Begin MSAdodcLib.Adodc adodc_participant_category 
      Height          =   375
      Left            =   3360
      Top             =   10440
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data\fight_db.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data\fight_db.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from participant_category_master"
      Caption         =   "participant_category"
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
   Begin MSAdodcLib.Adodc adodc_participants2 
      Height          =   375
      Left            =   8760
      Top             =   10440
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data\fight_db.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data\fight_db.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from participants_query"
      Caption         =   "participants2"
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
   Begin MSAdodcLib.Adodc adodc_entries2 
      Height          =   375
      Left            =   6480
      Top             =   10440
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data\fight_db.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data\fight_db.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from entries_query"
      Caption         =   "entries2"
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
   Begin MSAdodcLib.Adodc adodc_events2 
      Height          =   375
      Left            =   1080
      Top             =   10440
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data\fight_db.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data\fight_db.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from events_query order by event_id desc"
      Caption         =   "events2"
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
   Begin VB.Frame frame_slider 
      BackColor       =   &H00404040&
      Height          =   6255
      Left            =   7200
      TabIndex        =   50
      Top             =   2040
      Visible         =   0   'False
      Width           =   11775
      Begin VB.CommandButton command_participant_set_no_match 
         Caption         =   "Set No Match"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   57
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton command_frame_slider_delete 
         Caption         =   "Delete"
         Height          =   375
         Left            =   4080
         TabIndex        =   55
         Top             =   5280
         Width           =   1695
      End
      Begin VB.CommandButton command_frame_slider_save 
         Caption         =   "Save"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2280
         TabIndex        =   54
         Top             =   5280
         Width           =   1695
      End
      Begin VB.CommandButton command_frame_slider_edit 
         Caption         =   "Edit"
         Height          =   375
         Left            =   120
         TabIndex        =   53
         Top             =   5280
         Width           =   2055
      End
      Begin MSDataGridLib.DataGrid datagrid_frame_slider 
         Height          =   4455
         Left            =   120
         TabIndex        =   52
         Top             =   720
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   7858
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777215
         HeadLines       =   1
         RowHeight       =   24
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
               LCID            =   1033
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
               LCID            =   1033
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
      Begin VB.Timer timer_frame_slider 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   9840
         Top             =   240
      End
      Begin VB.CommandButton command_frame_slider_close 
         Caption         =   "Close"
         Height          =   375
         Left            =   10440
         TabIndex        =   51
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton command_entry_new 
      Caption         =   "New"
      Height          =   375
      Left            =   6480
      TabIndex        =   19
      Top             =   8760
      Width           =   735
   End
   Begin VB.Frame frame_entries 
      Caption         =   "Entries"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   1080
      TabIndex        =   41
      Top             =   6720
      Width           =   6495
      Begin MSDataListLib.DataCombo datacombo_entry_schedule 
         Bindings        =   "form_new_event.frx":00C0
         DataField       =   "entry_schedule_type"
         DataSource      =   "adodc_entries"
         Height          =   420
         Left            =   1800
         TabIndex        =   16
         Top             =   1560
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   741
         _Version        =   393216
         ListField       =   "entry_schedule_type"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton command_entry_save 
         Caption         =   "Save Entry"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4200
         TabIndex        =   18
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox text_entry_weight 
         DataField       =   "entry_weight"
         DataSource      =   "adodc_entries"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   17
         Top             =   2040
         Width           =   2295
      End
      Begin VB.TextBox text_entry_wing_band 
         DataField       =   "entry_wing_band"
         DataSource      =   "adodc_entries"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1800
         TabIndex        =   15
         Top             =   1080
         Width           =   4335
      End
      Begin VB.TextBox text_entry_leg_band 
         DataField       =   "entry_leg_band"
         DataSource      =   "adodc_entries"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1800
         TabIndex        =   14
         Top             =   600
         Width           =   4335
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Fight Schedule"
         Height          =   255
         Left            =   360
         TabIndex        =   56
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label label_participant_name 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         DataField       =   "participant_name"
         DataSource      =   "adodc_participants"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   840
         TabIndex        =   46
         Top             =   240
         Width           =   5535
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Weight"
         Height          =   255
         Left            =   720
         TabIndex        =   44
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Wing Band"
         Height          =   255
         Left            =   720
         TabIndex        =   43
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Leg Band"
         Height          =   255
         Left            =   720
         TabIndex        =   42
         Top             =   600
         Width           =   975
      End
      Begin VB.Label label_participant_id 
         BackColor       =   &H00404040&
         DataField       =   "participant_id"
         DataSource      =   "adodc_participants"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CommandButton command_new_participant 
      Caption         =   "New"
      Height          =   375
      Left            =   6480
      TabIndex        =   13
      Top             =   6120
      Width           =   735
   End
   Begin MSAdodcLib.Adodc adodc_entries 
      Height          =   375
      Left            =   5640
      Top             =   9960
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data\fight_db.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data\fight_db.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from entries_query"
      Caption         =   "entries"
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
   Begin MSAdodcLib.Adodc adodc_participants 
      Height          =   375
      Left            =   7920
      Top             =   9960
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data\fight_db.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data\fight_db.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from participants_query order by participant_id desc"
      Caption         =   "participants"
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
   Begin VB.Frame frame_overview 
      Caption         =   "Overview"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9495
      Left            =   7680
      TabIndex        =   30
      Top             =   360
      Width           =   10815
      Begin VB.CommandButton command_load_all_events 
         Caption         =   "All Events"
         Height          =   375
         Left            =   3600
         TabIndex        =   48
         Top             =   240
         Width           =   1335
      End
      Begin VB.Frame Frame8 
         Caption         =   "Entries"
         Height          =   2895
         Left            =   120
         TabIndex        =   35
         Top             =   6480
         Width           =   10575
         Begin MSDataGridLib.DataGrid datagrid_entries 
            Bindings        =   "form_new_event.frx":00E8
            Height          =   2535
            Left            =   120
            TabIndex        =   38
            Top             =   240
            Width           =   10335
            _ExtentX        =   18230
            _ExtentY        =   4471
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   0   'False
            HeadLines       =   1
            RowHeight       =   24
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   11
            BeginProperty Column00 
               DataField       =   "participant_name"
               Caption         =   "participant_name"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "entry_leg_band"
               Caption         =   "entry_leg_band"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column02 
               DataField       =   "entry_wing_band"
               Caption         =   "entry_wing_band"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   "entry_weight"
               Caption         =   "entry_weight"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   "entry_owner"
               Caption         =   "entry_owner"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column05 
               DataField       =   "participant_id"
               Caption         =   "participant_id"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column06 
               DataField       =   "entry_schedule_type"
               Caption         =   "entry_schedule_type"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column07 
               DataField       =   "entry_matching_status"
               Caption         =   "entry_matching_status"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column08 
               DataField       =   "participant_event"
               Caption         =   "participant_event"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column09 
               DataField       =   "participant_category"
               Caption         =   "participant_category"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column10 
               DataField       =   "participant_bet"
               Caption         =   "participant_bet"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  ColumnWidth     =   2775.118
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   2775.118
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   2775.118
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1454.74
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   1454.74
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   1454.74
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   2775.118
               EndProperty
               BeginProperty Column07 
                  ColumnWidth     =   2775.118
               EndProperty
               BeginProperty Column08 
                  ColumnWidth     =   1844.787
               EndProperty
               BeginProperty Column09 
                  ColumnWidth     =   2174.74
               EndProperty
               BeginProperty Column10 
                  ColumnWidth     =   2775.118
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Participants"
         Height          =   3855
         Left            =   120
         TabIndex        =   34
         Top             =   2640
         Width           =   10575
         Begin MSDataGridLib.DataGrid datagrid_participants 
            Bindings        =   "form_new_event.frx":0104
            Height          =   3495
            Left            =   120
            TabIndex        =   37
            Top             =   240
            Width           =   10335
            _ExtentX        =   18230
            _ExtentY        =   6165
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   0   'False
            HeadLines       =   1
            RowHeight       =   24
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   9
            BeginProperty Column00 
               DataField       =   "participant_name"
               Caption         =   "participant_name"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "participant_bet"
               Caption         =   "participant_bet"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column02 
               DataField       =   "participant_address"
               Caption         =   "participant_address"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   "participant_company"
               Caption         =   "participant_company"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   "participant_event"
               Caption         =   "participant_event"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column05 
               DataField       =   "event_name"
               Caption         =   "event_name"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column06 
               DataField       =   "event_id"
               Caption         =   "event_id"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column07 
               DataField       =   "participant_category"
               Caption         =   "participant_category"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column08 
               DataField       =   "participant_category_id"
               Caption         =   "participant_category_id"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  ColumnWidth     =   2775.118
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   2775.118
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   2775.118
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   2775.118
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   1844.787
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   2775.118
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   1454.74
               EndProperty
               BeginProperty Column07 
                  ColumnWidth     =   2775.118
               EndProperty
               BeginProperty Column08 
                  ColumnWidth     =   2489.953
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Events"
         Height          =   1935
         Left            =   120
         TabIndex        =   33
         Top             =   720
         Width           =   10575
         Begin MSDataGridLib.DataGrid datagrid_events 
            Bindings        =   "form_new_event.frx":0125
            Height          =   1455
            Left            =   120
            TabIndex        =   36
            Top             =   240
            Width           =   10335
            _ExtentX        =   18230
            _ExtentY        =   2566
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   -1  'True
            HeadLines       =   1
            RowHeight       =   24
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   6
            BeginProperty Column00 
               DataField       =   "event_name"
               Caption         =   "event_name"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "event_type"
               Caption         =   "event_type"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column02 
               DataField       =   "event_schedule"
               Caption         =   "event_schedule"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   "event_minimum_bet"
               Caption         =   "event_minimum_bet"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   "event_type_id"
               Caption         =   "event_type_id"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column05 
               DataField       =   "event_maximum_entries"
               Caption         =   "event_maximum_entries"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  ColumnWidth     =   2775.118
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   2775.118
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   2775.118
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   2775.118
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   1500.095
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   2594.835
               EndProperty
            EndProperty
         End
      End
      Begin MSComCtl2.DTPicker date_browser 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "M/d/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   31
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "m/dd/yyyy"
         Format          =   78970881
         CurrentDate     =   42089
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Browse by date:"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton command_new_event 
      Caption         =   "New"
      Height          =   375
      Left            =   6480
      TabIndex        =   6
      Top             =   2040
      Width           =   735
   End
   Begin MSAdodcLib.Adodc adodc_events 
      Height          =   330
      Left            =   1080
      Top             =   9960
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data\fight_db.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data\fight_db.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from events_query order by event_id desc"
      Caption         =   "events"
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
   Begin MSAdodcLib.Adodc adodc_event_type 
      Height          =   375
      Left            =   3360
      Top             =   9960
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data\fight_db.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data\fight_db.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from event_type_master_query"
      Caption         =   "event_type"
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
   Begin VB.Frame frame_participants 
      Caption         =   "Participants"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   1080
      TabIndex        =   21
      Top             =   2520
      Width           =   6495
      Begin VB.Frame Frame3 
         Caption         =   "Participant Details"
         Height          =   3615
         Left            =   120
         TabIndex        =   26
         Top             =   480
         Width           =   6255
         Begin VB.TextBox text_participant_address 
            DataField       =   "participant_address"
            DataSource      =   "adodc_participants"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1680
            TabIndex        =   10
            Top             =   2040
            Width           =   4335
         End
         Begin VB.TextBox text_participant_company 
            DataField       =   "participant_company"
            DataSource      =   "adodc_participants"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1680
            TabIndex        =   9
            Top             =   1440
            Width           =   4335
         End
         Begin VB.CommandButton command_participant_save 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Save Participant"
            Enabled         =   0   'False
            Height          =   375
            Left            =   3840
            TabIndex        =   12
            Top             =   3120
            Width           =   1335
         End
         Begin MSDataListLib.DataCombo datacombo_participant_category 
            Bindings        =   "form_new_event.frx":0140
            DataSource      =   "adodc_participants"
            Height          =   420
            Left            =   1680
            TabIndex        =   11
            Top             =   2640
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   741
            _Version        =   393216
            ListField       =   "participant_category"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.TextBox text_participant_bet 
            BackColor       =   &H00FFFFFF&
            DataField       =   "participant_bet"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """Php""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   13321
               SubFormatType   =   2
            EndProperty
            DataSource      =   "adodc_participants"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1680
            TabIndex        =   8
            Top             =   840
            Width           =   4335
         End
         Begin VB.TextBox text_participant_name 
            BackColor       =   &H00FFFFFF&
            DataField       =   "participant_name"
            DataSource      =   "adodc_participants"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1680
            TabIndex        =   7
            Top             =   240
            Width           =   4335
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Caption         =   "Address"
            Height          =   255
            Left            =   240
            TabIndex        =   40
            Top             =   2160
            Width           =   1335
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "Company"
            Height          =   255
            Left            =   240
            TabIndex        =   39
            Top             =   1560
            Width           =   1335
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "Category"
            Height          =   375
            Left            =   720
            TabIndex        =   29
            Top             =   2640
            Width           =   855
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Bet"
            Height          =   375
            Left            =   720
            TabIndex        =   28
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Name"
            Height          =   375
            Left            =   720
            TabIndex        =   27
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Label label_event_name 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         DataField       =   "event_name"
         DataSource      =   "adodc_events"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   720
         TabIndex        =   25
         Top             =   240
         Width           =   5655
      End
      Begin VB.Label label_event_id 
         BackColor       =   &H00404040&
         DataField       =   "event_id"
         DataSource      =   "adodc_events"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame frame_event_details 
      Caption         =   "Event Details"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   1080
      TabIndex        =   20
      Top             =   360
      Width           =   6495
      Begin VB.TextBox text_event_minimum_bet 
         DataField       =   "event_minimum_bet"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Php""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   13321
            SubFormatType   =   2
         EndProperty
         DataSource      =   "adodc_events"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1800
         TabIndex        =   3
         Top             =   1200
         Width           =   4335
      End
      Begin VB.TextBox text_event_name 
         DataField       =   "event_name"
         DataSource      =   "adodc_events"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   1
         Top             =   240
         Width           =   4335
      End
      Begin VB.CommandButton command_event_save 
         Caption         =   "Save"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4080
         TabIndex        =   5
         Top             =   1680
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker date_event_schedule 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "M/d/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         DataSource      =   "adodc_events"
         Height          =   375
         Left            =   1800
         TabIndex        =   4
         Top             =   1680
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   78970881
         CurrentDate     =   42088
      End
      Begin MSDataListLib.DataCombo datacombo_event_type 
         Bindings        =   "form_new_event.frx":0169
         DataField       =   "event_type"
         DataSource      =   "adodc_events"
         Height          =   420
         Left            =   1800
         TabIndex        =   2
         Top             =   720
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   741
         _Version        =   393216
         ListField       =   "event_type"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "Min. Bet"
         Height          =   255
         Left            =   840
         TabIndex        =   45
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Schedule"
         Height          =   255
         Left            =   840
         TabIndex        =   24
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Event Type"
         Height          =   255
         Left            =   720
         TabIndex        =   23
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Event Name"
         Height          =   255
         Left            =   720
         TabIndex        =   22
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00404040&
      Caption         =   "Total Entries:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   15120
      TabIndex        =   70
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00404040&
      Caption         =   "Total Participants:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   9360
      TabIndex        =   69
      Top             =   0
      Width           =   3375
   End
   Begin VB.Label label_total_entries 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   16800
      TabIndex        =   68
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label label_total_participants 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   12720
      TabIndex        =   67
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      DataField       =   "event_name"
      DataSource      =   "adodc_events"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   0
      Width           =   8295
   End
End
Attribute VB_Name = "form_new_event"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim frame_slider_datagrid_data_source As String
Dim selected_entry_schedule_type_id As Integer
Dim exiter As Boolean

Sub adodc_participant_no_match_union_default()
    With adodc_participant_no_match_union
        .RecordSource = "SELECT no_match_name from participant_no_match_union_view_query where participant_id = " & adodc_participants.Recordset("participant_id") & " order by id"
        .Refresh
    End With
End Sub

Sub update_total_labels()
    With adodc_participants
        .RecordSource = "select * from participants_query where participant_event = " & adodc_events.Recordset("event_id") & ""
        .Refresh
        label_total_participants.Caption = .Recordset.RecordCount
    End With
    
    With adodc_entries
        .RecordSource = "select * from entries_query where participant_event = " & adodc_events.Recordset("event_id") & ""
        .Refresh
        label_total_entries.Caption = .Recordset.RecordCount
    End With
    
    Call reset_recordsets
End Sub

Sub reset_recordsets()
    With adodc_participants
        .RecordSource = "select * from participants_query where participant_event = " & adodc_events.Recordset("event_id") & " order by participant_id desc"
        .Refresh
    End With
    
    With adodc_entries
        If adodc_participants.Recordset.EOF = False Then
            .RecordSource = "select * from entries_query where participant_id = " & adodc_participants.Recordset("participant_id") & " order by entry_id desc"
            .Refresh
        End If
    End With
End Sub

Sub entry_limit_check()
    'check if the current participant already have exceeded the maximum entries for this event
    If adodc_entries.Recordset.RecordCount >= adodc_events.Recordset("event_maximum_entries") Then
        If MsgBox("Maximum entries reached! Do you want to continue adding?", vbYesNo, "System Message") = vbYes Then
            exiter = False
        Else
            exiter = True
        End If
    End If
End Sub

Sub duplicate_participant_name_check()
    'check if this participants name entered aleady exist in this event
    With adodc_participants
        .RecordSource = "Select * from participants where participant_name = '" & text_participant_name.Text & "' and participant_event = " & adodc_events.Recordset("event_id") & ""
        .Refresh
        
        If .Recordset.RecordCount > 0 Then
            MsgBox "Participant name already exist in this event!", vbCritical, "System Message"
            
            exiter = True
        Else
            exiter = False
        End If
        
        .RecordSource = "select * from participants_query where participant_event = " & adodc_events.Recordset("event_id") & " order by participant_id desc"
        .Refresh
        
    End With
End Sub

Sub load_defaults()
    If adodc_events.Recordset.RecordCount <> 0 Then
        adodc_event_type.RecordSource = "SELECT * from event_type_master"
        adodc_event_type.Refresh
        
        adodc_participant_category.RecordSource = "Select participant_category FROM participant_category_master"
        adodc_participant_category.Refresh
        'adodc_events.RecordSource = "select * from events_query order by event_id desc"
        'adodc_events.Refresh
        'adodc_participants.RecordSource = "select * from participants_query"
        'adodc_participants.Refresh
        'adodc_entries.RecordSource = "select * from entries_query"
        'adodc_entries.Refresh
        
        frame_event_details.Enabled = False
        frame_participants.Enabled = False
        
        'event details controls
        command_event_save.Enabled = False
        command_participant_save.Enabled = False
        
        If adodc_events.Recordset("event_schedule") <> "" Then
            text_event_name.DataField = "event_name"
            datacombo_event_type.DataField = "event_type"
            date_event_schedule.DataField = "event_schedule"
            text_event_minimum_bet.DataField = "event_minimum_bet"
        End If
        
        command_new_event.Caption = "New"
        
        'participants frame controls
        command_new_participant.Caption = "New"
        
        frame_participants.Enabled = False
        text_participant_name.DataField = "participant_name"
        text_participant_bet.DataField = "participant_bet"
        text_participant_company.DataField = "participant_company"
        text_participant_address.DataField = "participant_address"
        datacombo_participant_category.DataField = "participant_category"
        
        text_participant_name.BackColor = &H80000005
        text_participant_bet.BackColor = &H80000005
        
        'entries frame controls and defaults
        text_entry_leg_band.Text = ""
        text_entry_wing_band.Text = ""
        text_entry_weight.Text = ""
        datacombo_entry_schedule.Text = ""
        
        frame_entries.Enabled = False
        text_entry_leg_band.DataField = "entry_leg_band"
        text_entry_wing_band.DataField = "entry_wing_band"
        text_entry_weight.DataField = "entry_weight"
        datacombo_entry_schedule.DataField = "entry_schedule_type"
        command_entry_save.Enabled = False
        command_entry_new.Caption = "New"
        
        'frame_slider defaults
        frame_slider.Left = 18480
        frame_slider.Width = 495
        frame_slider.Visible = False
        timer_frame_slider.Enabled = False
        command_frame_slider_save.Enabled = False
    End If
End Sub

Private Sub command_add_no_match_Click()
    Dim selected_participant_no_match As String
    
    If Not datacombo_frame_set_no_match.Text = "" Then
        'check if this no match already exist for this participant in this event
        With adodc_participant_no_match
            .RecordSource = "select * from participant_no_match_query where participant_no_match.participant_id = " & adodc_participants.Recordset("participant_id") & " and participant_name = '" & datacombo_frame_set_no_match.Text & "'"
            .Refresh
            
            If .Recordset.RecordCount > 0 Then
                MsgBox "This participant is already no matched with this id.", vbOKOnly, "System Message"
                Call adodc_participant_no_match_default
                Exit Sub
            End If
        End With
    
        'get selected participant_no_match_id
        adodc_participant_no_match.RecordSource = "select * from participants where participant_name = '" & datacombo_frame_set_no_match.Text & "' and participant_event = " & adodc_events.Recordset!event_id & ""
        adodc_participant_no_match.Refresh
        If adodc_participant_no_match.Recordset.RecordCount > 1 Then
            MsgBox "Selected participant not found!", vbCritical, "System Message"
        Else
            'if own name selected
            If adodc_participant_no_match.Recordset("participant_id") = adodc_participants.Recordset("participant_id") Then
                Call adodc_participant_no_match_default
                Exit Sub
            End If
            
            selected_participant_no_match_id = adodc_participant_no_match.Recordset("participant_id")
                
            'requery the current participants no match table
            With adodc_participant_no_match
                .RecordSource = "select * from participant_no_match where participant_id = " & adodc_participants.Recordset("participant_id") & " and participant_event = " & adodc_events.Recordset!event_id & ""
                .Refresh
            End With
            
            With adodc_participant_no_match
                .Recordset.AddNew
                .Recordset("participant_id") = adodc_participants.Recordset("participant_id")
                .Recordset("participant_no_match") = selected_participant_no_match_id
                .Recordset("participant_event") = adodc_participants.Recordset("participant_event")
                .Recordset.Update
            End With
            
            Call adodc_participant_no_match_default
            Call adodc_participant_no_match_union_default
            
            MsgBox "Save successful!", vbOKOnly
        End If
    End If
    
    Call adodc_participant_no_match_default
End Sub

Sub adodc_participant_no_match_default()
    If adodc_participant_no_match.Recordset.RecordCount <> 0 Then
        With adodc_participant_no_match
            .RecordSource = "select participant_name as no_match from participant_no_match_query where participant_no_match.participant_id = " & adodc_participants.Recordset("participant_id") & " order by participant_no_match_id desc"
            .Refresh
        End With
    End If
End Sub

Private Sub command_delete_no_match_Click()
    If adodc_participant_no_match.Recordset.EOF = False Then
        If MsgBox("Are you sure you want to delete this record?", vbYesNo, "System Message") = vbYes Then
            With adodc_participant_no_match
                .RecordSource = "select * from participant_no_match_query where participant_name = '" & .Recordset("no_match") & "'"
                .Refresh
                .RecordSource = "select * from participant_no_match_query where participants.participant_id = " & .Recordset("participants.participant_id") & ""
                .Refresh
                .RecordSource = "select * from participant_no_match where participant_no_match = " & .Recordset("participants.participant_id") & " and participant_event = " & adodc_events.Recordset("event_id") & " and participant_id = " & adodc_participants.Recordset("participant_id") & ""
                .Refresh
                .Recordset.Delete
                .Refresh
                
                Call adodc_participant_no_match_default
                Call adodc_participant_no_match_union_default
                
                MsgBox "Delete successful!", vbOKOnly
            End With
        End If
    End If
End Sub

Private Sub command_edit_entry_close_Click()
    frame_edit_entry.Visible = False
    frame_slider.Enabled = True
    frame_overview.Enabled = True
End Sub

Private Sub command_edit_entry_save_Click()
    'ask for confirmation
    If MsgBox("Are you sure you want to save this record?", vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    'get the schedule_type_id of the chosen entry_schedule_type
    With adodc_getter
        .RecordSource = "Select * from entry_schedule_type where entry_schedule_type = '" & datacombo_edit_entry_schedule_type & "'"
        .Refresh
        
        If .Recordset.RecordCount < 1 Then
            MsgBox "Schedule type not found!", vbCritical
            Exit Sub
        End If
    End With
    
    'start saving
    With adodc_entries.Recordset
        !entry_leg_band = text_edit_entry_leg_band.Text
        !entry_wing_band = text_edit_entry_wing_band.Text
        !entry_weight = text_edit_entry_weight.Text
        !entry_schedule_type_id = adodc_getter.Recordset!entry_schedule_type_id
        On Error GoTo edit_entry_error
        .Update
        
        frame_edit_entry.Visible = False
        frame_slider.Enabled = True
        frame_overview.Enabled = True
        MsgBox "Edit entry successful!", vbInformation
        .Requery
        Exit Sub
        
edit_entry_error:
        MsgBox "Edit entry error!", vbCritical
        Exit Sub
    End With
End Sub

Private Sub command_edit_event_close_Click()
    frame_edit_event.Visible = False
    frame_slider.Enabled = True
    frame_overview.Enabled = True
End Sub

Private Sub command_edit_event_save_Click()
    'ask for confirmation
    If MsgBox("Are you sure you want to edit this data?", vbYesNo) = vbNo Then
        Exit Sub
    End If

    With adodc_events
        'get the id of the chosen event_type
        adodc_getter.RecordSource = "select * from event_type_master where event_type = '" & datacombo_edit_event_type.Text & "'"
        adodc_getter.Refresh
        
        If adodc_getter.Recordset.RecordCount < 1 Then
            MsgBox "Event type not found! (Ln:240)", vbCritical
            GoTo edit_event_error
        End If
        
        .Recordset!event_name = text_edit_event_name.Text
        .Recordset!event_type_id = adodc_getter.Recordset!event_type_id
        .Recordset!event_schedule = date_edit_event_schedule.Value
        .Recordset!event_minimum_bet = text_edit_event_minimum_bet.Text
        
        On Error GoTo edit_event_error
        .Recordset.Update
        .Recordset.Requery
        frame_slider.Enabled = True
        frame_edit_event.Visible = False
        frame_overview.Enabled = True
        MsgBox "Edit successful!", vbInformation
        Exit Sub
        
edit_event_error:
        MsgBox "Error editing event! (Ln:237)", vbCritical
    End With
End Sub

Private Sub command_edit_participant_close_Click()
    frame_slider.Enabled = True
    frame_edit_participant.Visible = False
    frame_overview.Enabled = True
End Sub

Private Sub command_edit_participant_save_Click()
    'ask for confirmation
    If MsgBox("Are you sure you want to save edit this data?", vbYesNo) = vbNo Then
        Exit Sub
    End If

    'get the category_id of the chosen category
    adodc_getter.RecordSource = "select * from participant_category_master where participant_category = '" & datacombo_edit_participant_category.Text & "'"
    adodc_getter.Refresh
    If adodc_getter.Recordset.RecordCount < 1 Then
        MsgBox "Participant category not found!", vbCritical
        GoTo edit_participant_error
    End If
    
    With adodc_participants.Recordset
        !participant_category_id = adodc_getter.Recordset!participant_category_id
        !participant_name = text_edit_participant_name.Text
        !participant_bet = text_edit_participant_bet.Text
        !participant_address = text_edit_participant_address.Text
        !participant_company = text_edit_participant_company.Text
        '!participant_category = datacombo_edit_participant_category.Text
        
        On Error GoTo edit_participant_error
        .Update
        .Requery
        frame_slider.Enabled = True
        frame_edit_participant.Visible = False
        frame_overview.Enabled = True
        MsgBox "Participant edit successful!", vbInformation
        Exit Sub
        
edit_participant_error:
        MsgBox "Error editing participant! (Ln:263)", vbCritical
    End With
End Sub

Private Sub command_entry_new_Click()
    exiter = False
    If command_entry_new.Caption = "New" Then
        Call entry_limit_check
    End If
    
    If exiter = True Then
        Exit Sub
    Else
        If command_entry_new.Caption = "New" And adodc_participants.Recordset.RecordCount > 0 Then
            Call load_defaults
            frame_entries.Enabled = True
            command_entry_save.Enabled = True
            text_entry_leg_band.DataField = ""
            text_entry_wing_band.DataField = ""
            text_entry_weight.DataField = ""
            datacombo_entry_schedule.DataField = ""
            
            text_entry_leg_band.Text = ""
            text_entry_wing_band.Text = ""
            text_entry_weight.Text = ""
            datacombo_entry_schedule.Text = ""
            
            command_entry_new.Caption = "Cancel"
            
            text_entry_leg_band.SetFocus
        ElseIf command_entry_new.Caption = "Cancel" Then
            Call load_defaults
        Else
            MsgBox "Please select an event participant first.", vbOKOnly, "Sytem Message"
        End If
    End If
End Sub







Private Sub command_entry_save_Click()
    Dim current_participant_id As Integer
    
    'check if this entry LB or WB already exist for this event
    adodc_entries.RecordSource = "select * from entries_query where (entry_leg_band <> '' and entry_leg_band = '" & text_entry_leg_band.Text & "') or (entry_wing_band <> '' and entry_wing_band = '" & text_entry_wing_band.Text & "')"
    adodc_entries.Refresh
    If adodc_entries.Recordset.RecordCount > 0 Then
        'this wing band or legband already exist!
        MsgBox "This wing band/ leg band already exist! Try again!", vbCritical
        adodc_entries.RecordSource = "SELECT * FROM entries_query WHERE entry_owner = " & adodc_participants.Recordset("participant_id") & ""
        adodc_entries.Refresh
        Call load_defaults
        Exit Sub
    End If
    
    Call get_entry_schedule_type_id
        
    current_participant_id = adodc_participants.Recordset("participant_id")
    
    If exiter = True Then
        adodc_entries.RecordSource = "SELECT * FROM entries_query WHERE entry_owner = " & adodc_participants.Recordset("participant_id") & ""
        adodc_entries.Refresh
        Call load_defaults
        Exit Sub
    Else
        If text_entry_leg_band.Text = "" And text_entry_wing_band.Text = "" And text_entry_weight.Text = "" Then
        Else
            adodc_entries.RecordSource = "select * from entries where entry_owner = " & adodc_participants.Recordset("participant_id") & ""
            adodc_entries.Refresh
            
            With adodc_entries.Recordset
                .AddNew
                !entry_owner = adodc_participants.Recordset("participant_id")
                !entry_leg_band = text_entry_leg_band.Text
                !entry_wing_band = text_entry_wing_band.Text
                !entry_weight = text_entry_weight.Text
                !entry_schedule_type = selected_entry_schedule_type_id
                .Update
                
                adodc_entries.RecordSource = "SELECT * FROM entries_query WHERE entry_owner = " & adodc_participants.Recordset("participant_id") & " order by entry_id desc"
                adodc_entries.Refresh
                
                'Call update_total_labels
                'Call reset_recordsets
                Call load_defaults
            End With
        End If
    End If
End Sub

Sub get_entry_schedule_type_id()
    'get selected entry_schedule_type_id
    adodc_entry_schedule_type.RecordSource = "select entry_schedule_type_id from entry_schedule_type where entry_schedule_type = '" & datacombo_entry_schedule.Text & "'"
    adodc_entry_schedule_type.Refresh
    
    If adodc_entry_schedule_type.Recordset.RecordCount <> 1 Then
        MsgBox "Preferred entry schedule error.", vbCritical
        exiter = True
    Else
        selected_entry_schedule_type_id = adodc_entry_schedule_type.Recordset("entry_schedule_type_id")
        'load default recordset for entry_schedule_type
        adodc_entry_schedule_type.RecordSource = "select * from entry_schedule_type"
        adodc_entry_schedule_type.Refresh
    End If
End Sub

Private Sub command_event_save_Click()
    If command_event_save.Caption = "Save" Then
        If Not text_event_name.Text = "" Or datacombo_event_type.Text = "" Then
            'get the id of the chosen event_type
save_event:
            Dim event_type_id As Integer
            adodc_event_type.RecordSource = "SELECT event_type_id FROM event_type_master_query WHERE event_type LIKE '%" & datacombo_event_type.Text & "%'"
            adodc_event_type.Refresh
            
            If adodc_event_type.Recordset.RecordCount > 0 Then
                adodc_events.RecordSource = "SELECT * FROM events order by event_id desc"
                adodc_events.Refresh
                
                adodc_events.Recordset.AddNew
                
                adodc_events.Recordset("event_name") = text_event_name.Text
                adodc_events.Recordset("event_type") = adodc_event_type.Recordset("event_type_id")
                adodc_events.Recordset("event_minimum_bet") = text_event_minimum_bet.Text
                adodc_events.Recordset("event_schedule") = date_event_schedule.Value
                
                adodc_events.Recordset.Update
                
                adodc_events.RecordSource = "SELECT * FROM events_query order by event_id desc"
                adodc_events.Refresh
                'update the event name label in the add participant frame
                label_event_name.Caption = text_event_name.Text
                Call load_defaults
            Else
                If MsgBox("Do you want to add this new event type?", vbYesNo) = vbYes Then
                    Dim new_event_type_maximum_entries As Integer
enter_maximum_entries:
                    On Error Resume Next
                    new_event_type_maximum_entries = InputBox("Maximum entries per participant:", "Maximum Entries", "0")
                    If IsNumeric(new_event_type_maximum_entries) And new_event_type_maximum_entries > 0 Then
                        'save to database
                        adodc_event_type.RecordSource = "select * from event_type_master_query"
                        adodc_event_type.Refresh
                        adodc_event_type.Recordset.AddNew
                        adodc_event_type.Recordset!event_type = datacombo_event_type.Text
                        adodc_event_type.Recordset!event_maximum_entries = new_event_type_maximum_entries
                        adodc_event_type.Recordset.Update
                        GoTo save_event
                    ElseIf Not IsNumeric(new_event_type_maximum_entries) Then
                        MsgBox "Value entered not allowed!", vbCritical
                        GoTo enter_maximum_entries
                    End If
                Else
                    
                End If
            End If
        End If
    End If
    
    adodc_event_type.RecordSource = "select * from event_type_master_query"
    adodc_event_type.Refresh
    
End Sub



Private Sub command_frame_slider_close_Click()
    frame_slider.Left = 18480
    frame_slider.Width = 495
    frame_slider.Visible = False
    timer_frame_slider.Enabled = False
    command_participant_set_no_match.Enabled = False
End Sub

Private Sub command_frame_slider_delete_Click()
    Dim delete_response As String
    delete_response = MsgBox("Are you sure you want to delete this record?", vbYesNo, "System Message")
    If delete_response = vbYes Then
        If frame_slider_datagrid_data_source = "adodc_events" Then
            If adodc_events.Recordset.RecordCount <> 0 Then
                adodc_events2.Refresh
                adodc_events2.RecordSource = "Select * from events where event_id = " & adodc_events.Recordset("event_id") & ""
                adodc_events2.Refresh
                adodc_events2.Recordset.Delete
                adodc_events2.Refresh
                
                
                adodc_events.Refresh
                
                adodc_participants.Recordset.Requery
                adodc_entries.Recordset.Requery
            End If
        ElseIf frame_slider_datagrid_data_source = "adodc_participants" Then
            If adodc_participants.Recordset.RecordCount <> 0 Then
                adodc_participants2.Refresh
                adodc_participants2.RecordSource = "SELECT * from participants where participant_id = " & adodc_participants.Recordset("participant_id") & ""
                adodc_participants2.Refresh
                adodc_participants2.Recordset.Delete
                adodc_participants2.Refresh
                
                adodc_participants.RecordSource = "SELECT * from participants_query where participant_event = " & adodc_events.Recordset("event_id") & " order by participant_id desc"
                adodc_participants.Refresh
                
                adodc_entries.Recordset.Requery
            End If
        ElseIf frame_slider_datagrid_data_source = "adodc_entries" Then
            If adodc_entries.Recordset.RecordCount <> 0 Then
                adodc_entries2.Refresh
                adodc_entries2.RecordSource = "select * from entries where entry_id = " & adodc_entries.Recordset("entry_id") & ""
                adodc_entries2.Refresh
                adodc_entries2.Recordset.Delete
                adodc_entries2.Refresh
                adodc_entries.Refresh
            End If
        End If
    End If
    
    'Call update_total_labels
    'Call load_defaults
End Sub

Private Sub command_frame_slider_edit_Click()
    frame_slider.Enabled = False
    frame_overview.Enabled = False
    
    If command_frame_slider_edit.Caption = "Edit Event" Then
        If adodc_events.Recordset.RecordCount = 0 Then
            MsgBox "No record found!", vbInformation
            frame_slider.Enabled = True
            frame_overview.Enabled = True
        Exit Sub
        End If
        With frame_edit_event
            .Height = 3975
            .Left = 11040
            .Top = 2160
            .Width = 5535
            .Visible = True
        End With
        
        text_edit_event_name.Text = adodc_events.Recordset!event_name
        datacombo_edit_event_type.Text = adodc_events.Recordset!event_type
        date_edit_event_schedule.Value = adodc_events.Recordset!event_schedule
        text_edit_event_minimum_bet.Text = adodc_events.Recordset!event_minimum_bet
    ElseIf command_frame_slider_edit.Caption = "Edit Participant" Then
        If adodc_participants.Recordset.RecordCount = 0 Then
            MsgBox "No record found!", vbInformation
            frame_slider.Enabled = True
            frame_overview.Enabled = True
            Exit Sub
        End If
        With frame_edit_participant
            .Height = 5175
            .Left = 9120
            .Top = 1080
            .Width = 6135
            .Visible = True
        End With
        
        With adodc_participants
            text_edit_participant_name.Text = .Recordset!participant_name
            text_edit_participant_bet = .Recordset!participant_bet
            text_edit_participant_address = .Recordset!participant_address
            text_edit_participant_company = .Recordset!participant_company
            datacombo_edit_participant_category = .Recordset!participant_category
        End With
    ElseIf command_frame_slider_edit.Caption = "Edit Entry" Then
        If adodc_entries.Recordset.RecordCount = 0 Then
            MsgBox "No record found!", vbInformation
            frame_slider.Enabled = True
            frame_overview.Enabled = True
            Exit Sub
        End If
        With frame_edit_entry
            .Height = 4215
            .Left = 10440
            .Top = 1920
            .Width = 4575
            .Visible = True
        End With
        
        With adodc_entries
            text_edit_entry_leg_band.Text = .Recordset!entry_leg_band
            text_edit_entry_wing_band.Text = .Recordset!entry_wing_band
            text_edit_entry_weight.Text = .Recordset!entry_weight
            datacombo_edit_entry_schedule_type.Text = .Recordset!entry_schedule_type
        End With
    End If
End Sub

Private Sub command_frame_slider_save_Click()
    If frame_slider_datagrid_data_source = "adodc_events" Then
        If adodc_events.Recordset.RecordCount <> 0 Then
            adodc_events.Recordset.Update
            adodc_events.Recordset.Requery
            adodc_events.Refresh
            
            adodc_participants.Recordset.Requery
            adodc_entries.Recordset.Requery
        End If
    ElseIf frame_slider_datagrid_data_source = "adodc_participants" Then
        If adodc_participants.Recordset.RecordCount <> 0 Then
            adodc_participants.Recordset.Update
            adodc_participants.Recordset.Requery
            adodc_participants.Refresh
            
            adodc_entries.Recordset.Requery
        End If
    ElseIf frame_slider_datagrid_data_source = "adodc_entries" Then
        If adodc_participants.Recordset.RecordCount <> 0 Then
            adodc_entries.Recordset.Update
            adodc_entries.Recordset.Requery
            adodc_entries.Refresh
        End If
    End If
End Sub

Private Sub command_load_all_events_Click()
    adodc_events.RecordSource = "SELECT * FROM events_query where event_name <> 'None' order by event_id desc"
    adodc_events.Refresh
End Sub

Private Sub command_new_event_Click()
    If command_new_event.Caption = "New" Then
        Call load_defaults
        text_event_name.DataField = ""
        datacombo_event_type.DataField = ""
        date_event_schedule.DataField = ""
        text_event_minimum_bet.DataField = ""
        
        text_event_name.Text = ""
        datacombo_event_type.Text = ""
        text_event_minimum_bet.Text = ""
        
        frame_event_details.Enabled = True
        command_new_event.Caption = "Cancel"
        command_event_save.Caption = "Save"
        command_event_save.Enabled = True
        
        date_event_schedule.Value = Date
        
        text_event_name.SetFocus
    ElseIf command_new_event.Caption = "Cancel" Then
        command_new_event.Caption = "New"
        Call load_defaults
    End If
End Sub

Private Sub command_new_participant_Click()
    If adodc_events.Recordset.EOF = True Then
        MsgBox "Please add an event first.", vbOKOnly, "System Message"
    Else
        If command_new_participant.Caption = "New" Then
            Call load_defaults
            frame_participants.Enabled = True
            text_participant_name.DataField = ""
            text_participant_bet.DataField = ""
            text_participant_company.DataField = ""
            text_participant_address.DataField = ""
            datacombo_participant_category.DataField = ""
            
            text_participant_name.Text = ""
            text_participant_bet.Text = ""
            text_participant_company.Text = ""
            text_participant_address.Text = ""
            datacombo_participant_category.Text = ""
            
            command_participant_save.Enabled = True
            command_new_participant.Caption = "Cancel"
            
            text_participant_name.SetFocus
        ElseIf command_new_participant.Caption = "Cancel" Then
            command_new_participant.Caption = "New"
            Call load_defaults
        End If
    End If
End Sub

Private Sub command_participant_save_Click()
    'get the chosen dont match's id
    Dim category As Integer
    adodc_participant_category.RecordSource = "Select participant_category_id FROM participant_category_master where participant_category = '" & datacombo_participant_category.Text & "'"
    adodc_participant_category.Refresh
    
    Call duplicate_participant_name_check
    
    If exiter = True Then
        MsgBox "Please solve the issues before saving.", vbCritical
    Else
        If adodc_participant_category.Recordset.RecordCount <> 0 Or datacombo_participant_category.Text = "" Then
            If adodc_participant_category.Recordset.RecordCount <> 0 Then
                category = adodc_participant_category.Recordset("participant_category_id")
            End If
            
            adodc_participants.RecordSource = "Select * FROM participants WHERE participant_event like '" & adodc_events.Recordset("event_id") & "'"
            adodc_participants.Refresh
            If text_participant_name.Text = "" Or text_participant_bet.Text = "" Or text_participant_company.Text = "" Or datacombo_participant_category.Text = "" Then
                MsgBox "Please fill the form up.", vbCritical
            Else
                'adding new participant to event
                'get event id then add the participant to participants table
                With adodc_participants.Recordset
                    .AddNew
                    !participant_name = text_participant_name.Text
                    !participant_bet = text_participant_bet.Text
                    !participant_event = adodc_events.Recordset("event_id")
                    !participant_address = text_participant_address.Text
                    !participant_company = text_participant_company.Text
                    !participant_category = category
                    .Update
                    
                    adodc_participants.RecordSource = "Select * from participants_query where participant_event = " & adodc_events.Recordset("event_id") & " order by participant_id desc"
                    adodc_participants.Refresh
                    
                    Call update_total_labels
                    Call reset_recordsets
                    Call load_defaults
                End With
            End If
        Else
            MsgBox "Please solve the issues before saving.", vbCritical
        End If
    End If
End Sub



Private Sub datacombo_frame_set_no_match_Change()
    If adodc_participants.Recordset("participant_name") = datacombo_frame_set_no_match.Text Then
        MsgBox "Own ID selected!", vbCritical, "System Message"
    End If
End Sub


Private Sub datacombo_participant_category_Change()
'    If text_participant_name.Text <> "" And text_participant_bet.Text >= adodc_events.Recordset("event_minimum_bet") And text_participant_company.Text <> "" And datacombo_participant_category.Text <> "" Then
'        command_participant_save.Enabled = True
'    Else
'       MsgBox "Please make sure the data entered are correct.", vbOKOnly, "System Message"
'    End If
End Sub

Private Sub datagrid_entries_DblClick()
    Call frame_slider_on
    Set datagrid_frame_slider.DataSource = adodc_entries
    frame_slider_datagrid_data_source = "adodc_entries"
    
    command_frame_slider_edit.Caption = "Edit Entry"
End Sub

Private Sub datagrid_events_Click()
    If adodc_events.Recordset.RecordCount > 0 Then
        adodc_participants.RecordSource = "SELECT * FROM participants_query WHERE participant_event = " & adodc_events.Recordset("event_id") & " order by participant_id desc"
        adodc_participants.Refresh
        If adodc_participants.Recordset.RecordCount > 0 Then
            adodc_entries.RecordSource = "SELECT * FROM entries_query WHERE entry_owner = " & adodc_participants.Recordset("participant_id") & ""
            adodc_entries.Refresh
        Else
            adodc_entries.RecordSource = "SELECT * FROM entries_query WHERE entry_owner = null"
            adodc_entries.Refresh
        End If
    End If
End Sub



Private Sub datagrid_events_DblClick()
    Call update_total_labels
    Call frame_slider_on
    Set datagrid_frame_slider.DataSource = adodc_events
    frame_slider_datagrid_data_source = "adodc_events"
    
    command_frame_slider_edit.Caption = "Edit Event"
End Sub

Private Sub datagrid_participants_Click()
    command_participant_set_no_match.Enabled = True
    If adodc_participants.Recordset.RecordCount > 0 Then
        adodc_entries.RecordSource = "SELECT * FROM entries_query WHERE entry_owner = " & adodc_participants.Recordset("participant_id") & ""
        adodc_entries.Refresh
    Else
        adodc_entries.RecordSource = "SELECT * FROM entries_query WHERE entry_owner = null"
        adodc_entries.Refresh
    End If
End Sub

Private Sub datagrid_participants_DblClick()
    Call frame_slider_on
    Set datagrid_frame_slider.DataSource = adodc_participants
    frame_slider_datagrid_data_source = "adodc_participants"
    
    command_frame_slider_edit.Caption = "Edit Participant"
End Sub

Private Sub date_browser_Change()
    adodc_events.RecordSource = "SELECT * FROM events_query WHERE event_name AND event_schedule like '" & date_browser.Value & "'"
    adodc_events.Refresh
End Sub

Private Sub Form_Load()
    If adodc_events.Recordset.RecordCount <> 0 And adodc_events.Recordset.EOF = False Then
        If adodc_events.Recordset("event_schedule") <> 0 Then
            date_event_schedule.DataField = "event_schedule"
            date_browser.Value = Date
            Call load_defaults
        End If
    End If
End Sub

Private Sub label_event_id_Change()
    If label_event_id.Caption = "" Then
        adodc_participants.RecordSource = "SELECT * FROM events_query WHERE event_name = null"
        adodc_participants.Refresh
        adodc_entries.RecordSource = "SELECT * FROM entries_query WHERE entry_owner = null"
        adodc_entries.Refresh
    Else
        adodc_participants.RecordSource = "SELECT * FROM participants_query WHERE participant_event = " & adodc_events.Recordset("event_id") & " order by participant_id desc"
        adodc_participants.Refresh
        
        label_total_participants.Caption = adodc_participants.Recordset.RecordCount
        
        If adodc_participants.Recordset.RecordCount > 0 Then
            adodc_entries.RecordSource = "Select * from entries_query where participant_event = " & adodc_events.Recordset("event_id") & ""
            adodc_entries.Refresh
            label_total_entries.Caption = adodc_entries.Recordset.RecordCount
            
            adodc_entries.RecordSource = "SELECT * FROM entries_query WHERE entry_owner = " & adodc_participants.Recordset("participant_id") & ""
            adodc_entries.Refresh
        Else
            adodc_entries.RecordSource = "SELECT * FROM entries_query WHERE entry_owner = null"
            adodc_entries.Refresh
            label_total_entries.Caption = adodc_entries.Recordset.RecordCount
        End If
    End If
End Sub


Private Sub label_participant_id_Change()
    If label_participant_id.Caption = "" Then
        adodc_entries.RecordSource = "SELECT * FROM entries_query WHERE entry_owner = null"
        adodc_entries.Refresh
    Else
        If adodc_participants.Recordset.RecordCount > 0 Then
            adodc_entries.RecordSource = "SELECT * FROM entries_query WHERE entry_owner = " & adodc_participants.Recordset("participant_id") & ""
            adodc_entries.Refresh
        Else
            adodc_entries.RecordSource = "SELECT * FROM entries_query WHERE entry_owner = null"
            adodc_entries.Refresh
        End If
    End If
End Sub




Private Sub label_participant_no_match_Change()
    If adodc_participants.Recordset.RecordCount <> 0 And adodc_participants.Recordset.EOF = False And command_participant_save.Enabled = False Then
        adodc_participant_no_match.RecordSource = "select participant_name as no_match from participant_no_match_query where participant_no_match.participant_id = " & adodc_participants.Recordset("participant_id") & ""
        adodc_participant_no_match.Refresh
    End If
End Sub






Private Sub text_participant_bet_LostFocus()
    If text_participant_bet.DataField = "" And text_participant_bet.Text <> "" Then
        Dim participant_bet As Currency
        participant_bet = Val(text_participant_bet.Text)
        If participant_bet < adodc_events.Recordset("event_minimum_bet") Then
            MsgBox "Must be higher or equal to " & adodc_events.Recordset("event_minimum_bet") & "(Event's Minimum Bet).", vbCritical, "System Message"
            text_participant_bet.BackColor = &H8080FF
        Else
            text_participant_bet.BackColor = &H80000005
        End If
        text_participant_bet.Text = participant_bet
    End If
End Sub


Private Sub text_participant_name_LostFocus()
    Call duplicate_participant_name_check
    If exiter = True Then
        text_participant_name.BackColor = &H8080FF
        Exit Sub
    End If
    text_participant_name.BackColor = &H80000005
End Sub

Private Sub timer_frame_slider_Timer()
    If frame_slider.Left > 7200 Then
        frame_slider.Left = frame_slider.Left - 500
        If frame_slider.Width < 11775 Then
            frame_slider.Width = frame_slider.Width + 500
        End If
    Else
        timer_frame_slider.Enabled = False
    End If
End Sub

Sub frame_slider_on()
    frame_slider.Visible = True
    frame_slider.Top = 2040
    frame_slider.Left = 18360
    timer_frame_slider.Enabled = True
End Sub

Sub frame_set_no_match_default()
    timer_frame_set_no_match.Enabled = False
    frame_set_no_match.Visible = False
    With frame_set_no_match
        .Left = 18480   'left = 12600
        .Width = 615    'width = 6495
    End With
End Sub

Private Sub command_participant_set_no_match_Click()
    Call adodc_participant_no_match_default
    Call adodc_participant_no_match_union_default
    Call frame_set_no_match_default
    frame_set_no_match.Visible = True
    timer_frame_set_no_match.Enabled = True
End Sub


Private Sub timer_frame_set_no_match_Timer()
    With frame_set_no_match
        If .Left > 12600 Then
            .Left = .Left - 200
            If .Width < 6495 Then
                .Width = .Width + 200
            End If
        Else
            timer_frame_set_no_match.Enabled = False
        End If
    End With
End Sub

Private Sub command_frame_set_no_match_close_Click()
    Call frame_set_no_match_default
End Sub
