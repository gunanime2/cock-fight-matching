VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form form_matches_2 
   Caption         =   "Matches"
   ClientHeight    =   9225
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   20250
   Begin VB.Frame frame_edit_unmatched 
      BackColor       =   &H00C0C0C0&
      Height          =   2655
      Left            =   12000
      TabIndex        =   72
      Top             =   2640
      Visible         =   0   'False
      Width           =   8055
      Begin VB.CommandButton command_add_to_matches 
         Caption         =   "Add To Matches"
         Height          =   375
         Left            =   1080
         TabIndex        =   82
         Top             =   2160
         Width           =   2415
      End
      Begin VB.CommandButton command_close_edit_unmatched 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Close"
         Height          =   375
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Caption         =   "By clcking on the ""Add To Matches"" button this entry will be added to the matches as an entry for a new fight slot."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   4800
         TabIndex        =   83
         Top             =   1320
         Width           =   3135
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         Caption         =   "entry_weight"
         DataField       =   "entry_weight"
         DataSource      =   "adodc_unmatched"
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
         Left            =   3240
         TabIndex        =   81
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         Caption         =   "entry_wing_band"
         DataField       =   "entry_wing_band"
         DataSource      =   "adodc_unmatched"
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
         Left            =   1680
         TabIndex        =   80
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         Caption         =   "entry_leg_band"
         DataField       =   "entry_leg_band"
         DataSource      =   "adodc_unmatched"
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
         Left            =   240
         TabIndex        =   79
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Line Line12 
         X1              =   120
         X2              =   4560
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line11 
         X1              =   3120
         X2              =   3120
         Y1              =   1200
         Y2              =   2040
      End
      Begin VB.Line Line10 
         X1              =   1560
         X2              =   1560
         Y1              =   1200
         Y2              =   2040
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Weight"
         Height          =   255
         Left            =   3240
         TabIndex        =   78
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Wing Band"
         Height          =   255
         Left            =   1680
         TabIndex        =   77
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Leg Band"
         Height          =   255
         Left            =   120
         TabIndex        =   76
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label label21 
         Caption         =   "participant_name"
         DataField       =   "participant_name"
         DataSource      =   "adodc_unmatched"
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
         Left            =   120
         TabIndex        =   75
         Top             =   720
         Width           =   6135
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "Modifying Unmatched Entry"
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
         TabIndex        =   73
         Top             =   240
         Width           =   6135
      End
   End
   Begin VB.Timer timer_search 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   4080
      Top             =   10200
   End
   Begin VB.Frame frame 
      Caption         =   "Search"
      Height          =   855
      Left            =   240
      TabIndex        =   69
      Top             =   9240
      Width           =   4335
      Begin VB.CommandButton command_search 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Search"
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
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox text_search 
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
         Left            =   120
         TabIndex        =   70
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame frame_move_match 
      BackColor       =   &H00808080&
      Height          =   3015
      Left            =   9600
      TabIndex        =   63
      Top             =   5040
      Visible         =   0   'False
      Width           =   3975
      Begin VB.CommandButton command_move 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Move Match"
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   1440
         Width           =   3735
      End
      Begin MSDataListLib.DataCombo datacombo_move_match 
         Bindings        =   "form_matches_2.frx":0000
         Height          =   420
         Left            =   120
         TabIndex        =   66
         Top             =   840
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   741
         _Version        =   393216
         ListField       =   "match_schedule"
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
      Begin VB.CommandButton command_close_move_match 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Close"
         Height          =   375
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label19 
         BackColor       =   &H00404040&
         Caption         =   "Select the slot where you want to move this match then click the ""Move Match"" button."
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
         Height          =   975
         Left            =   120
         TabIndex        =   68
         Top             =   1920
         Width           =   3735
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "Move Match"
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
         TabIndex        =   64
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame frame_replace 
      BackColor       =   &H00808080&
      Height          =   7215
      Left            =   7440
      TabIndex        =   57
      Top             =   5640
      Visible         =   0   'False
      Width           =   11055
      Begin VB.CommandButton command_replace_select 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Select"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7920
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   6720
         Width           =   2895
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "form_matches_2.frx":001C
         Height          =   5895
         Left            =   120
         TabIndex        =   60
         Top             =   720
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   10398
         _Version        =   393216
         AllowUpdate     =   0   'False
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
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "participant_name"
            Caption         =   "ENTRYNAME"
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
            Caption         =   "LB"
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
            Caption         =   "WB"
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
            Caption         =   "WT"
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
            DataField       =   "entry_schedule_type"
            Caption         =   "SCHEDULE"
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
            DataField       =   "participant_category"
            Caption         =   "CATEGORY"
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
            DataField       =   "participant_bet"
            Caption         =   "BET"
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
               ColumnWidth     =   2385.071
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1604.976
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1454.74
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1739.906
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton command_frame_replace_close 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Close"
         Height          =   375
         Left            =   9720
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "Choose Unmatched Entry"
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
         TabIndex        =   58
         Top             =   240
         Width           =   9495
      End
   End
   Begin VB.Frame frame_edit_matches 
      BackColor       =   &H00808080&
      Height          =   2895
      Left            =   2760
      TabIndex        =   35
      Top             =   1680
      Visible         =   0   'False
      Width           =   8175
      Begin VB.CommandButton command_delete_match 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Delete Match"
         Height          =   375
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   2400
         Width           =   1575
      End
      Begin VB.CommandButton command_move_match 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Move Match"
         Height          =   375
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   2400
         Width           =   1575
      End
      Begin VB.CommandButton command_replace_puti 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Replace"
         Height          =   375
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   1800
         Width           =   1575
      End
      Begin VB.CommandButton command_replace_pula 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Replace"
         Height          =   375
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   1800
         Width           =   1695
      End
      Begin MSAdodcLib.Adodc adodc_edit_matches 
         Height          =   375
         Left            =   120
         Top             =   2400
         Visible         =   0   'False
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
         RecordSource    =   "select * from matches_view_query_proto where match_event_id = 0 "
         Caption         =   "edit_matches"
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
      Begin VB.CommandButton command_matches_edit_close 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Close"
         Height          =   375
         Left            =   6480
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   2400
         Width           =   1455
      End
      Begin VB.CommandButton command_puti_remove 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Remove"
         Height          =   375
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   1800
         Width           =   1935
      End
      Begin VB.CommandButton command_pula_remove 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Remove"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Line Line9 
         X1              =   120
         X2              =   7920
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line8 
         X1              =   120
         X2              =   7920
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Label label_puti_leg_band 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Puti LB"
         DataField       =   "puti_leg_band"
         DataSource      =   "adodc_edit_matches"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7080
         TabIndex        =   53
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label label_puti_wing_band 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Puti WB"
         DataField       =   "puti_wing_band"
         DataSource      =   "adodc_edit_matches"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5880
         TabIndex        =   52
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label label_puti_weight 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Puti WT"
         DataField       =   "puti_weight"
         DataSource      =   "adodc_edit_matches"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         TabIndex        =   51
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label label_pula_weight 
         BackStyle       =   0  'Transparent
         Caption         =   "Pula WT"
         DataField       =   "pula_weight"
         DataSource      =   "adodc_edit_matches"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   50
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label label_pula_wing_band 
         BackStyle       =   0  'Transparent
         Caption         =   "Pula WB"
         DataField       =   "pula_wing_band"
         DataSource      =   "adodc_edit_matches"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   49
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label label_pula_leg_band 
         BackStyle       =   0  'Transparent
         Caption         =   "Pula LB"
         DataField       =   "pula_leg_band"
         DataSource      =   "adodc_edit_matches"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   48
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label label_puti_address 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         DataField       =   "puti_address"
         DataSource      =   "adodc_edit_matches"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   47
         Top             =   480
         Width           =   3615
      End
      Begin VB.Label label_pula_address 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         DataField       =   "pula_address"
         DataSource      =   "adodc_edit_matches"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   480
         Width           =   3615
      End
      Begin VB.Line Line7 
         X1              =   120
         X2              =   7920
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line6 
         X1              =   6960
         X2              =   6960
         Y1              =   840
         Y2              =   1680
      End
      Begin VB.Line Line5 
         X1              =   5760
         X2              =   5760
         Y1              =   840
         Y2              =   1680
      End
      Begin VB.Line Line4 
         X1              =   2280
         X2              =   2280
         Y1              =   840
         Y2              =   1680
      End
      Begin VB.Line Line3 
         X1              =   1080
         X2              =   1080
         Y1              =   840
         Y2              =   1680
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Weight"
         DataSource      =   "adodc_edit_matches"
         Height          =   255
         Left            =   4320
         TabIndex        =   45
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Wing Band"
         DataSource      =   "adodc_edit_matches"
         Height          =   255
         Left            =   5760
         TabIndex        =   44
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Leg Band"
         DataSource      =   "adodc_edit_matches"
         Height          =   255
         Left            =   7080
         TabIndex        =   43
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Weight"
         DataSource      =   "adodc_edit_matches"
         Height          =   255
         Left            =   2640
         TabIndex        =   42
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Wing Band"
         DataSource      =   "adodc_edit_matches"
         Height          =   255
         Left            =   1200
         TabIndex        =   41
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Leg Band"
         DataSource      =   "adodc_edit_matches"
         Height          =   375
         Left            =   120
         TabIndex        =   40
         Top             =   960
         Width           =   855
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   7920
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line1 
         X1              =   4080
         X2              =   4080
         Y1              =   240
         Y2              =   2280
      End
      Begin VB.Label label_puti_entry_name 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Puti Entry Name"
         DataField       =   "puti_owner"
         DataSource      =   "adodc_edit_matches"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   39
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label label_pula_entry_name 
         BackStyle       =   0  'Transparent
         Caption         =   "Pula Entry Name"
         DataField       =   "pula_owner"
         DataSource      =   "adodc_edit_matches"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame frame_highlighter 
      Caption         =   "Highlighter"
      Height          =   855
      Left            =   4680
      TabIndex        =   32
      Top             =   9240
      Width           =   9735
      Begin MSAdodcLib.Adodc adodc_highlighter 
         Height          =   375
         Left            =   840
         Top             =   240
         Visible         =   0   'False
         Width           =   3495
         _ExtentX        =   6165
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
         RecordSource    =   "select * from matching_settings_summary_query"
         Caption         =   "highlighter"
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
      Begin MSDataListLib.DataCombo datacombo_highlighter 
         Bindings        =   "form_matches_2.frx":003A
         Height          =   420
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   741
         _Version        =   393216
         ListField       =   "criteria_summary"
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
      Begin VB.CommandButton command_highlighter 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Highlight"
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
         Left            =   7920
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Timer timer_bookmark 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   13920
      Top             =   10200
   End
   Begin VB.Frame frame_unmatched 
      Height          =   8655
      Left            =   14520
      TabIndex        =   29
      Top             =   480
      Width           =   5535
      Begin MSAdodcLib.Adodc adodc_unmatched 
         Height          =   375
         Left            =   240
         Top             =   7920
         Visible         =   0   'False
         Width           =   5055
         _ExtentX        =   8916
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
         RecordSource    =   "select * from entries_unmatched_query where participant_event = 0"
         Caption         =   "unmatched"
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
      Begin MSDataGridLib.DataGrid datagrid_unmatched 
         Bindings        =   "form_matches_2.frx":005A
         Height          =   7695
         Left            =   120
         TabIndex        =   31
         Top             =   720
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   13573
         _Version        =   393216
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
               ColumnWidth     =   1395.213
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1244.976
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1365.165
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   959.811
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1244.976
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1454.74
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   1739.906
            EndProperty
         EndProperty
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "Unmatched"
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
         TabIndex        =   30
         Top             =   240
         Width           =   5295
      End
   End
   Begin MSAdodcLib.Adodc adodc_matcher2 
      Height          =   330
      Left            =   4320
      Top             =   2280
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
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
      RecordSource    =   ""
      Caption         =   "matcher2"
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
   Begin MSAdodcLib.Adodc adodc_criteria 
      Height          =   330
      Left            =   4320
      Top             =   1920
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
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
      RecordSource    =   "select * from matching_settings_query_proto where criteria_summary not like '%manual%'"
      Caption         =   "criteria"
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
   Begin MSAdodcLib.Adodc adodc_entries 
      Height          =   330
      Left            =   4320
      Top             =   1560
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
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
      RecordSource    =   ""
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
   Begin MSAdodcLib.Adodc adodc_matcher 
      Height          =   330
      Left            =   600
      Top             =   2280
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
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
      RecordSource    =   ""
      Caption         =   "matcher"
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
   Begin MSAdodcLib.Adodc adodc_matches 
      Height          =   330
      Left            =   600
      Top             =   1920
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
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
      RecordSource    =   "select * from matches_view_query_proto where match_event_id = 0 "
      Caption         =   "matches"
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
   Begin MSAdodcLib.Adodc adodc_events 
      Height          =   330
      Left            =   600
      Top             =   1560
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
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
      RecordSource    =   "select * from events_query"
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
   Begin VB.Frame frame_matching_settings 
      Height          =   6615
      Left            =   240
      TabIndex        =   7
      Top             =   2040
      Visible         =   0   'False
      Width           =   14175
      Begin MSAdodcLib.Adodc adodc_matching_settings_master 
         Height          =   375
         Left            =   3720
         Top             =   5880
         Visible         =   0   'False
         Width           =   3255
         _ExtentX        =   5741
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
         RecordSource    =   "select * from matching_settings_master where criteria_name <> 'manual'"
         Caption         =   "matching_settings_master"
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
      Begin MSAdodcLib.Adodc adodc_matching_settings 
         Height          =   375
         Left            =   3720
         Top             =   5400
         Visible         =   0   'False
         Width           =   3255
         _ExtentX        =   5741
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
         RecordSource    =   "select * from matching_settings_query"
         Caption         =   "matching_settings"
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
      Begin VB.CommandButton command_settings_use_defaults 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Use Defaults"
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
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   5880
         Width           =   3015
      End
      Begin VB.CommandButton command_settings_delete 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Delete"
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
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   5280
         Width           =   1455
      End
      Begin VB.CommandButton command_settings_save 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Save"
         Enabled         =   0   'False
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
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   5280
         Width           =   1455
      End
      Begin VB.CommandButton command_settings_edit 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Edit"
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
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   4680
         Width           =   1455
      End
      Begin VB.CommandButton command_settings_add 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Add"
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
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   4680
         Width           =   1455
      End
      Begin MSDataListLib.DataCombo datacombo_sixth_criteria 
         Bindings        =   "form_matches_2.frx":0078
         DataField       =   "sixth_criteria"
         DataSource      =   "adodc_matching_settings"
         Height          =   420
         Left            =   480
         TabIndex        =   15
         Top             =   4200
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   741
         _Version        =   393216
         Enabled         =   0   'False
         ListField       =   "criteria_name"
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
      Begin MSDataListLib.DataCombo datacombo_fifth_criteria 
         Bindings        =   "form_matches_2.frx":00A5
         DataField       =   "fifth_criteria"
         DataSource      =   "adodc_matching_settings"
         Height          =   420
         Left            =   480
         TabIndex        =   14
         Top             =   3510
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   741
         _Version        =   393216
         Enabled         =   0   'False
         ListField       =   "criteria_name"
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
      Begin MSDataListLib.DataCombo datacombo_fourth_criteria 
         Bindings        =   "form_matches_2.frx":00D2
         DataField       =   "fourth_criteria"
         DataSource      =   "adodc_matching_settings"
         Height          =   420
         Left            =   480
         TabIndex        =   13
         Top             =   2760
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   741
         _Version        =   393216
         Enabled         =   0   'False
         ListField       =   "criteria_name"
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
      Begin MSDataListLib.DataCombo datacombo_third_criteria 
         Bindings        =   "form_matches_2.frx":00FF
         DataField       =   "third_criteria"
         DataSource      =   "adodc_matching_settings"
         Height          =   420
         Left            =   480
         TabIndex        =   12
         Top             =   2115
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   741
         _Version        =   393216
         Enabled         =   0   'False
         ListField       =   "criteria_name"
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
      Begin MSDataListLib.DataCombo datacombo_second_criteria 
         Bindings        =   "form_matches_2.frx":012C
         DataField       =   "second_criteria"
         DataSource      =   "adodc_matching_settings"
         Height          =   420
         Left            =   480
         TabIndex        =   11
         Top             =   1440
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   741
         _Version        =   393216
         Enabled         =   0   'False
         ListField       =   "criteria_name"
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
      Begin MSDataListLib.DataCombo datacombo_first_criteria 
         Bindings        =   "form_matches_2.frx":0159
         DataField       =   "first_criteria"
         DataSource      =   "adodc_matching_settings"
         Height          =   420
         Left            =   480
         TabIndex        =   10
         Top             =   720
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   741
         _Version        =   393216
         Enabled         =   0   'False
         ListField       =   "criteria_name"
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
      Begin MSDataGridLib.DataGrid datagrid_matching_settings 
         Bindings        =   "form_matches_2.frx":0186
         Height          =   5655
         Left            =   3600
         TabIndex        =   9
         Top             =   720
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   9975
         _Version        =   393216
         AllowUpdate     =   0   'False
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
            DataField       =   "first_criteria"
            Caption         =   "first_criteria"
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
            DataField       =   "second_criteria"
            Caption         =   "second_criteria"
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
            DataField       =   "third_criteria"
            Caption         =   "third_criteria"
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
            DataField       =   "fourth_criteria"
            Caption         =   "fourth_criteria"
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
            DataField       =   "fifth_criteria"
            Caption         =   "fifth_criteria"
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
            DataField       =   "sixth_criteria"
            Caption         =   "sixth_criteria"
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
               ColumnWidth     =   2775.118
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   2775.118
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton command_frame_settings_close 
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
         Left            =   12840
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "Matching Settings"
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
         Left            =   3600
         TabIndex        =   28
         Top             =   240
         Width           =   9135
      End
      Begin VB.Label Label8 
         Caption         =   "6th"
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   4200
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "5th"
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "4th"
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "3rd"
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "2nd"
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "1st"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "Criterias"
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
         TabIndex        =   21
         Top             =   240
         Width           =   3375
      End
   End
   Begin MSDataGridLib.DataGrid datagrid_matches 
      Bindings        =   "form_matches_2.frx":01AC
      Height          =   8055
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   14208
      _Version        =   393216
      AllowUpdate     =   0   'False
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
         DataField       =   "pula_owner"
         Caption         =   "ENTRYNAME"
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
         DataField       =   "pula_address"
         Caption         =   "ADDRESS"
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
         DataField       =   "pula_leg_band"
         Caption         =   "LB"
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
         DataField       =   "pula_wing_band"
         Caption         =   "WB"
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
         DataField       =   "pula_weight"
         Caption         =   "WEIGHT"
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
         DataField       =   "match_schedule"
         Caption         =   "VS"
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
         DataField       =   "puti_weight"
         Caption         =   "WEIGHT"
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
         DataField       =   "puti_wing_band"
         Caption         =   "WB"
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
         DataField       =   "puti_leg_band"
         Caption         =   "LB"
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
         DataField       =   "puti_address"
         Caption         =   "ADDRESS"
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
         DataField       =   "puti_owner"
         Caption         =   "ENTRYNAME"
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
            ColumnWidth     =   1604.976
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1395.213
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1049.953
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   945.071
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1110.047
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   780.095
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   959.811
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   959.811
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1395.213
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   2445.166
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton command_settings 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Settings"
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
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton command_print 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Print"
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
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton command_clear_matches 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Clear"
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
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton command_generate_matches 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Match"
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
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
   Begin MSDataListLib.DataCombo datacombo_event_name 
      Bindings        =   "form_matches_2.frx":01C8
      Height          =   420
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   741
      _Version        =   393216
      ListField       =   "event_name"
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Matches"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   450
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   19815
   End
End
Attribute VB_Name = "form_matches_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim search As String
Private sAppName As String, sAppPath As String

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub frame_matching_settings_default()
    With frame_matching_settings
        .Height = 8175
        .Left = 13560
        .Top = 960
        .Width = 855
        .Visible = False
    End With
End Sub

Sub check_datacombo_duplicate()
    If (datacombo_first_criteria.Text <> "" And datacombo_first_criteria.Text <> "none") And (datacombo_first_criteria.Text = datacombo_second_criteria.Text Or datacombo_first_criteria.Text = datacombo_third_criteria.Text Or datacombo_first_criteria.Text = datacombo_fourth_criteria.Text Or datacombo_first_criteria.Text = datacombo_fifth_criteria.Text Or datacombo_first_criteria.Text = datacombo_sixth_criteria.Text) Then
        MsgBox "Duplicate criteria in one set not allowed!", vbCritical
        'GoTo edit_error
    End If
End Sub

Sub controls_default()
    adodc_matching_settings.RecordSource = "select * from matching_settings_query where criteria_summary not like '%manual%'"
    adodc_matching_settings.Refresh
    
    If adodc_matching_settings.Recordset.RecordCount = 0 Then
        command_settings_add.Enabled = True
        command_settings_edit.Enabled = False
        command_settings_save.Enabled = False
        command_settings_delete.Enabled = False
        command_settings_use_defaults.Enabled = True
    Else
        command_settings_add.Enabled = True
        command_settings_edit.Enabled = True
        command_settings_save.Enabled = False
        command_settings_delete.Enabled = True
        command_settings_use_defaults.Enabled = True
        command_settings_add.Caption = "Add"
        command_settings_edit.Caption = "Edit"
    End If

    datacombo_first_criteria.Enabled = False
    datacombo_second_criteria.Enabled = False
    datacombo_third_criteria.Enabled = False
    datacombo_fourth_criteria.Enabled = False
    datacombo_fifth_criteria.Enabled = False
    datacombo_sixth_criteria.Enabled = False
    
    datacombo_first_criteria.DataField = "first_criteria"
    datacombo_second_criteria.DataField = "second_criteria"
    datacombo_third_criteria.DataField = "third_criteria"
    datacombo_fourth_criteria.DataField = "fourth_criteria"
    datacombo_fifth_criteria.DataField = "fifth_criteria"
    datacombo_sixth_criteria.DataField = "sixth_criteria"
    
    datagrid_matching_settings.Enabled = True
End Sub

Private Sub command_bookmark_Click()
    timer_bookmark.Enabled = True
End Sub

Private Sub command_add_to_matches_Click()
    Dim new_match_schedule As Integer
    If MsgBox("This entry will be added to a new fight slot. Are you sure you want to continue?", vbYesNo) = vbYes Then
        With adodc_matcher2
            .RecordSource = "select max(match_schedule) as match_schedule from matches where match_event_id = " & adodc_matches.Recordset!match_event_id & ""
            .Refresh
            new_match_schedule = .Recordset!match_schedule + 1
            
            .RecordSource = "select * from matches"
            .Refresh
            
            .Recordset.AddNew
            .Recordset!criteria_set_used = 42 'criteria used is manual
            .Recordset!match_event_id = adodc_matches.Recordset!match_event_id
            .Recordset!match_schedule = new_match_schedule
            .Recordset!entry_pula_id = adodc_unmatched.Recordset!entry_id
            'add opponent use the 'None' entry, the id is 171
            .Recordset!entry_puti_id = 171
            .Recordset.Update
            
            'now mark the entry's entry_matching_status as matched in the entries table
            .RecordSource = "select * from entries where entry_id = " & adodc_unmatched.Recordset!entry_id & ""
            .Refresh
            .Recordset!entry_matching_status = 2
            .Recordset.Update
        End With
        
        frame_edit_unmatched.Visible = False
        frame_unmatched.Enabled = True
        
        adodc_unmatched.Refresh
        adodc_matches.Refresh
        
        MsgBox "Entry succesfuly added to the matches.", vbInformation
    Else
        'do nothing
    End If
End Sub

Private Sub command_clear_matches_Click()
    'get event_name's event_id then mark 'Unmatched' all entries under it that are marked as 'Matched'
    If datacombo_event_name.Text <> "" Then
        adodc_matches.RecordSource = "select * from events where event_name = '" & datacombo_event_name.Text & "'"
        adodc_matches.Refresh
        If adodc_matches.Recordset.RecordCount = 1 Then
            adodc_matches.RecordSource = "select * from matches where match_event_id =" & adodc_matches.Recordset!event_id & ""
            adodc_matches.Refresh
            If adodc_matches.Recordset.RecordCount > 0 Then
                With adodc_matches
                    .Recordset.MoveFirst
                'delete the found matches
                    While Not .Recordset.EOF
                        .Recordset.Delete
                        .Recordset.MoveNext
                    Wend
                'get the entries that belongs to the current event
                    .RecordSource = "select * from events where event_name = '" & datacombo_event_name.Text & "'"
                    .Refresh
                    .RecordSource = "select * from entries_query_raw where participant_event = " & .Recordset!event_id & ""
                    .Refresh
                'mark the found entries as unmatched
                    .Recordset.MoveFirst
                    While Not .Recordset.EOF
                        If .Recordset!entry_matching_status = 2 Then
                            .Recordset!entry_matching_status = 1
                            .Recordset.Update
                            
                            adodc_unmatched.Refresh
                        End If
                        .Recordset.MoveNext
                    Wend
                    .RecordSource = "select * from matches_view_query_proto where event_name = '" & datacombo_event_name.Text & "'"
                    .Refresh
                    
                    MsgBox "Matches Cleared!"
                End With
            Else
                MsgBox "No matched record found!", vbInformation
                Exit Sub
            End If
        Else
            MsgBox "Event error!", vbInformation
            Exit Sub
        End If
    Else
        MsgBox "Choose an event first!", vbInformation
        Exit Sub
    End If
End Sub

Private Sub command_close_edit_unmatched_Click()
    frame_unmatched.Enabled = True
    frame_edit_unmatched.Visible = False
End Sub

Private Sub command_close_move_match_Click()
    frame_move_match.Visible = False
    frame_edit_matches.Enabled = True
End Sub

Private Sub command_delete_match_Click()
    If MsgBox("Are you sure you want to delete this match?", vbYesNo) = vbYes Then
        'first change the entry's entry_matching_status to 1(unmatched) in the entries table
        'then delete this match from the matches table
        'lastly resequence the match_schedule
        
        'first change the entry's entry_matching_status to 1(unmatched) in the entries table
        With adodc_matcher2
            'start with the pula
            .RecordSource = "select * from entries where entry_id = " & adodc_matches.Recordset!pula_id & ""
            .Refresh
            .Recordset!entry_matching_status = 1
            .Recordset.Update
            
            'then the puti
            .RecordSource = "select * from entries where entry_id = " & adodc_matches.Recordset!puti_id & ""
            .Refresh
            .Recordset!entry_matching_status = 1
            .Recordset.Update
        End With
        'adodc_matches.Refresh
        adodc_unmatched.Refresh
        
        'then delete this match from the matches table
        With adodc_matcher2
            .RecordSource = "select * from matches where match_id = " & adodc_matches.Recordset!match_id & ""
            .Refresh
            .Recordset.Delete
            .Refresh
        End With
        
        'lastly resequence the match_schedule
        With adodc_matcher2
            .RecordSource = "select * from matches where match_event_id = " & adodc_matches.Recordset!match_event_id & ""
            .Refresh
            .Recordset.MoveFirst
            
            Dim match_schedule As Integer
            match_schedule = 1
            
            While Not .Recordset.EOF
                .Recordset!match_schedule = match_schedule
                .Recordset.Update
                .Recordset.MoveNext
                match_schedule = match_schedule + 1
            Wend
        End With
        
        adodc_matcher2.Refresh
        adodc_matches.Refresh
        adodc_unmatched.Refresh
    
        MsgBox "Match Deleted!"
    Else
        'do nothing
    End If
End Sub

Private Sub command_frame_replace_close_Click()
    frame_edit_matches.Enabled = True
    frame_replace.Visible = False
End Sub

Private Sub command_frame_settings_close_Click()
    Call frame_matching_settings_default
End Sub

Private Sub command_generate_matches_Click()
    If datacombo_event_name.Text <> "" Then
        With adodc_matcher
            .RecordSource = "select * from events where event_name = '" & datacombo_event_name.Text & "'"
            .Refresh
            
            If .Recordset.RecordCount = 1 Then
            
                Dim event_id As Integer
                Dim event_name As String
                event_id = .Recordset!event_id
                event_name = .Recordset!event_name
                
                .RecordSource = "select * from entries_query where participant_event = " & event_id & " and entry_matching_status = 'Unmatched'"
                .Refresh
                
                If .Recordset.RecordCount > 1 Then
                    Dim pula_id As Integer
                    Dim puti_id As Integer
                    Dim puti_participant_id As Integer
                    Dim puti_weight_difference As Integer
                    Dim weight_difference As Integer
                    Dim puti_participant_bet As Currency
                    Dim puti_bet_difference As Currency
                    Dim bet_difference As Currency
                
                    Dim total_entries, total_possible_matches, total_daily_matches, total_event_days, _
                        event_day_segment, event_day_early, event_day_none, event_day_late, _
                        event_day_segment_counter, event_day_counter, event_match_counter, event_matche_number As Integer

                    Dim total_event_days_double As Double
                    
                    Dim event_fight_schedule, entry_weight_criteria_query, _
                        participant_category_criteria_query, entry_schedule_type_criteria_query, _
                        criteria_query As String
                    
                    Dim no_match_criteria, if_ready_criteria, entry_weight_criteria, _
                        entry_schedule_type_criteria, participant_category_criteria, _
                        participant_bet_criteria, user_ready As Boolean
                    
                        event_match_number = 1
                        
                        total_entries = .Recordset.RecordCount
                        total_possible_matches = total_entries / 2
                        total_daily_matches = 30 'this must be set by the user in the settings
                        
                        'get total event days
                        total_event_days_double = total_possible_matches / total_daily_matches
                        total_event_days = total_possible_matches / total_daily_matches
                        If total_event_days_double > total_event_days Then
                            total_event_days = total_event_days + 1
                        End If
                        
                        'code for splasher
                        'adodc_matcher.RecordSource = "Select * from process"
                        'adodc_matcher.Refresh
                        'adodc_matcher.Recordset!current_event_id = event_id
                        'adodc_matcher.Recordset!current_event_name = event_name
                        'adodc_matcher.Recordset!current_total = total_possible_matches
                        'adodc_matcher.Recordset.Update
                        'mdi_main.Visible = False
                        'res = Shell("splasher.exe " & sAppPath, vbHide)
                        
                        'get event day segments
                        event_day_segment = total_daily_matches / 3
                        
                        'declare default counter values
                        event_match_counter = 0
                        
                        'reset daily counters
reset_daily_counters:
                        event_day_segment_counter = 0
                        event_day_counter = 0
                        
counter_updated:
                        If Not event_match_counter > total_possible_matches Then
                            If Not event_day_segment_counter > total_daily_matches Then
                                If event_day_segment_counter <= event_day_segment Then
                                    event_fight_schedule = "Early Fight"
                                ElseIf (event_day_segment_counter > event_day_segment) And (event_day_segment_counter <= (event_day_segment * 2)) Then
                                    event_fight_schedule = "None"
                                ElseIf (event_day_segment_counter > (event_day_segment * 2)) And (event_day_segment_counter <= (event_day_segment * 3)) Then
                                    event_fight_schedule = "Late Fight"
                                Else
                                    GoTo counter_updated
                                End If
                                
                                'start the real matching
                                
                                'get entries
                                adodc_entries.RecordSource = "select * from entries_query " & _
                                    " where entry_matching_status = 'Unmatched'" & _
                                    " and participant_event = " & event_id & "" & _
                                    " and entry_schedule_type = '" & event_fight_schedule & "'" & _
                                    " order by participant_category"
                                adodc_entries.Refresh
                                
                                'get criterias
                                adodc_criteria.RecordSource = "select * from matching_settings_query " & _
                                    " where criteria_summary not like '%manual%' order by criteria_set_id"
                                adodc_criteria.Refresh
                                
                                If adodc_entries.Recordset.RecordCount > 0 Then
                                    If adodc_criteria.Recordset.RecordCount > 0 Then
entries_move_first:
                                        adodc_entries.Recordset.MoveFirst

get_matches:
                                        'translate criteria_set to query
                                        GoSub sub_translate_criteria_set
                                        
                                        'get matches using translated/generated query from
                                        'current criteria
                                        
                                        'check if this entry is already matched
                                        adodc_matcher2.RecordSource = "select * from entries where entry_id=" & adodc_entries.Recordset!entry_id & " and entry_matching_status = 2"
                                        adodc_matcher2.Refresh
                                        If adodc_matcher2.Recordset.RecordCount > 0 Then
                                            'YES, entry is already matched
                                            GoTo entries_eof_check
                                        End If
                                        
                                        'if the if_ready_criteria = true
                                        If if_ready_criteria = True Then
                                            'check if this entry_owner is ready
                                            adodc_matcher2.RecordSource = " select top 5 * from matches_view_query where match_event_id = " & adodc_events.Recordset!event_id & " order by match_id desc"
                                            adodc_matcher2.Refresh
                                            
                                            user_ready = True
                                            If adodc_matcher2.Recordset.RecordCount > 0 Then
                                                adodc_matcher2.Recordset.MoveFirst
                                                While Not adodc_matcher2.Recordset.EOF
                                                    If adodc_matcher2.Recordset("entries_query_for_pula.entry_owner") = adodc_entries.Recordset!entry_owner Or adodc_matcher2.Recordset("entries_query_for_puti.entry_owner") = adodc_entries.Recordset!entry_owner Then
                                                        user_ready = False
                                                        adodc_matcher2.Recordset.MoveNext
                                                    Else
                                                        adodc_matcher2.Recordset.MoveNext
                                                    End If
                                                Wend
                                            End If
                                            
                                            If user_ready = False Then
                                                'movenext entry
                                                GoTo entries_eof_check
                                            Else
                                                'continue
                                            End If
                                        End If

                                        .RecordSource = criteria_query
                                        'Sleep 500
                                        .Refresh
                                        
                                        
                                        If entry_weight_criteria = True Then
                                            .Recordset.Sort = "weight_difference"
                                        End If
                                        
                                        If .Recordset.RecordCount > 0 Then
                                            .Recordset.MoveFirst
                                            
                                            'get_no_match_criteria
get_no_match_criteria:
                                            If no_match_criteria = True Then
                                                'check_no_match
                                                adodc_matcher2.RecordSource = "Select * from participant_no_match_union_view_query " & _
                                                " where no_match_id = " & adodc_entries.Recordset!entry_owner & " " & _
                                                " and participant_id = " & .Recordset!entry_owner & "" & _
                                                " and event_id = " & event_id & ""
                                                adodc_matcher2.Refresh
                                                
                                                If adodc_matcher2.Recordset.RecordCount > 0 Then
                                                    'no match found, YES
                                                    'check if matches.eof
check_matcher_eof:
                                                    If .Recordset.EOF = True Then
                                                        GoTo entries_eof_check:
                                                    Else
                                                        'matcher.movenext
                                                        .Recordset.MoveNext
                                                        If .Recordset.EOF = False Then
                                                            GoTo get_no_match_criteria
                                                        End If
                                                        GoTo check_matcher_eof
                                                    End If
                                                Else
                                                    'no_match, NO
                                                    'get_if_ready_criteria
get_if_ready_criteria:
                                                    If if_ready_criteria = True Then
                                                        'check if matcher owner is ready
                                                        adodc_matcher2.RecordSource = " select top 5 * from matches_view_query where match_event_id = " & adodc_events.Recordset!event_id & " order by match_id desc"
                                                        adodc_matcher2.Refresh
                                                        
                                                        user_ready = True
                                                        If adodc_matcher2.Recordset.RecordCount > 0 Then
                                                            adodc_matcher2.Recordset.MoveFirst
                                                            While Not adodc_matcher2.Recordset.EOF
                                                                If adodc_matcher2.Recordset("entries_query_for_pula.entry_owner") = adodc_matcher.Recordset!entry_owner Or adodc_matcher2.Recordset("entries_query_for_puti.entry_owner") = .Recordset!entry_owner Then
                                                                    user_ready = False
                                                                    adodc_matcher2.Recordset.MoveNext
                                                                Else
                                                                    adodc_matcher2.Recordset.MoveNext
                                                                End If
                                                            Wend
                                                        End If
                                                        
                                                        If user_ready = True Then
                                                            'matcher is ready
                                                            'get entry_weight_criteria
get_entry_weight_criteria:
                                                            If entry_weight_criteria = True Then
                                                                'entry_weight_criteria = true, YES
                                                                'get allowed weight_difference from database
                                                                Dim allowed_weight_difference As Integer
                                                                
                                                                '20150422 code
                                                                'adodc_matcher2.RecordSource = "select * from settings_master where settings_name = 'allowed_weight_difference'"
                                                                'adodc_matcher2.Refresh
                                                                
                                                                adodc_matcher2.RecordSource = "select * from event_type_master where event_type_id = " & adodc_events.Recordset!event_type_id & ""
                                                                adodc_matcher2.Refresh
                                                                
                                                                If adodc_matcher2.Recordset.RecordCount <> 1 Then
                                                                    MsgBox "Settings master error!", vbCritical
                                                                Else
                                                                    'allowed_weight_difference = adodc_matcher2.Recordset!settings_value
                                                                    allowed_weight_difference = adodc_matcher2.Recordset!event_allowed_weight_range
                                                                    'get weight difference
                                                                    If adodc_entries.Recordset!entry_weight >= .Recordset!entry_weight Then
                                                                        weight_difference = adodc_entries.Recordset!entry_weight - .Recordset!entry_weight
                                                                    ElseIf adodc_entries.Recordset!entry_weight < .Recordset!entry_weight Then
                                                                        weight_difference = .Recordset!entry_weight - adodc_entries.Recordset!entry_weight
                                                                    End If
                                                                    'check if entryweight passed
                                                                    If weight_difference > allowed_weight_difference Then
                                                                    'If weight_difference > 50 Then
                                                                        'NO, matcher weight failed
                                                                        GoTo check_matcher_eof
                                                                    Else
                                                                        'YES, matcher weight passed
                                                                        GoTo get_bet_difference_criteria
                                                                    End If
                                                                End If
                                                                
                                                            Else
                                                                'NO, entry_weight_criteria = false
                                                                'get bet_difference_criteria
get_bet_difference_criteria:
                                                                If participant_bet_criteria = True Then
                                                                    'YES, bet_difference_criteria = true
                                                                    'save variables
save_matcher_variables:
                                                                    puti_id = .Recordset!entry_id
                                                                    puti_participant_id = .Recordset!entry_owner
                                                                    puti_participant_bet = .Recordset!participant_bet
                                                                    
                                                                    'get puti_bet_difference
                                                                    If adodc_entries.Recordset!participant_bet >= .Recordset!participant_bet Then
                                                                        puti_bet_difference = adodc_entries.Recordset!participant_bet - .Recordset!participant_bet
                                                                    ElseIf adodc_entries.Recordset!participant_bet < .Recordset!participant_bet Then
                                                                        puti_bet_difference = .Recordset!participant_bet - adodc_entries.Recordset!participant_bet
                                                                    End If
                                                                    
                                                                    'get puti_weight_difference
                                                                    If adodc_entries.Recordset!entry_weight >= .Recordset!entry_weight Then
                                                                        puti_weight_difference = adodc_entries.Recordset!entry_weight - .Recordset!entry_weight
                                                                    ElseIf adodc_entries.Recordset!entry_weight < .Recordset!entry_weight Then
                                                                        puti_weight_difference = .Recordset!entry_weight - adodc_entries.Recordset!entry_weight
                                                                    End If
                                                                    
                                                                    'check if matcher eof again
check_matcher_eof_2:
                                                                    'If .Recordset.EOF = False Then
                                                                    '    .Recordset.MoveNext
                                                                    'End If
                                                                    If .Recordset.EOF = False Then
                                                                        'YES, matcher eof = false
                                                                        .Recordset.MoveNext
                                                                        If .Recordset.EOF = True Then
                                                                            GoTo save_match
                                                                        End If
                                                                        'compare saved variables to current recordset
                                                                        
                                                                        'get current recordset bet_difference
                                                                        If adodc_entries.Recordset!participant_bet >= .Recordset!participant_bet Then
                                                                            bet_difference = adodc_entries.Recordset!participant_bet - .Recordset!participant_bet
                                                                        ElseIf adodc_entries.Recordset!participant_bet < .Recordset!participant_bet Then
                                                                            bet_difference = .Recordset!participant_bet - adodc_entries.Recordset!participant_bet
                                                                        End If
                                                                        
                                                                        'get current recordset weight_difference
                                                                        If adodc_entries.Recordset!entry_weight >= .Recordset!entry_weight Then
                                                                            weight_difference = adodc_entries.Recordset!entry_weight - .Recordset!entry_weight
                                                                        ElseIf adodc_entries.Recordset!entry_weight < .Recordset!entry_weight Then
                                                                            weight_difference = .Recordset!entry_weight - adodc_entries.Recordset!entry_weight
                                                                        End If
                                                                        
                                                                        'start comparing
                                                                        If puti_weight_difference < weight_difference And puti_bet_difference <= bet_difference Then
                                                                            'YES, saved matcher won
                                                                            'GoTo check_matcher_eof_2
                                                                            
                                                                            'updated 20150420
                                                                            GoTo save_match
                                                                        Else
                                                                            'NO, saved matcher lost
                                                                            GoTo save_matcher_variables
                                                                        End If
                                                                    Else
                                                                        'NO, matcher eof = true
                                                                        'start saving match to database
                                                                        'mark entries as matched
                                                                        'add + to counters
                                                                    GoTo check_if_the_same
save_match:
                                                                        'MsgBox criteria_query & " - " & puti_weight_difference & " - Criteria used: " & adodc_criteria.Recordset!criteria_set_id
                                                                        adodc_matches.RecordSource = "select * from matches"
                                                                        adodc_matches.Refresh
                                                                        'save to database
                                                                        With adodc_matches
                                                                            .Recordset.AddNew
                                                                            .Recordset!match_event_id = event_id
                                                                            .Recordset!match_schedule = event_match_number
                                                                            .Recordset!entry_pula_id = adodc_entries.Recordset!entry_id
                                                                            .Recordset!entry_puti_id = puti_id
                                                                            .Recordset!criteria_set_used = adodc_criteria.Recordset!criteria_set_id
                                                                            .Recordset.Update
                                                                        End With
                                                                        
                                                                        'mark entries as matched
                                                                        adodc_matcher2.RecordSource = "select * from entries where entry_id = " & adodc_entries.Recordset!entry_id & ""
                                                                        adodc_matcher2.Refresh
                                                                        adodc_matcher2.Recordset!entry_matching_status = 2
                                                                        adodc_matcher2.Recordset.Update
                                                                        
                                                                        adodc_matcher2.RecordSource = "select * from entries where entry_id = " & puti_id & ""
                                                                        adodc_matcher2.Refresh
                                                                        adodc_matcher2.Recordset!entry_matching_status = 2
                                                                        adodc_matcher2.Recordset.Update
                                                                        
                                                                        'add +1 to counters
                                                                        event_match_number = event_match_number + 1
                                                                        event_match_counter = event_match_counter + 1
                                                                        event_day_counter = event_day_counter + 1
                                                                        event_day_segment_counter = event_day_segment_counter + 1
                                                                        
                                                                        datagrid_matches.SelBookmarks.Add datagrid_matches.Bookmark
                                                                        
                                                                        adodc_matches.RecordSource = "select * from matches_view_query_proto where match_event_id = " & event_id & " order by match_schedule desc"
                                                                        adodc_matches.Refresh
                                                                        datagrid_matches.Refresh
                                                                       
                                                                        adodc_unmatched.Refresh
                                                                        
                                                                        GoSub sub_event_day_segment_update
                                                                        GoTo entries_eof_check
                                                                    End If
                                                                Else
                                                                    'NO, bet_difference_criteria = false
                                                                    'save variables
                                                                    puti_id = .Recordset!entry_id
                                                                    puti_participant_id = .Recordset!entry_owner
check_if_the_same:
                                                                    If .Recordset!entry_owner = adodc_entries.Recordset!entry_owner Then
                                                                        GoTo check_matcher_eof
                                                                    End If
                                                                    
                                                                    GoTo save_match
                                                                End If
                                                            End If
                                                        Else
                                                            'matcher not ready
                                                            GoTo check_matcher_eof
                                                        End If
                                                    Else
                                                        'if_ready_criteria = false
                                                        GoTo get_entry_weight_criteria
                                                    End If
                                                End If
                                            Else
                                                'no_match_criteria = false
                                                GoTo get_if_ready_criteria
                                            End If
                                        Else
criteria_eof_check:
                                            If adodc_criteria.Recordset.EOF = False Then
                                                'MsgBox adodc_criteria.Recordset!criteria_set_id
                                                adodc_criteria.Recordset.MoveNext
                                                'MsgBox adodc_criteria.Recordset!criteria_set_id
                                                
                                                If adodc_criteria.Recordset.EOF Then
                                                    GoTo entries_eof_check
                                                End If
                                                
                                                If adodc_entries.Recordset.EOF Then
                                                    GoTo entries_move_first
                                                End If
                                                
                                                GoSub sub_translate_criteria_set
                                                
                                                GoTo get_matches
                                            Else
                                                'check if entries.eof
entries_eof_check:
                                                If adodc_entries.Recordset.EOF = False Then
                                                    adodc_entries.Recordset.MoveNext
                                                    If adodc_entries.Recordset.EOF = True Then
                                                        
                                                        If adodc_criteria.Recordset.EOF = False Then
                                                            GoTo criteria_eof_check
                                                        End If
                                                        
                                                        GoTo entries_eof_true
                                                    Else
                                                        If adodc_criteria.Recordset.EOF Then
                                                            GoTo entries_eof_check
                                                        Else
                                                            GoTo get_matches
                                                        End If
                                                    End If
                                                Else
entries_eof_true:
                                                    event_day_segment_counter = event_day_segment_counter + 1
                                                    
                                                    If adodc_criteria.Recordset.EOF And adodc_entries.Recordset.EOF Then
                                                        event_day_segments_counter = event_day_segments_counter + 1
                                                        event_match_counter = event_match_counter + 1
                                                        GoTo counter_updated
                                                    End If
                                                    
                                                    GoTo counter_updated
                                                End If
                                            End If
                                        End If
                                    Else
                                        MsgBox "No criteria found!", vbCritical
                                        GoTo matching_complete
                                    End If
                                Else
                                    'MsgBox "Entries not enough for matching!", vbCritical
                                    'GoTo matching_complete
                                    event_day_segment_counter = event_day_segment_counter + 1
                                    GoTo counter_updated
                                End If
                            Else
                                GoTo reset_daily_counters
                            End If
                        Else
                            MsgBox "Total possible matches reached!", vbInformation
                            GoTo matching_complete
                        End If
                Else
                    MsgBox "Not enough/ No entries!", vbCritical
                    Exit Sub
                End If
            Else
                MsgBox "Event don't exist or event duplicate found!", vbCritical
                Exit Sub
            End If
        End With
    Else
        MsgBox "Select event first!", vbCritical
        Exit Sub
    End If
    Exit Sub

sub_translate_criteria_set:
    Dim criteria_counter As Integer
    Dim criterias(5), and_query, criteria_query_union As String
    Dim where_on As Boolean
    
    where_on = False
    no_match_criteria = False
    if_ready_criteria = False
    entry_weight_criteria = False
    entry_schedule_type_criteria = False
    participant_category_criteria = False
    participant_bet_criteria = False
    
    criteria_query = ""
    and_query = ""
    entry_weight_criteria_query = ""
    entry_schedule_type_criteria_query = ""
    participant_category_criteria_query = ""
    
    With adodc_criteria
        criterias(0) = .Recordset!first_criteria
        criterias(1) = .Recordset!second_criteria
        criterias(2) = .Recordset!third_criteria
        criterias(3) = .Recordset!fourth_criteria
        criterias(4) = .Recordset!fifth_criteria
        criterias(5) = .Recordset!sixth_criteria
        
        For criteria_counter = 0 To UBound(criterias)
            If criterias(criteria_counter) = "by weight" Then
                where_on = True
                entry_weight_criteria = True
                entry_weight_criteria_query = " (entry_weight - " & adodc_entries.Recordset!entry_weight & ") as weight_difference from entries_query "
            ElseIf criterias(criteria_counter) = "by schedule" Then
                where_on = True
                entry_schedule_type_criteria = True
                entry_schedule_type_criteria_query = " entry_schedule_type = '" & match_schedule_type & "' "
            ElseIf criterias(criteria_counter) = "by category" Then
                where_on = True
                participant_category_criteria = True
                participant_category_criteria_query = " participant_category = " & adodc_entries.Recordset!participant_category & " "
            ElseIf criterias(criteria_counter) = "by bet" Then
                participant_bet_criteria = True
            ElseIf criterias(criteria_counter) = "no match" Then
                no_match_criteria = True
            ElseIf criterias(criteria_counter) = "5 fights interval" Then
                if_ready_criteria = True
            ElseIf criterias(criteria_counter) = "none" Then
            End If
        Next
    End With

    'start building criteria_query
    If entry_weight_criteria = True Then
        criteria_query = "Select entry_weight, entry_id, entry_owner, participant_bet, " & _
            " (entry_weight - " & adodc_entries.Recordset!entry_weight & ") as weight_difference from entries_query"
        If entry_schedule_type_criteria = True And participant_category_criteria = True Then
            and_query = " and entry_schedule_type = '" & event_fight_schedule & "' and participant_category = " & adodc_entries.Recordset!participant_category
        ElseIf entry_schedule_type_criteria = False And participant_category_criteria = True Then
            and_query = " and participant_category = " & adodc_entries.Recordset!participant_category
        ElseIf entry_schedule_type_criteria = True And participant_category_criteria = False Then
            and_query = " and entry_schedule_type = '" & event_fight_schedule & "'"
        ElseIf entry_schedule_type_criteria = False And participant_category_criteria = False Then
        End If
        
        criteria_query = "select entry_weight, entry_id, entry_owner, participant_bet, " & _
            " (entry_weight - " & adodc_entries.Recordset!entry_weight & ") as weight_difference from entries_query " & _
            " where participant_event = " & adodc_events.Recordset!event_id & "" & _
            " and entry_matching_status = 'Unmatched' " & _
            " and entry_weight >= " & adodc_entries.Recordset!entry_weight & "" & _
            " and entry_id <> " & adodc_entries.Recordset!entry_id & "" & _
            " and entry_owner <> " & adodc_entries.Recordset!entry_owner & "" & _
            " " & and_query & " " & _
            " union select entry_weight, entry_id, entry_owner, participant_bet, " & _
            " (" & adodc_entries.Recordset!entry_weight & " - entry_weight) as weight_difference from entries_query " & _
            " where participant_event = " & adodc_events.Recordset!event_id & "" & _
            " and entry_matching_status = 'Unmatched' " & _
            " and entry_weight < " & adodc_entries.Recordset!entry_weight & "" & _
            " and entry_id <> " & adodc_entries.Recordset!entry_id & "" & _
            " and entry_owner <> " & adodc_entries.Recordset!entry_owner & "" & _
            " " & and_query & " "
    Else
        If entry_schedule_type_criteria = True And participant_category_criteria = True Then
            and_query = " and entry_schedule_type = '" & event_fight_schedule & "' and participant_category = " & adodc_entries.Recordset!participant_category
        ElseIf entry_schedule_type_criteria = False And participant_category_criteria = True Then
            and_query = " and participant_category = " & adodc_entries.Recordset!participant_category
        ElseIf entry_schedule_type_criteria = True And participant_category_criteria = False Then
            and_query = " and entry_schedule_type = '" & event_fight_schedule
        ElseIf entry_schedule_type_criteria = False And participant_category_criteria = False Then
        End If
        
        criteria_query = "Select entry_weight, entry_id, entry_owner, participant_bet from entries_query " & _
            " where participant_event = " & adodc_events.Recordset!event_id & "" & _
            " and entry_matching_status = 'Unmatched' " & _
            " and entry_weight >= " & adodc_entries.Recordset!entry_weight & "" & _
            " and entry_id <> " & adodc_entries.Recordset!entry_id & "" & _
            " and entry_owner <> " & adodc_entries.Recordset!entry_owner & "" & _
            " " & and_query & " "
    End If
    
    'MsgBox criteria_query, vbInformation
Return

sub_event_day_segment_update:
    If event_day_segment_counter <= event_day_segment Then
        event_fight_schedule = "Early Fight"
    ElseIf (event_day_segment_counter > event_day_segment) And (event_day_segment_counter <= (event_day_segment * 2)) Then
        event_fight_schedule = "None"
    ElseIf (event_day_segment_counter > (event_day_segment * 2)) And (event_day_segment_counter <= (event_day_segment * 3)) Then
        event_fight_schedule = "Late Fight"
    Else
        GoTo counter_updated
    End If
Return

sub_matched_check:
    'check if this entry is already matched
    adodc_matcher2.RecordSource = "select * from entries where entry_id=" & adodc_entries.Recordset!entry_id & " and entry_matching_status = 2"
    adodc_matcher2.Refresh
    If adodc_matcher2.Recordset.RecordCount > 0 Then
        'YES, entry is already matched
        GoTo entries_eof_check
    End If
Return

sub_check_if_ready:

Return

matching_complete:
    'adodc_matcher.RecordSource = "Select * from process"
    'adodc_matcher.Refresh
    
    'adodc_matcher.Recordset!current_event_id = 0
    'adodc_matcher.Recordset!current_event_name = "None"
    'adodc_matcher.Recordset!current_total = 0
    'adodc_matcher.Recordset.Update
    
    'Shell "taskkill.exe /f /t /im splasher.exe"
    
    'mdi_main.Visible = True
    
    MsgBox "Finished Matching!", vbOKOnly
    
    adodc_matches.RecordSource = "select * from matches_view_query_proto where match_event_id = " & adodc_events.Recordset!event_id & ""
    adodc_matches.Refresh
    adodc_unmatched.Refresh
    adodc_highlighter.Refresh
    datagrid_matches.Refresh
End Sub



Private Sub command_matches_edit_close_Click()
    adodc_matches.Refresh
    adodc_unmatched.Recordset.Requery
    adodc_unmatched.Refresh
    Call unmatched_update
    frame_edit_matches.Visible = False
End Sub

Private Sub command_move_Click()
    'none entry_id 110
    Dim move_pula_id, move_puti_id, move_slot, mover_pula_id, mover_puti_id, mover_slot As Integer
    mover_pula_id = adodc_edit_matches.Recordset!pula_id
    mover_puti_id = adodc_edit_matches.Recordset!puti_id
    mover_slot = adodc_edit_matches.Recordset!match_id
    
    If datacombo_move_match.Text <> " " Then
        'get the to be replace values using the chosen match_schedule in the datacombo
        With adodc_matcher
            .RecordSource = "select * from matches where match_schedule = " & datacombo_move_match.Text & " and match_event_id = " & adodc_edit_matches.Recordset!match_event_id & " "
            .Refresh
            If .Recordset.RecordCount = 0 Then
                MsgBox "Match not found!", vbCritical
                Exit Sub
            End If
            
            move_pula_id = .Recordset!entry_pula_id
            move_puti_id = .Recordset!entry_puti_id
            move_slot = .Recordset!match_id
    
            'start chaning the values, start with the current entry
            .Recordset!entry_pula_id = mover_pula_id
            .Recordset!entry_puti_id = mover_puti_id
            .Recordset.Update
            .Refresh
            
            'then with the replaced entry
            adodc_edit_matches.Recordset!pula_id = move_pula_id
            adodc_edit_matches.Recordset!puti_id = move_puti_id
            adodc_edit_matches.Recordset.Update
            adodc_edit_matches.Refresh
            
            frame_edit_matches.Enabled = True
            frame_move_match.Visible = False
            
            adodc_matches.Refresh
            
            MsgBox "Match succesfuly moved!", vbInformation
            Exit Sub
        End With
    Else
        MsgBox "Match schedule not found!", vbCritical
        Exit Sub
    End If
End Sub

Private Sub command_move_match_Click()
    frame_edit_matches.Enabled = False
    With frame_move_match
        .Height = 3015
        .Left = 5400
        .Top = 1920
        .Width = 3975
    End With
    frame_move_match.Visible = True
End Sub

Private Sub command_print_Click()
    If Not adodc_matches.Recordset.RecordCount = 0 Then
        Set DataReport1.DataSource = adodc_matches
        
        DataReport1.Title = adodc_matches.Recordset!event_name & vbCrLf & adodc_matches.Recordset!schedule
        
        DataReport1.Show
    End If
End Sub

Private Sub command_pula_remove_Click()
    'unmatch the current entry before removing it to the matches table
    With adodc_entries
        .RecordSource = "select * from entries where entry_id = " & adodc_matches.Recordset!pula_id & ""
        .Refresh
        If .Recordset.RecordCount = 0 Then
            MsgBox "Can't find entry!", vbCritical
            Exit Sub
        Else
            .Recordset!entry_matching_status = 1
            .Recordset.Update
        End If
    End With

    adodc_edit_matches.Recordset!pula_id = 171
    adodc_edit_matches.Recordset.Update
    adodc_edit_matches.Recordset.Requery
    adodc_edit_matches.Refresh
    
    Call unmatched_update
    
    adodc_unmatched.Recordset.Requery
    adodc_unmatched.Refresh
    Call unmatched_update
End Sub

Private Sub command_puti_remove_Click()
    'unmatch the current entry before removing it to the matches table
    With adodc_entries
        .RecordSource = "select * from entries where entry_id = " & adodc_matches.Recordset!puti_id & ""
        .Refresh
        If .Recordset.RecordCount = 0 Then
            MsgBox "Can't find entry!", vbCritical
            Exit Sub
        Else
            .Recordset!entry_matching_status = 1
            .Recordset.Update
        End If
    End With

    adodc_edit_matches.Recordset!puti_id = 171
    adodc_edit_matches.Recordset.Update
    adodc_edit_matches.Recordset.Requery
    adodc_edit_matches.Refresh
    
    Call unmatched_update
    
    adodc_unmatched.Recordset.Requery
    adodc_unmatched.Refresh
    Call unmatched_update
End Sub

Private Sub command_replace_pula_Click()
    frame_edit_matches.Enabled = False
    With frame_replace
        .Caption = "Replace Pula"
        .Height = 7215
        .Left = 2400
        .Top = 2040
        .Width = 11055
        .Visible = True
    End With
End Sub

Private Sub command_replace_puti_Click()
    frame_edit_matches.Enabled = False
    With frame_replace
        .Caption = "Replace Puti"
        .Height = 7215
        .Left = 2400
        .Top = 2040
        .Width = 11055
        .Visible = True
    End With
End Sub

Private Sub command_replace_select_Click()
    If frame_replace.Caption = "Replace Pula" Then
        'the user clicked the command_replace_pula
        'get the current_recordset by using the match_id
        'adodc_entries.RecordSource = "select * from matches where match_id = " & adodc_edit_matches.Recordset!match_id & ""
        'adodc_entries.Refresh
        'If adodc_entries.Recordset.Recordset = 0 Then
            'MsgBox "Entry not found!", vbCritical
            'Exit Sub
        'End If
        
        'mark the current_record or the to be replace record as unmatched
        adodc_matcher.RecordSource = "select * from entries where entry_id  = " & adodc_edit_matches.Recordset!pula_id & ""
        adodc_matcher.Refresh
        If adodc_matcher.Recordset.RecordCount = 0 Then
            MsgBox "Entry not found!", vbCritical
            Exit Sub
        End If
        
        adodc_matcher.Recordset!entry_matching_status = 1
        adodc_matcher.Recordset.Update
        adodc_matcher.Refresh
        
        'mark the selected replacement as matched
        adodc_entries.RecordSource = "select * from entries where entry_id = " & adodc_unmatched.Recordset!entry_id & ""
        adodc_entries.Refresh
        If adodc_entries.Recordset.RecordCount = 0 Then
            MsgBox "Entry not found!", vbCritical
            Exit Sub
        End If
        
        adodc_entries.Recordset!entry_matching_status = 2
        adodc_entries.Recordset.Update
        adodc_entries.Refresh
        
        adodc_edit_matches.Recordset!criteria_set_used = 42
        adodc_edit_matches.Recordset!pula_id = adodc_unmatched.Recordset!entry_id
        adodc_edit_matches.Recordset.Update
        adodc_edit_matches.Recordset.Requery
        adodc_edit_matches.Refresh
    ElseIf frame_replace.Caption = "Replace Puti" Then
        'the user clicked the command_replace_puti
        
        'mark the current_record or the to be replace record as unmatched
        adodc_matcher.RecordSource = "select * from entries where entry_id  = " & adodc_edit_matches.Recordset!puti_id & ""
        adodc_matcher.Refresh
        If adodc_matcher.Recordset.RecordCount = 0 Then
            MsgBox "Entry not found!", vbCritical
            Exit Sub
        End If
        
        adodc_matcher.Recordset!entry_matching_status = 1
        adodc_matcher.Recordset.Update
        adodc_matcher.Refresh
        
        'mark the selected replacement as matched
        adodc_entries.RecordSource = "select * from entries where entry_id = " & adodc_unmatched.Recordset!entry_id & ""
        adodc_entries.Refresh
        If adodc_entries.Recordset.RecordCount = 0 Then
            MsgBox "Entry not found!", vbCritical
            Exit Sub
        End If
        
        adodc_entries.Recordset!entry_matching_status = 2
        adodc_entries.Recordset.Update
        adodc_entries.Refresh
        
        adodc_edit_matches.Recordset!criteria_set_used = 42
        adodc_edit_matches.Recordset!puti_id = adodc_unmatched.Recordset!entry_id
        adodc_edit_matches.Recordset.Update
        adodc_edit_matches.Recordset.Requery
        adodc_edit_matches.Refresh
    End If
    
    adodc_matches.Refresh
    adodc_unmatched.Refresh
    frame_replace.Visible = False
    frame_edit_matches.Enabled = True
    MsgBox "Replace succesful!", vbInformation
End Sub



Private Sub command_search_Click()
    If text_search.Text <> "" Then
        search = text_search.Text
        Do While datagrid_matches.SelBookmarks.Count > 0
            datagrid_matches.SelBookmarks.Remove 0
        Loop
    
        timer_search.Enabled = True
        adodc_matches.Recordset.MoveFirst
    Else
        MsgBox "Please enter something to search.", vbInformation
        text_search.SetFocus
    End If
End Sub

Private Sub command_settings_add_Click()
    If command_settings_add.Caption = "Add" Then
        command_settings_save.Enabled = True
        command_settings_edit.Enabled = False
        command_settings_delete.Enabled = False
        command_settings_use_defaults.Enabled = False
        
        datacombo_first_criteria.DataField = ""
        datacombo_second_criteria.DataField = ""
        datacombo_third_criteria.DataField = ""
        datacombo_fourth_criteria.DataField = ""
        datacombo_fifth_criteria.DataField = ""
        datacombo_sixth_criteria.DataField = ""
        
        datacombo_first_criteria.Text = ""
        datacombo_second_criteria.Text = ""
        datacombo_third_criteria.Text = ""
        datacombo_fourth_criteria.Text = ""
        datacombo_fifth_criteria.Text = ""
        datacombo_sixth_criteria.Text = ""
        
        datacombo_first_criteria.Enabled = True
        datacombo_second_criteria.Enabled = True
        datacombo_third_criteria.Enabled = True
        datacombo_fourth_criteria.Enabled = True
        datacombo_fifth_criteria.Enabled = True
        datacombo_sixth_criteria.Enabled = True
        
        command_settings_add.Caption = "Cancel"
    Else
        command_settings_add.Caption = "Add"
        Call controls_default
    End If
End Sub

Private Sub command_settings_Click()
    With frame_matching_settings
        .Height = 6615
        .Left = 240
        .Top = 480
        .Width = 14175
        .Visible = True
    End With
    Call controls_default
End Sub

Private Sub command_settings_delete_Click()
    If adodc_matching_settings.Recordset.RecordCount <> 0 Then
        If MsgBox("Are you sure you want to delete this record?", vbYesNo) = vbYes Then
            With adodc_matching_settings
                .RecordSource = "select * from matching_settings where criteria_set_id =" & .Recordset!criteria_set_id & ""
                .Refresh
                
                If .Recordset.RecordCount = 1 Then
                    .Recordset.Delete
                    Call controls_default
                    MsgBox "Criteria Set Deleted!", vbInformation
                Else
                    Call controls_default
                    MsgBox "More than 1 record returned!", vbCritical
                End If
            End With
        Else
            
        End If
    Else
        MsgBox "No record found!", vbCritical
    End If
    Call controls_default
End Sub

Private Sub command_settings_edit_Click()
    If command_settings_edit.Caption = "Edit" Then
        command_settings_add.Enabled = False
        command_settings_delete.Enabled = False
        command_settings_use_defaults.Enabled = False
        
        datacombo_first_criteria.Enabled = True
        datacombo_second_criteria.Enabled = True
        datacombo_third_criteria.Enabled = True
        datacombo_fourth_criteria.Enabled = True
        datacombo_fifth_criteria.Enabled = True
        datacombo_sixth_criteria.Enabled = True
        
        datacombo_first_criteria.DataField = ""
        datacombo_second_criteria.DataField = ""
        datacombo_third_criteria.DataField = ""
        datacombo_fourth_criteria.DataField = ""
        datacombo_fifth_criteria.DataField = ""
        datacombo_sixth_criteria.DataField = ""
        
        datagrid_matching_settings.Enabled = False
        
        command_settings_edit.Caption = "Update"
        
        Exit Sub
    Else
        GoSub sub_check_datacombo_duplicate
        Dim id_to_edit As Integer
        id_to_edit = adodc_matching_settings.Recordset!criteria_set_id
        With adodc_matching_settings
            .RecordSource = "select * from matching_settings_master where criteria_name ='" & datacombo_first_criteria.Text & "'"
            .Refresh
            If .Recordset.RecordCount = 1 Then
                datacombo_first_criteria.Text = .Recordset!criteria_id
                .RecordSource = "select * from matching_settings_master where criteria_name ='" & datacombo_second_criteria.Text & "'"
                .Refresh
                If .Recordset.RecordCount = 1 Then
                    datacombo_second_criteria.Text = .Recordset!criteria_id
                    .RecordSource = "select * from matching_settings_master where criteria_name ='" & datacombo_third_criteria.Text & "'"
                    .Refresh
                    If .Recordset.RecordCount = 1 Then
                        datacombo_third_criteria.Text = .Recordset!criteria_id
                        .RecordSource = "select * from matching_settings_master where criteria_name ='" & datacombo_fourth_criteria.Text & "'"
                        .Refresh
                        If .Recordset.RecordCount = 1 Then
                            datacombo_fourth_criteria.Text = .Recordset!criteria_id
                            .RecordSource = "select * from matching_settings_master where criteria_name ='" & datacombo_fifth_criteria.Text & "'"
                            .Refresh
                            If .Recordset.RecordCount = 1 Then
                                datacombo_fifth_criteria.Text = .Recordset!criteria_id
                                .RecordSource = "select * from matching_settings_master where criteria_name ='" & datacombo_sixth_criteria.Text & "'"
                                .Refresh
                                If .Recordset.RecordCount = 1 Then
                                    datacombo_sixth_criteria.Text = .Recordset!criteria_id
                                    'search the to be edited record by using the id_to_edit variable
                                    .RecordSource = "select * from matching_settings where criteria_set_id = " & id_to_edit & ""
                                    .Refresh
                                    If .Recordset.RecordCount = 1 Then
                                        'start saving
                                        .Recordset!first_criteria = datacombo_first_criteria.Text
                                        .Recordset!second_criteria = datacombo_second_criteria.Text
                                        .Recordset!third_criteria = datacombo_third_criteria.Text
                                        .Recordset!fourth_criteria = datacombo_fourth_criteria.Text
                                        .Recordset!fifth_criteria = datacombo_fifth_criteria.Text
                                        .Recordset!sixth_criteria = datacombo_sixth_criteria.Text
                                        'On Error GoTo edit_error
                                        .Recordset.Update
                                        Call controls_default
                                        MsgBox "Record successfuly edited!", vbInformation
                                        Call controls_default
                                        Exit Sub
edit_error:
                                        MsgBox "Error Editing Record!", vbCritical
                                        Call controls_default
                                        Exit Sub
                                    Else
                                        MsgBox "Criteria Settings Error!", vbCritical
                                        Call controls_default
                                    End If
                                Else
                                    MsgBox "Criteria Settings Master Error!", vbCritical
                                    Call controls_default
                                End If
                            Else
                                MsgBox "Criteria Settings Master Error!", vbCritical
                                Call controls_default
                            End If
                        Else
                            MsgBox "Criteria Settings Master Error!", vbCritical
                            Call controls_default
                        End If
                    Else
                        MsgBox "Criteria Settings Master Error!", vbCritical
                        Call controls_default
                    End If
                Else
                    MsgBox "Criteria Settings Master Error!", vbCritical
                    Call controls_default
                End If
            Else
                MsgBox "Criteria Settings Master Error!", vbCritical
                Call controls_default
            End If
        End With
    End If
    
sub_check_datacombo_duplicate:
    Dim counter, set_counter As Long
    Dim criteria_array(5), set_criteria_array(5) As String
    
    criteria_array(0) = datacombo_first_criteria.Text
    criteria_array(1) = datacombo_second_criteria.Text
    criteria_array(2) = datacombo_third_criteria.Text
    criteria_array(3) = datacombo_fourth_criteria.Text
    criteria_array(4) = datacombo_fifth_criteria.Text
    criteria_array(5) = datacombo_sixth_criteria.Text
    
    'copy criteria_array to set_criteria_array
    For counter = 0 To UBound(criteria_array)
        set_criteria_array(counter) = criteria_array(counter)
    Next
    
    For set_counter = 0 To UBound(set_criteria_array)
        For counter = 0 To UBound(criteria_array)
            If set_criteria_array(set_counter) = criteria_array(counter) And counter <> set_counter And set_criteria_array(set_counter) <> "none" And criteria_array(counter) <> "" Then
                MsgBox "Duplicate criteria in one set not allowed!", vbCritical
                GoTo edit_error
            End If
        Next
    Next
On Error Resume Next
Return
End Sub

Private Sub command_settings_save_Click()
    If MsgBox("Are you sure you want to save this?", vbYesNo) = vbYes Then
        Dim criteria_set(0 To 5) As String
    
        If datacombo_first_criteria.Text = "" Then
            datacombo_first_criteria.Text = "none"
        End If
        If datacombo_second_criteria.Text = "" Then
            datacombo_second_criteria.Text = "none"
        End If
        If datacombo_third_criteria.Text = "" Then
            datacombo_third_criteria.Text = "none"
        End If
        If datacombo_fourth_criteria.Text = "" Then
            datacombo_fourth_criteria.Text = "none"
        End If
        If datacombo_fifth_criteria.Text = "" Then
            datacombo_fifth_criteria.Text = "none"
        End If
        If datacombo_sixth_criteria.Text = "" Then
            datacombo_sixth_criteria.Text = "none"
        End If
        
        GoSub sub_check_datacombo_duplicate
        
        With adodc_matching_settings
            .RecordSource = "select * from matching_settings_master where criteria_name = '" & datacombo_first_criteria.Text & "'"
            .Refresh
            If .Recordset.RecordCount = 1 Then
                datacombo_first_criteria.Text = .Recordset!criteria_id
                
                .RecordSource = "select * from matching_settings_master where criteria_name = '" & datacombo_second_criteria.Text & "'"
                .Refresh
                If .Recordset.RecordCount = 1 Then
                    datacombo_second_criteria.Text = .Recordset!criteria_id
                    
                    .RecordSource = "select * from matching_settings_master where criteria_name = '" & datacombo_third_criteria.Text & "'"
                    .Refresh
                    If .Recordset.RecordCount = 1 Then
                        datacombo_third_criteria.Text = .Recordset!criteria_id
                        
                        .RecordSource = "select * from matching_settings_master where criteria_name = '" & datacombo_fourth_criteria.Text & "'"
                        .Refresh
                        If .Recordset.RecordCount = 1 Then
                            datacombo_fourth_criteria.Text = .Recordset!criteria_id
                            
                            .RecordSource = "select * from matching_settings_master where criteria_name = '" & datacombo_fifth_criteria.Text & "'"
                            .Refresh
                            If .Recordset.RecordCount = 1 Then
                                datacombo_fifth_criteria.Text = .Recordset!criteria_id
                                
                                .RecordSource = "select * from matching_settings_master where criteria_name = '" & datacombo_sixth_criteria.Text & "'"
                                .Refresh
                                If .Recordset.RecordCount = 1 Then
                                    datacombo_sixth_criteria.Text = .Recordset!criteria_id
                                    
                                    'start adding
                                    .RecordSource = "select * from matching_settings"
                                    .Refresh
                                    
                                    .Recordset.AddNew
                                    .Recordset!first_criteria = datacombo_first_criteria.Text
                                    .Recordset!second_criteria = datacombo_second_criteria.Text
                                    .Recordset!third_criteria = datacombo_third_criteria.Text
                                    .Recordset!fourth_criteria = datacombo_fourth_criteria.Text
                                    .Recordset!fifth_criteria = datacombo_fifth_criteria.Text
                                    .Recordset!sixth_criteria = datacombo_sixth_criteria.Text
                                    On Error GoTo save_error
                                    .Recordset.Update
                                    Call controls_default
                                    MsgBox "Adding Criteria Set Successful!", vbInformation
                                    Exit Sub
save_error:
                                    Call controls_default
                                    MsgBox "Error Adding Criteria Set!", vbCritical
                                Else
                                    MsgBox "Multiple/ No Result Error!", vbCritical, "Criteria Error"
                                    Call controls_default
                                End If
                            Else
                                MsgBox "Multiple/ No Result Error!", vbCritical, "Criteria Error"
                                Call controls_default
                            End If
                        Else
                            MsgBox "Multiple/ No Result Error!", vbCritical, "Criteria Error"
                            Call controls_default
                        End If
                    Else
                        MsgBox "Multiple/ No Result Error!", vbCritical, "Criteria Error"
                        Call controls_default
                    End If
                Else
                    MsgBox "Multiple/ No Result Error!", vbCritical, "Criteria Error"
                    Call controls_default
                End If
            Else
                MsgBox "Multiple/ No Result Error!", vbCritical, "Criteria Error"
                Call controls_default
            End If
        End With
    Else
    
    End If
    
    Exit Sub
    
sub_check_datacombo_duplicate:
    Dim counter, set_counter As Long
    Dim criteria_array(5), set_criteria_array(5) As String
    
    criteria_array(0) = datacombo_first_criteria.Text
    criteria_array(1) = datacombo_second_criteria.Text
    criteria_array(2) = datacombo_third_criteria.Text
    criteria_array(3) = datacombo_fourth_criteria.Text
    criteria_array(4) = datacombo_fifth_criteria.Text
    criteria_array(5) = datacombo_sixth_criteria.Text
    
    'copy criteria_array to set_criteria_array
    For counter = 0 To UBound(criteria_array)
        set_criteria_array(counter) = criteria_array(counter)
    Next
    
    For set_counter = 0 To UBound(set_criteria_array)
        For counter = 0 To UBound(criteria_array)
            If set_criteria_array(set_counter) = criteria_array(counter) And counter <> set_counter And set_criteria_array(set_counter) <> "none" And criteria_array(counter) <> "" Then
                MsgBox "Duplicate criteria in one set not allowed!", vbCritical
                GoTo save_error
            End If
        Next
    Next
Return
End Sub



Private Sub Command1_Click()
    timer_bookmark.Enabled = False
End Sub

Private Sub datacombo_event_name_Change()
    If datacombo_event_name.Text = "" Then
    Else
        With adodc_events.Recordset
            .MoveFirst
            While Not .EOF
                If !event_name = datacombo_event_name.Text Then
                    adodc_matches.RecordSource = "select * from matches_view_query_proto where match_event_id = " & !event_id & " order by match_schedule asc"
                    adodc_matches.Refresh
                    
                    adodc_unmatched.RecordSource = "select * from entries_unmatched_query where participant_event = " & !event_id & " and entry_matching_status = 'Unmatched'"
                    adodc_unmatched.Refresh
                    
                    Exit Sub
                End If
                .MoveNext
                
                If .EOF Then
                    adodc_matches.RecordSource = "select * from matches_view_query_proto where match_event_id = 0 "
                    adodc_matches.Refresh
                End If
            Wend
        End With
    End If
End Sub


Private Sub datagrid_matches_Change()
    adodc_unmatched.Recordset.Requery
End Sub

Private Sub datagrid_matches_DblClick()
    If adodc_matches.Recordset.RecordCount <> 0 Then
        With frame_edit_matches
            With adodc_edit_matches
                .RecordSource = "select * from matches_view_query_proto where match_id = " & adodc_matches.Recordset!match_id & ""
                .Refresh
            End With
        
            .Height = 2895
            .Left = 3120
            .Top = 2040
            .Width = 8175
            .Visible = True
        End With
    End If
End Sub





Private Sub datagrid_unmatched_DblClick()
    With frame_edit_unmatched
        .Height = 2655
        .Left = 12000
        .Top = 2640
        .Width = 8055
        .Visible = True
    End With
    frame_unmatched.Enabled = False
End Sub

Private Sub Form_Load()
    Call controls_default
    sAppName = "splasher.exe"
    sAppPath = "splasher.exe"
End Sub

Private Sub command_highlighter_Click()
    'verify if adodc_matches is not empty
    If adodc_matches.Recordset.RecordCount > 0 Then
        'get the criteria_id of the selected criteria_summary
        With adodc_matcher
            .RecordSource = "Select * from matching_settings_summary_query where criteria_summary = '" & datacombo_highlighter.Text & "'"
            .Refresh
            If .Recordset.RecordCount <> 1 Then
                'display error message
                MsgBox "None or more than 1 result(s) found!", vbCritical
                Exit Sub
            Else
                'start highlighting matches that were matched using the found criteria_id
                
                'clear existing highlights first
                Do While datagrid_matches.SelBookmarks.Count > 0
                    frame_highlighter.Caption = "Highlighter - Clearing existing bookmarks..."
                    datagrid_matches.SelBookmarks.Remove 0
                Loop
    
                timer_bookmark.Enabled = True
                adodc_matches.Recordset.MoveFirst
            End If
        End With
    Else
        MsgBox "Select matches first!", vbInformation
    End If
End Sub


Private Sub timer_bookmark_Timer()
    'Do While datagrid_matches.SelBookmarks.Count > 0
            'datagrid_matches.SelBookmarks.Remove 0
    'Loop
    
    'If Not adodc_matches.Recordset.EOF Then
        'adodc_matches.Recordset.MoveNext
    'End If
    'If adodc_matches.Recordset.EOF Then
        'adodc_matches.Recordset.MoveFirst
    'End If
    'datagrid_matches.SelBookmarks.Add adodc_matches.Recordset.Bookmark
    
    If Not adodc_matches.Recordset.EOF Then
        frame_highlighter.Caption = "Highlighter - Highlighting..."
        If adodc_matches.Recordset!criteria_set_used = adodc_matcher.Recordset!criteria_set_id Then
            datagrid_matches.SelBookmarks.Add adodc_matches.Recordset.Bookmark
        End If
        adodc_matches.Recordset.MoveNext
    Else
        frame_highlighter.Caption = "Highlighter - Done."
        timer_bookmark.Enabled = False
    End If
    
End Sub

Sub unmatched_update()
    Dim unmatched_query As String
    
    unmatched_query = adodc_unmatched.RecordSource
    adodc_unmatched.RecordSource = "select * from entries_unmatched_query where participant_event = 0"
    adodc_unmatched.Refresh
    
    adodc_unmatched.RecordSource = unmatched_query
    adodc_unmatched.Refresh
End Sub

Private Sub timer_search_Timer()
    If Not adodc_matches.Recordset.EOF Then
        'If adodc_matches.Recordset!entry_ = adodc_matcher.Recordset!criteria_set_id Then
            'datagrid_matches.SelBookmarks.Add adodc_matches.Recordset.Bookmark
        'End If
        'adodc_matches.Recordset.MoveNext
        'pula_owner, pula_address, pula_leg_band, pula_wing_band, pula_weight
        'puti_owner, puti_address, puti_leg_band, puti_wing_band, puti_weight
        With adodc_matches.Recordset
            If !pula_owner = search Or !pula_address = search Or !pula_leg_band = search Or !pula_wing_band = search Or !pula_weight = search Or !puti_owner = search Or !puti_address = search Or !puti_leg_band = search Or !puti_wing_band = search Or !puti_weight = search Then
                datagrid_matches.SelBookmarks.Add adodc_matches.Recordset.Bookmark
            End If
            .MoveNext
        End With
    Else
        timer_search.Enabled = False
    End If
End Sub
