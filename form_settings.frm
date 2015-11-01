VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form form_settings 
   Caption         =   "Settings"
   ClientHeight    =   9195
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9195
   ScaleWidth      =   13755
   Begin VB.Frame frame_category_type_settings 
      BackColor       =   &H00404040&
      Caption         =   "Participant Category Type Settings"
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
      Height          =   3255
      Left            =   720
      TabIndex        =   18
      Top             =   5400
      Width           =   5415
      Begin MSAdodcLib.Adodc adodc_participant_category_type 
         Height          =   375
         Left            =   240
         Top             =   1560
         Visible         =   0   'False
         Width           =   4935
         _ExtentX        =   8705
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
         RecordSource    =   "select participant_category from participant_category_master where participant_category <> 'None'"
         Caption         =   "category type"
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
      Begin VB.TextBox text_category 
         DataField       =   "participant_category"
         DataSource      =   "adodc_participant_category_type"
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
         Height          =   375
         Left            =   1320
         TabIndex        =   8
         Top             =   2160
         Width           =   3975
      End
      Begin VB.CommandButton command_participant_category_settings_save 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Save"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2640
         Width           =   1095
      End
      Begin VB.CommandButton command_participant_category_settings_delete 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Delete"
         Height          =   375
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2640
         Width           =   1095
      End
      Begin VB.CommandButton command_participant_category_settings_edit 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Edit"
         Height          =   375
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2640
         Width           =   1095
      End
      Begin VB.CommandButton command_participant_category_settings_add 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Add"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2640
         Width           =   1095
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "form_settings.frx":0000
         Height          =   1815
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   3201
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
         ColumnCount     =   1
         BeginProperty Column00 
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   1739.906
            EndProperty
         EndProperty
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Category:"
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
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   2160
         Width           =   1095
      End
   End
   Begin VB.Frame frame_event_type_settings 
      BackColor       =   &H00404040&
      Caption         =   "Event Type Settings"
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
      Height          =   4335
      Left            =   720
      TabIndex        =   13
      Top             =   720
      Width           =   5415
      Begin MSAdodcLib.Adodc adodc_event_type 
         Height          =   375
         Left            =   240
         Top             =   1560
         Visible         =   0   'False
         Width           =   4935
         _ExtentX        =   8705
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
         RecordSource    =   $"form_settings.frx":002E
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
      Begin VB.TextBox text_weight_range 
         DataField       =   "weight_range"
         DataSource      =   "adodc_event_type"
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
         Height          =   375
         Left            =   1800
         TabIndex        =   3
         Top             =   3360
         Width           =   3135
      End
      Begin VB.TextBox text_maximum_entries 
         DataField       =   "max_entries"
         DataSource      =   "adodc_event_type"
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
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   2760
         Width           =   3135
      End
      Begin VB.CommandButton command_event_type_settings_save 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Save"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3840
         Width           =   1215
      End
      Begin VB.CommandButton command_event_type_settings_delete 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Delete"
         Height          =   375
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3840
         Width           =   1215
      End
      Begin VB.CommandButton command_event_type_settings_edit 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Edit"
         Height          =   375
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3840
         Width           =   1215
      End
      Begin VB.CommandButton command_event_type_settings_add 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Add"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox text_event_type 
         DataField       =   "event_type"
         DataSource      =   "adodc_event_type"
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
         Height          =   375
         Left            =   1800
         TabIndex        =   1
         Top             =   2160
         Width           =   3135
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "form_settings.frx":00C7
         Height          =   1815
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   3201
         _Version        =   393216
         AllowUpdate     =   0   'False
         Enabled         =   -1  'True
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
         ColumnCount     =   3
         BeginProperty Column00 
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
         BeginProperty Column01 
            DataField       =   "max_entries"
            Caption         =   "max_entries"
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
            DataField       =   "weight_range"
            Caption         =   "weight_range"
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
               ColumnWidth     =   1635.024
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1170.142
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1530.142
            EndProperty
         EndProperty
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Weight Range:"
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
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Max. Entries"
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
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Event Type:"
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
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   2160
         Width           =   1575
      End
   End
   Begin VB.Line Line1 
      X1              =   720
      X2              =   13080
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   720
      TabIndex        =   0
      Top             =   0
      Width           =   12495
   End
End
Attribute VB_Name = "form_settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub event_type_settings_default()
    text_event_type.Enabled = False
    text_maximum_entries.Enabled = False
    text_weight_range.Enabled = False
    
    text_event_type.DataField = "event_type"
    text_maximum_entries.DataField = "max_entries"
    text_weight_range.DataField = "weight_range"
    
    command_event_type_settings_add.Enabled = True
    command_event_type_settings_edit.Enabled = True
    command_event_type_settings_delete.Enabled = True
    command_event_type_settings_save.Enabled = False
    
    command_event_type_settings_add.Caption = "Add"
    command_event_type_settings_edit.Caption = "Edit"
End Sub

Sub participant_category_settings_default()
    text_category.Enabled = False
    
    text_category.DataField = "participant_category"
    
    command_participant_category_settings_add.Enabled = True
    command_participant_category_settings_edit.Enabled = True
    command_participant_category_settings_delete.Enabled = True
    command_participant_category_settings_save.Enabled = False
    
    command_participant_category_settings_add.Caption = "Add"
    command_participant_category_settings_edit.Caption = "Edit"
End Sub

Private Sub command_event_type_settings_add_Click()
    If command_event_type_settings_add.Caption = "Add" Then
        command_event_type_settings_add.Caption = "Cancel"
        
        text_event_type.Enabled = True
        text_maximum_entries.Enabled = True
        text_weight_range.Enabled = True
        
        text_event_type.DataField = ""
        text_maximum_entries.DataField = ""
        text_weight_range.DataField = ""
        
        text_event_type.Text = ""
        text_maximum_entries.Text = ""
        text_weight_range.Text = ""
        
        command_event_type_settings_edit.Enabled = False
        command_event_type_settings_save.Enabled = True
        command_event_type_settings_delete.Enabled = False
        
        text_event_type.SetFocus
    ElseIf command_event_type_settings_add.Caption = "Cancel" Then
        Call event_type_settings_default
    End If
End Sub

Private Sub command_event_type_settings_delete_Click()
    If adodc_event_type.Recordset.RecordCount > 0 Then
        If MsgBox("Are you sure you want to delete this record?", vbYesNo) = vbYes Then
            adodc_event_type.Recordset.Delete
        Else
            Exit Sub
        End If
    Else
        MsgBox "No record found!", vbCritical
    End If
End Sub

Private Sub command_event_type_settings_edit_Click()
    With command_event_type_settings_edit
        If .Caption = "Edit" Then
            .Caption = "Cancel"
            command_event_type_settings_save.Enabled = True
            command_event_type_settings_add.Enabled = False
            command_event_type_settings_delete.Enabled = False
            
            text_event_type.DataField = ""
            text_maximum_entries.DataField = ""
            text_weight_range.DataField = ""
            
            text_event_type.Enabled = True
            text_maximum_entries.Enabled = True
            text_weight_range.Enabled = True
        ElseIf .Caption = "Cancel" Then
            Call event_type_settings_default
        End If
    End With
End Sub

Private Sub command_event_type_settings_save_Click()
    If text_event_type.Text = "" Or text_maximum_entries.Text = "" Or text_weight_range.Text = "" Then
        GoTo save_error
    End If
    
    If MsgBox("Are you sure you want to save this record?", vbYesNo) = vbNo Then
        Call event_type_settings_default
        Exit Sub
    End If
    
    If command_event_type_settings_add.Caption = "Cancel" Then
        'start saving
        With adodc_event_type
            .Recordset.AddNew
            .Recordset!event_type = text_event_type.Text
            .Recordset!max_entries = text_maximum_entries.Text
            .Recordset!weight_range = text_weight_range.Text
            On Error GoTo save_error
            .Recordset.Update
            MsgBox "Save successful!", vbInformation
            Call event_type_settings_default
            Exit Sub
            
save_error:
            MsgBox "Error Saving Data!", vbCritical
            Call event_type_settings_default
            Exit Sub
        End With
    ElseIf command_event_type_settings_edit.Caption = "Cancel" Then
        'start updating
            With adodc_event_type
            .Recordset!event_type = text_event_type.Text
            .Recordset!max_entries = text_maximum_entries.Text
            .Recordset!weight_range = text_weight_range.Text
            On Error GoTo save_error
            .Recordset.Update
            MsgBox "Save successful!", vbInformation
            Call event_type_settings_default
            Exit Sub
        End With
    End If
End Sub

Private Sub command_participant_category_settings_add_Click()
    With command_participant_category_settings_add
        If .Caption = "Add" Then
            .Caption = "Cancel"
            command_participant_category_settings_edit.Enabled = False
            command_participant_category_settings_delete.Enabled = False
            command_participant_category_settings_add.Enabled = True
             command_participant_category_settings_save.Enabled = True
            
            text_category.Enabled = True
            text_category.DataField = ""
            text_category.Text = ""
        ElseIf .Caption = "Cancel" Then
            Call participant_category_settings_default
            Exit Sub
        End If
    End With
End Sub

Private Sub command_participant_category_settings_delete_Click()
    If MsgBox("Are you sure you want to delete this record?", vbYesNo) = vbNo Then
        Exit Sub
    End If
    If adodc_participant_category_type.Recordset.RecordCount > 0 Then
        adodc_participant_category_type.Recordset.Delete
        MsgBox "Record deleted!", vbInformation
    Else
        MsgBox "No record found!", vbCritical
    End If
End Sub

Private Sub command_participant_category_settings_edit_Click()
    With command_participant_category_settings_edit
        If .Caption = "Edit" Then
            .Caption = "Cancel"
            
            text_category.DataField = ""
            text_category.Enabled = True
            
            command_participant_category_settings_add.Enabled = False
            command_participant_category_settings_delete.Enabled = False
             command_participant_category_settings_save.Enabled = True
        ElseIf .Caption = "Cancel" Then
            Call participant_category_settings_default
            Exit Sub
        End If
    End With
End Sub

Private Sub command_participant_category_settings_save_Click()
    If MsgBox("Are you sure you want to save this record?", vbYesNo) = vbNo Then
        Call participant_category_settings_default
        Exit Sub
    End If
    
    If text_category.Text = "" Then
        GoTo category_save_error
    End If
    
    If command_participant_category_settings_add.Caption = "Cancel" Then
        'start adding
        With adodc_participant_category_type
            .Recordset.AddNew
            .Recordset!participant_category = text_category.Text
            On Error GoTo category_save_error
            .Recordset.Update
            MsgBox "Saved Succesfuly!", vbInformation
            Call participant_category_settings_default
            Exit Sub
            
category_save_error:
            MsgBox "Error Saving Record!", vbCritical
            Call participant_category_settings_default
            Exit Sub
        End With
    ElseIf command_participant_category_settings_edit.Caption = "Cancel" Then
        'start updating
        With adodc_participant_category_type
            .Recordset!participant_category = text_category.Text
            On Error GoTo category_save_error
            .Recordset.Update
            MsgBox "Saved Succesfuly!", vbInformation
            Call participant_category_settings_default
            Exit Sub
        End With
    End If
End Sub
