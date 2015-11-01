VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm mdi_main 
   BackColor       =   &H8000000C&
   Caption         =   "BGT Cockfight Matching[Beta]"
   ClientHeight    =   8745
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   10905
   Icon            =   "mdi_main.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10905
      _ExtentX        =   19235
      _ExtentY        =   1164
      ButtonWidth     =   609
      ButtonHeight    =   1005
      Appearance      =   1
      _Version        =   393216
      Begin VB.CommandButton command_matches_2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Matches"
         Height          =   375
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   1935
      End
      Begin VB.CommandButton command_show_settings 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Settings"
         Height          =   375
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   2055
      End
      Begin VB.CommandButton command_show_new_event 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Events"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   2055
      End
   End
End
Attribute VB_Name = "mdi_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub command_show_dashboard_Click()
    Load form_dashboard
    form_dashboard.Show
    form_dashboard.WindowState = 2
End Sub

Private Sub command_matches_2_Click()
    Load form_matches_2
    form_matches_2.Show
    form_matches_2.WindowState = 2
End Sub

Private Sub command_show_matching_Click()
    Load form_matches
    form_matches.Show
    form_matches.WindowState = 2
End Sub

Private Sub command_show_new_event_Click()
    Load form_new_event
    form_new_event.Show
    form_new_event.WindowState = 2
End Sub

Private Sub command_show_settings_Click()
    Load form_settings
    form_settings.Show
    form_settings.WindowState = 2
End Sub

Private Sub MDIForm_Load()
    Call Connect
    Call Entry

    'Load form_dashboard
    'form_dashboard.Show
    
    Load form_new_event
    form_new_event.Show
End Sub
