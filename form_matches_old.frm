VERSION 5.00
Begin VB.Form form_matches 
   Caption         =   "Matches"
   ClientHeight    =   10125
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14835
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10125
   ScaleWidth      =   14835
   WindowState     =   2  'Maximized
   Begin VB.CommandButton command_print 
      Caption         =   "Print"
      Height          =   375
      Left            =   10680
      TabIndex        =   1
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton command_delete_matches 
      Caption         =   "Delete Matches"
      Height          =   375
      Left            =   8760
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton command_generate_matches 
      Caption         =   "Generate Matches"
      Height          =   375
      Left            =   6840
      TabIndex        =   4
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Select Event:"
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
      Left            =   720
      TabIndex        =   3
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
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
      Height          =   615
      Left            =   720
      TabIndex        =   0
      Top             =   0
      Width           =   14895
   End
End
Attribute VB_Name = "form_matches"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private sAppName As String, sAppPath As String

Dim entry_pula_id As Integer
Dim entry_puti_id As Integer
Dim event_id As Integer

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub adodc_matches_default()
    With adodc_matches
        .RecordSource = "Select * from matches_view_query order by match_schedule desc"
        .Refresh
    End With
End Sub

Sub mark_as_matched()
    With adodc_matcher
        .RecordSource = "select * from entries where entry_id = " & entry_pula_id & ""
        .Refresh
        
        .Recordset("entry_matching_status") = 2
        .Recordset.Update
        
        .RecordSource = "select * from entries where entry_id = " & entry_puti_id & ""
        .Refresh
        
        .Recordset("entry_matching_status") = 2
        .Recordset.Update
    End With
End Sub

Private Sub command_delete_matches_Click()
    With adodc_matches
        .RecordSource = "select * from matches"
        .Refresh
        
        While Not .Recordset.EOF
            .Recordset.Delete
            .Recordset.MoveNext
        Wend
        
        .RecordSource = "select * from entries where entry_matching_status = 2"
        .Refresh
        
        While Not .Recordset.EOF
            .Recordset("entry_matching_status") = 1
            .Recordset.Update
            .Recordset.MoveNext
        Wend
        
        Call adodc_matches_default
    End With
End Sub

Private Sub command_generate_matches_Click_draft1()
    If adodc_events.Recordset.EOF = False Then
        'for addnew variables
        Dim match_event_id As Integer
        Dim match_schedule As String
        
        Dim event_segments As Integer
        
        'for filtering variables
        Dim entry_weight As Integer
        Dim entry_category As Integer
        Dim entry_owner As Integer
        Dim entry_schedule_type As String
        
        
        match_event_id = adodc_events.Recordset("event_id") '@@@@@@@@@@@@@@@@@@@@@@ for addnew
        
        'firts know how many segments of matches this event will be
        'by dividing the total entries by two then dividing the result
        'by 30 <--30 stands as the default number of fights within a
        'day or a session of derby, this im not sure :)
        
        With adodc_matcher
            .RecordSource = "select * from entries_query where participant_event = " & match_event_id & ""
            .Refresh
            
            If .Recordset.RecordCount <> 0 Then
                Dim total_entries As Integer
                Dim total_possible_matches As Integer
                Dim total_segments As Integer
                Dim total_segments_proto As Double
                Dim total_segments_fixed As Integer
                Dim default_segment_matches As Integer
                Dim match_schedule_type As String
                
                match_schedule = 0
                default_segment_matches = 30
                total_entries = .Recordset.RecordCount
                total_possible_matches = total_entries / 2
                total_segments_proto = total_possible_matches / 30
                total_segments = total_possible_matches / 30
                
                If total_segments < total_segments_proto Then
                    total_segments = total_segments + 1
                    total_segments_fixed = total_segments
                End If
                
                While total_segments > 0
                    default_segment_matches = total_possible_matches - match_schedule
                    If total_segments = total_segments_fixed Then
                        match_schedule_type = "Early Fight"
                    ElseIf total_segments = 1 Then
                        match_schedule_type = "Late Fight"
                    Else
                        match_schedule_type = "None"
                    End If
                
                    While default_segment_matches > 0
                        'get first entry, filters
                        .RecordSource = "select * from entries_query where participant_event = " & match_event_id & " and entry_schedule_type = '" & match_schedule_type & "' and participant_category = 1 and entry_matching_status = 'Unmatched'"
                        .Refresh
                        
                            'check if this participant is qualified or ready for this fight
                            'by checking if this participant have a fight that is within
                            '5 records from this row
                            '!!!!!!!!!
                                If .Recordset.RecordCount <> 0 Then
                                    'first entry found for pula
pula_found:
                                    
                                    entry_pula_id = .Recordset("entry_id") '@@@@@@@@@@@@@@@@@@@@@@ for addnew
                                    entry_weight = .Recordset("entry_weight")
                                    entry_category = .Recordset("participant_category")
                                    entry_schedule_type = .Recordset("entry_schedule_type")
                                    
                                    'get second entry for puti
                                    .RecordSource = "select top 1 * from entries_query where (entry_weight >= " & entry_weight & " or entry_weight <= " & entry_weight & ") and participant_category = " & entry_category & " and entry_schedule_type = '" & match_schedule_type & "' and entry_matching_status = 'Unmatched' and participant_event = " & match_event_id & " and entry_id <> " & entry_pula_id & ""
                                    .Refresh
                                    
                                    If .Recordset.RecordCount <> 0 Then
                                        'second entry found for puti
puti_found:
                                        
                                        match_schedule = match_schedule + 1 '@@@@@@@@@@@@@@@@@@@@@@ for addnew
                                        entry_puti_id = .Recordset("entry_id") '@@@@@@@@@@@@@@@@@@@@@@ for addnew
                                        Dim puti_weight As Integer
                                        puti_weight = .Recordset("entry_weight")
                                        .RecordSource = "select * from matches"
                                        .Refresh
                                        'start adding
                                        .Recordset.AddNew
                                        .Recordset("match_event_id") = match_event_id
                                        .Recordset("match_schedule") = match_schedule
                                        .Recordset("entry_pula_id") = entry_pula_id
                                        .Recordset("entry_puti_id") = entry_puti_id
                                        .Recordset.Update
                                        
                                        default_segment_matches = default_segment_matches - 1
                                        
                                        Call mark_as_matched
                                        
                                        Call adodc_matches_default
                                    Else
                                        'run less sensitive filter query to get second entry
                                        .RecordSource = "select top 1 * from entries_query where (entry_weight >= " & entry_weight & " or entry_weight <= " & entry_weight & ") and entry_schedule_type = '" & match_schedule_type & "' and participant_category = " & entry_category & " and entry_matching_status = 'Unmatched' and participant_event = " & match_event_id & " and entry_id <> " & entry_pula_id & ""
                                        .Refresh
                                        
                                        If .Recordset.RecordCount <> 0 Then
                                            GoTo puti_found
                                        Else
                                            'run less sensitive filter query to get second entry
                                            .RecordSource = "select top 1 * from entries_query where (entry_weight >= " & entry_weight & " or entry_weight <= " & entry_weight & ") and entry_schedule_type = '" & match_schedule_type & "' and entry_matching_status = 'Unmatched' and participant_event = " & match_event_id & " and entry_id <> " & entry_pula_id & ""
                                            .Refresh
                                            
                                            If .Recordset.RecordCount <> 0 Then
                                                GoTo puti_found
                                            Else
                                                'run less sensitive filter query to get second entry
                                                .RecordSource = "select top 1 * from entries_query where (entry_weight >= " & entry_weight & " or entry_weight <= " & entry_weight & ") and entry_schedule_type = '" & match_schedule_type & "' and entry_matching_status = 'Unmatched' and entry_id <> " & entry_pula_id & ""
                                                .Refresh
                                                
                                                If .Recordset.RecordCount <> 0 Then
                                                    GoTo puti_found
                                                Else
                                                    default_segment_matches = default_segment_matches - 1
                                                            
                                                    Call adodc_matches_default
                                                End If
                                            End If
                                        End If
                                    End If
                                Else
                                    'run less sensitive filer query to get first entry
                                    
                                    .RecordSource = "select * from entries_query where participant_event = " & match_event_id & " and entry_schedule_type = '" & match_schedule_type & "' and entry_matching_status = 'Unmatched'"
                                    .Refresh
                                    
                                    If .Recordset.RecordCount <> 0 Then
                                        GoTo pula_found
                                    Else
                                        'run less sensitive filter query to get first entry
                                        .RecordSource = "select * from entries_query where participant_event = " & match_event_id & " and entry_matching_status = 'Unmatched'"
                                        .Refresh
                                        
                                        If .Recordset.RecordCount <> 0 Then
                                            GoTo pula_found
                                        Else
                                            default_segment_matches = default_segment_matches - 1
                                        End If
                                    End If
                                End If
                        Wend
                    total_segments = total_segments - 1
                Wend
                
            End If
        End With
    End If
End Sub

Private Sub command_generate_matches_Click()
    Dim total_entries As Integer
    Dim total_possible_matches As Integer
    Dim total_segments As Integer
    Dim total_segments_proto As Double
    Dim total_segments_fixed As Integer
    Dim default_segment_matches As Integer
    Dim fixed_default_segment_matches As Integer
    Dim match_schedule_type As String
    Dim segment_sections As Integer
    Dim fixed_segment_sections As Integer
    Dim match_schedule As Integer
    Dim last_segment_section As Integer

    Dim pula_id As Integer
    Dim puti_id As Integer
    Dim puti_participant_id As Integer
    Dim puti_weight_difference As Integer
    Dim puti_participant_bet As Currency
    Dim puti_bet_difference As Currency
    Dim bet_difference As Currency
    
    Dim user_ready As Boolean
    
    'get total entries to know how many matches or matchups are possible
    With adodc_entries
        .RecordSource = "select * from entries_query where participant_event = " & adodc_events.Recordset("event_id") & ""
        .Refresh
        total_entries = .Recordset.RecordCount
        
        If .Recordset.RecordCount <> 0 Or .Recordset.RecordCount = 1 Then
            match_schedule = 0
            default_segment_matches = 30
            fixed_default_segment_matches = default_segment_matches
            fixed_segment_sections = default_segment_matches / 3 'first 1/3 early_fights, second 1/3 none, third 1/3 late_fights
            total_entries = .Recordset.RecordCount
            total_possible_matches = total_entries / 2
            total_segments_proto = total_possible_matches / default_segment_matches
            total_segments = total_possible_matches / default_segment_matches
            last_segment_section = fixed_default_segment_matches - (fixed_segment_sections * 2)
            
            adodc_matcher.RecordSource = "Select * from process"
            adodc_matcher.Refresh
            
            adodc_matcher.Recordset!current_event_id = adodc_events.Recordset!event_id
            adodc_matcher.Recordset!current_event_name = adodc_events.Recordset!event_name
            adodc_matcher.Recordset!current_total = total_possible_matches
            adodc_matcher.Recordset.Update
            
            If total_segments < total_segments_proto Then
                total_segments = total_segments + 1
                total_segments_fixed = total_segments
            End If
            
        mdi_main.Visible = False
            
        res = Shell("splasher.exe " & sAppPath, vbHide)
        Else
            MsgBox "This event has no participants!", vbCritical
            GoTo exit_now
        End If
        
    End With
    While total_segments > 0
        While default_segment_matches > 0
            segment_sections = fixed_segment_sections 'reset segment sections
change_match_schedule_type:
            'determine the current matching schedule type/preferrence
            If default_segment_matches = fixed_default_segment_matches Then
                match_schedule_type = "Early Fight"
            ElseIf default_segment_matches = last_segment_section Then
                match_schedule_type = "Late Fight"
            Else
                match_schedule_type = "None"
            End If
            'by this point the match_schedule_fight should be identified
            
            segment_sections = fixed_segment_sections
            While segment_sections > 0
                adodc_entries.RecordSource = "Select * from entries_query " & _
                " where participant_event = " & adodc_events.Recordset!event_id & "" & _
                " and entry_matching_status = 'Unmatched'" & _
                " and entry_schedule_type = '" & match_schedule_type & "' order by participant_category"
                adodc_entries.Refresh
                
                If adodc_entries.Recordset.RecordCount > 0 Then
new_first_entry:

                    

                    'check if this entry is already matched
                    adodc_matcher.RecordSource = "select * from entries_query where entry_id = " & adodc_entries.Recordset!entry_id & ""
                    Sleep 500
                    adodc_matcher.Refresh
                    
                    'check if this entry_owner is ready for this fightslot
                    adodc_matcher2.RecordSource = " select top 5 * from matches_view_query where match_event_id = " & adodc_events.Recordset!event_id & " order by match_id desc"
                    adodc_matcher2.Refresh
                    
                    user_ready = True
                    If adodc_matcher2.Recordset.RecordCount > 0 Then
                        adodc_matcher2.Recordset.MoveFirst
                        While Not adodc_matcher2.Recordset.EOF
                            If adodc_matcher2.Recordset("entries_query_for_pula.entry_owner") = adodc_entries.Recordset!entry_owner Or adodc_matcher2.Recordset("entries_query_for_puti.entry_owner") = adodc_entries.Recordset!entry_owner Then
                                user_ready = False
                                GoTo exit_while_1
                            Else
                                adodc_matcher2.Recordset.MoveNext
                            End If
                        Wend
                    End If
exit_while_1:
                    
                    If adodc_matcher.Recordset!entry_matching_status = "Matched" Or user_ready = False Then
                        GoTo entries_movenext
                    End If
                    
                    If adodc_entries.Recordset!entry_matching_status = "Unmatched" Then
    
    
                        'search match using first criteria
                        adodc_matcher.RecordSource = "Select entry_weight, entry_id, entry_owner, participant_bet, " & _
                        " (entry_weight - " & adodc_entries.Recordset!entry_weight & ") " & _
                        " as weight_difference from entries_query " & _
                        " where participant_event = " & adodc_events.Recordset!event_id & "" & _
                        " and entry_matching_status = 'Unmatched' " & _
                        " and entry_schedule_type = '" & match_schedule_type & "'" & _
                        " and participant_category = " & adodc_entries.Recordset!participant_category & "" & _
                        " and entry_weight >= " & adodc_entries.Recordset!entry_weight & "" & _
                        " and entry_id <> " & adodc_entries.Recordset!entry_id & "" & _
                        " and entry_owner <> " & adodc_entries.Recordset!entry_owner & "" & _
                        " union select entry_weight, entry_id, entry_owner, participant_bet, (" & adodc_entries.Recordset!entry_weight & " - entry_weight) as weight_difference from entries_query " & _
                        " where participant_event = " & adodc_events.Recordset!event_id & "" & _
                        " and entry_matching_status = 'Unmatched' " & _
                        " and entry_schedule_type = '" & match_schedule_type & "'" & _
                        " and participant_category = " & adodc_entries.Recordset!participant_category & "" & _
                        " and entry_weight < " & adodc_entries.Recordset!entry_weight & "" & _
                        " and entry_owner <> " & adodc_entries.Recordset!entry_owner & "" & _
                        " and entry_id <> " & adodc_entries.Recordset!entry_id & ""
                        adodc_matcher.Refresh
                        adodc_matcher.Recordset.Sort = "weight_difference"
                        
                        
                        While Not adodc_entries.Recordset.EOF And Not adodc_matcher.Recordset.EOF
                            
                            If adodc_matcher.Recordset.RecordCount > 0 Or adodc_matcher.Recordset.EOF = False Then
                                adodc_matcher.Recordset.MoveFirst
next_matcher:
                                If adodc_matcher.Recordset.EOF = True Then
                                    GoTo entries_movenext
                                End If
                                If adodc_matcher.Recordset!weight_difference < 50 Then
                                    
                                    pula_id = adodc_entries.Recordset!entry_id
                                    puti_id = adodc_matcher.Recordset!entry_id
                                    puti_participant_id = adodc_matcher.Recordset!entry_owner
                                    puti_weight_difference = adodc_matcher.Recordset!weight_difference
                                    puti_participant_bet = adodc_matcher.Recordset!participant_bet
                                    
                                    'get bet difference
                                    If puti_participant_bet > adodc_entries.Recordset!participant_bet Then
                                        puti_bet_difference = puti_participant_bet - adodc_entries.Recordset!participant_bet
                                    Else
                                        puti_bet_difference = adodc_entries.Recordset!participant_bet - puti_participant_bet
                                    End If
                                    
check_no_match:
                                    'check no_match
                                    adodc_matcher2.RecordSource = "Select * from participant_no_match_union_view_query " & _
                                    " where no_match_id = " & puti_participant_id & " " & _
                                    " and participant_id = " & adodc_entries.Recordset!entry_id & "" & _
                                    " and event_id = " & adodc_events.Recordset!event_id & ""
                                    adodc_matcher2.Refresh
                                    
                                    If adodc_matcher2.Recordset.RecordCount > 0 Then
new_matcher:
                                        If Not adodc_matcher.Recordset.EOF Then
                                            adodc_matcher.Recordset.MoveNext
                                            GoTo next_matcher
                                        ElseIf Not adodc_entries.Recordset.EOF Then
                                            adodc_entries.Recordset.MoveNext
                                            GoTo new_first_entry
                                        End If
                                    Else
                                        'check if entry_owner ready for this fightslot
                                        adodc_matcher2.RecordSource = " select top 5 * from matches_view_query where match_event_id = " & adodc_events.Recordset!event_id & " order by match_id desc"
                                        adodc_matcher2.Refresh
                                        
                                        
                                        user_ready = True
                                        If adodc_matcher2.Recordset.RecordCount > 0 Then
                                            adodc_matcher2.Recordset.MoveFirst
                                            While Not adodc_matcher2.Recordset.EOF
                                                If adodc_matcher2.Recordset("entries_query_for_pula.entry_owner") = adodc_matcher.Recordset!entry_owner Or adodc_matcher2.Recordset("entries_query_for_puti.entry_owner") = adodc_matcher.Recordset!entry_owner Then
                                                    user_ready = False
                                                    GoTo exit_while
                                                Else
                                                    adodc_matcher2.Recordset.MoveNext
                                                End If
                                            Wend
                                        End If
exit_while:
                                        If user_ready = False Then
                                            GoTo new_matcher
                                        Else
                                            While Not adodc_matcher.Recordset.EOF
                                                If adodc_matcher.Recordset!weight_difference < 50 Then
                                                    
                                                    If adodc_matcher.Recordset!participant_bet >= adodc_entries.Recordset!participant_bet Then
                                                        bet_difference = adodc_matcher.Recordset!participant_bet - adodc_entries.Recordset!participant_bet
                                                    Else
                                                        bet_difference = adodc_entries.Recordset!participant_bet - adodc_matcher.Recordset!participant_bet
                                                    End If
                                                    
                                                    If bet_difference < puti_bet_difference Then
                                                        'puti_bet_difference defeated, assign new variables
                                                        puti_id = adodc_matcher.Recordset!entry_id
                                                        adodc_matcher.Recordset.MoveNext
                                                    Else
                                                        adodc_matcher.Recordset.MoveNext
                                                    End If
                                                Else
                                                    adodc_matcher.Recordset.MoveNext
                                                End If
                                            Wend
                                            
                                            'start saving match record to database
                                            event_id = adodc_events.Recordset!event_id
                                            match_schedule = match_schedule + 1 'for fights numbering, or sequencing
                                            segment_sections = segment_sections - 1
                                            
                                            entry_pula_id = adodc_entries.Recordset!entry_id
                                            entry_puti_id = puti_id
                                            
                                            adodc_matches.RecordSource = "select * from matches"
                                            adodc_matches.Refresh
                                            
                                            With adodc_matches
                                                .Recordset.AddNew
                                                .Recordset!match_event_id = event_id
                                                .Recordset!match_schedule = match_schedule
                                                .Recordset!entry_pula_id = entry_pula_id
                                                .Recordset!entry_puti_id = entry_puti_id
                                                .Recordset.Update
                                            End With
                                            
                                            'mark matched entries as Matched
                                            adodc_matcher2.RecordSource = "select * from entries where entry_id = " & entry_pula_id & ""
                                            adodc_matcher2.Refresh
                                            adodc_matcher2.Recordset!entry_matching_status = 2
                                            adodc_matcher2.Recordset.Update
                                            
                                            adodc_matcher2.RecordSource = "select * from entries where entry_id = " & entry_puti_id & ""
                                            adodc_matcher2.Refresh
                                            adodc_matcher2.Recordset!entry_matching_status = 2
                                            adodc_matcher2.Recordset.Update
                                            adodc_entries.Recordset.MoveNext
                                            
                                            adodc_matches.Refresh
                                            datagrid_matches.Refresh
                                            
                                            If default_segment_matches = fixed_segment_sections Then
                                                GoTo change_match_schedule_type
                                            ElseIf default_segment_matches = last_segment_section Then
                                                GoTo change_match_schedule_type
                                            Else
                                                default_segment_matches = default_segment_matches - 1
                                            End If
                                            
                                            If Not adodc_entries.Recordset.EOF And segment_sections <> 0 Then
                                                GoTo new_first_entry
                                            Else
                                                'start matching using second criteria
                                            End If
                                        End If
                                    End If
                                Else
entries_movenext:
                                    adodc_entries.Recordset.MoveNext
                                    If adodc_entries.Recordset.EOF = False Then
                                        GoTo new_first_entry
                                    Else
                                        segment_sections = segment_sections - 1
                                    End If
                                End If
                            Else
                            End If
                        Wend
                        
                        segment_sections = segment_sections - 1
                    Else
                        adodc_entries.Recordset.MoveNext
                        If adodc_entries.Recordset.EOF = False Then
                            GoTo new_first_entry
                        Else
                            segment_sections = segment_sections - 1
                        End If
                    End If
                    
                Else
                 segment_sections = segment_sections - 1
                End If
                'segment_sections = segment_sections - 1
                
                
'!!!!!!!!!!!!!!!!!!!!!!!!!!!ROUND 2!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

                adodc_entries.RecordSource = "Select * from entries_query " & _
                " where participant_event = " & adodc_events.Recordset!event_id & "" & _
                " and entry_matching_status = 'Unmatched'" & _
                " and entry_schedule_type = '" & match_schedule_type & "' order by participant_category"
                adodc_entries.Refresh
                
                If adodc_entries.Recordset.RecordCount > 0 Then
new_first_entry2:

                    'check if this entry is already matched
                    adodc_matcher.RecordSource = "select * from entries_query where entry_id = " & adodc_entries.Recordset!entry_id & ""
                    Sleep 500
                    adodc_matcher.Refresh
                    
                    'check if this entry_owner is ready for this fightslot
                    adodc_matcher2.RecordSource = " select top 5 * from matches_view_query where match_event_id = " & adodc_events.Recordset!event_id & " order by match_id desc"
                    adodc_matcher2.Refresh
                    
                    user_ready = True
                    If adodc_matcher2.Recordset.RecordCount > 0 Then
                        adodc_matcher2.Recordset.MoveFirst
                        While Not adodc_matcher2.Recordset.EOF
                            If adodc_matcher2.Recordset("entries_query_for_pula.entry_owner") = adodc_entries.Recordset!entry_owner Or adodc_matcher2.Recordset("entries_query_for_puti.entry_owner") = adodc_entries.Recordset!entry_owner Then
                                user_ready = False
                                GoTo exit_while_12
                            Else
                                adodc_matcher2.Recordset.MoveNext
                            End If
                        Wend
                    End If
exit_while_12:
                    
                    If adodc_matcher.Recordset!entry_matching_status = "Matched" Or user_ready = False Then
                        GoTo entries_movenext2
                    End If
                    
                    If adodc_entries.Recordset!entry_matching_status = "Unmatched" Then
                        'search match using this criteria
                        adodc_matcher.RecordSource = "Select entry_weight, entry_id, entry_owner, participant_bet, " & _
                        " (entry_weight - " & adodc_entries.Recordset!entry_weight & ") " & _
                        " as weight_difference from entries_query " & _
                        " where participant_event = " & adodc_events.Recordset!event_id & "" & _
                        " and entry_matching_status = 'Unmatched' " & _
                        " and entry_schedule_type = '" & match_schedule_type & "'" & _
                        " and entry_weight >= " & adodc_entries.Recordset!entry_weight & "" & _
                        " and entry_id <> " & adodc_entries.Recordset!entry_id & "" & _
                        " and entry_owner <> " & adodc_entries.Recordset!entry_owner & "" & _
                        " union select entry_weight, entry_id, entry_owner, participant_bet, (" & adodc_entries.Recordset!entry_weight & " - entry_weight) as weight_difference from entries_query " & _
                        " where participant_event = " & adodc_events.Recordset!event_id & "" & _
                        " and entry_matching_status = 'Unmatched' " & _
                        " and entry_schedule_type = '" & match_schedule_type & "'" & _
                        " and entry_weight < " & adodc_entries.Recordset!entry_weight & "" & _
                        " and entry_owner <> " & adodc_entries.Recordset!entry_owner & "" & _
                        " and entry_id <> " & adodc_entries.Recordset!entry_id & ""
                        adodc_matcher.Refresh
                        adodc_matcher.Recordset.Sort = "weight_difference"
                        
                        
                        While Not adodc_entries.Recordset.EOF
                            
                            If adodc_matcher.Recordset.RecordCount > 0 Or adodc_matcher.Recordset.EOF = False Then
                                adodc_matcher.Recordset.MoveFirst
next_matcher2:
                                
                                If adodc_matcher.Recordset.EOF = True Then
                                    GoTo entries_movenext2
                                End If
                                
                                If adodc_matcher.Recordset!weight_difference < 50 Then
    
                                    
                                    pula_id = adodc_entries.Recordset!entry_id
                                    puti_id = adodc_matcher.Recordset!entry_id
                                    puti_participant_id = adodc_matcher.Recordset!entry_owner
                                    puti_weight_difference = adodc_matcher.Recordset!weight_difference
                                    puti_participant_bet = adodc_matcher.Recordset!participant_bet
                                    
                                    'get bet difference
                                    If puti_participant_bet > adodc_entries.Recordset!participant_bet Then
                                        puti_bet_difference = puti_participant_bet - adodc_entries.Recordset!participant_bet
                                    Else
                                        puti_bet_difference = adodc_entries.Recordset!participant_bet - puti_participant_bet
                                    End If
                                    
check_no_match2:
                                    'check no_match
                                    adodc_matcher2.RecordSource = "Select * from participant_no_match_union_view_query " & _
                                    " where no_match_id = " & puti_participant_id & " " & _
                                    " and participant_id = " & adodc_entries.Recordset!entry_id & "" & _
                                    " and event_id = " & adodc_events.Recordset!event_id & ""
                                    adodc_matcher2.Refresh
                                    
                                    If adodc_matcher2.Recordset.RecordCount > 0 Then
new_matcher2:
                                        If Not adodc_matcher.Recordset.EOF Then
                                            adodc_matcher.Recordset.MoveNext
                                            GoTo next_matcher2
                                        ElseIf Not adodc_entries.Recordset.EOF Then
                                            adodc_entries.Recordset.MoveNext
                                            GoTo new_first_entry2
                                        End If
                                    Else
                                        'check if entry_owner ready for this fightslot
                                        adodc_matcher2.RecordSource = " select top 5 * from matches_view_query where match_event_id = " & adodc_events.Recordset!event_id & " order by match_id desc"
                                        adodc_matcher2.Refresh
                                        
                                        user_ready = True
                                        If adodc_matcher2.Recordset.RecordCount > 0 Then
                                            adodc_matcher2.Recordset.MoveFirst
                                            While Not adodc_matcher2.Recordset.EOF
                                                If adodc_matcher2.Recordset("entries_query_for_pula.entry_owner") = adodc_matcher.Recordset!entry_owner Or adodc_matcher2.Recordset("entries_query_for_puti.entry_owner") = adodc_matcher.Recordset!entry_owner Then
                                                    user_ready = False
                                                    GoTo exit_while2
                                                Else
                                                    adodc_matcher2.Recordset.MoveNext
                                                End If
                                            Wend
                                        End If
exit_while2:
                                        
                                        If user_ready = False Then
                                            GoTo new_matcher2
                                        Else
                                            While Not adodc_matcher.Recordset.EOF
                                                If adodc_matcher.Recordset!weight_difference < 50 Then
                                                    
                                                    If adodc_matcher.Recordset!participant_bet >= adodc_entries.Recordset!participant_bet Then
                                                        bet_difference = adodc_matcher.Recordset!participant_bet - adodc_entries.Recordset!participant_bet
                                                    Else
                                                        bet_difference = adodc_entries.Recordset!participant_bet - adodc_matcher.Recordset!participant_bet
                                                    End If
                                                    
                                                    If bet_difference < puti_bet_difference Then
                                                        'puti_bet_difference defeated, assign new variables
                                                        puti_id = adodc_matcher.Recordset!entry_id
                                                        adodc_matcher.Recordset.MoveNext
                                                    Else
                                                        adodc_matcher.Recordset.MoveNext
                                                    End If
                                                Else
                                                    adodc_matcher.Recordset.MoveNext
                                                End If
                                            Wend
                                            
                                            'start saving match record to database
                                            event_id = adodc_events.Recordset!event_id
                                            match_schedule = match_schedule + 1 'for fights numbering, or sequencing
                                            segment_sections = segment_sections - 1
                                            
                                            entry_pula_id = adodc_entries.Recordset!entry_id
                                            entry_puti_id = puti_id
                                            
                                            adodc_matches.RecordSource = "select * from matches"
                                            adodc_matches.Refresh
                                            
                                            With adodc_matches
                                                .Recordset.AddNew
                                                .Recordset!match_event_id = event_id
                                                .Recordset!match_schedule = match_schedule
                                                .Recordset!entry_pula_id = entry_pula_id
                                                .Recordset!entry_puti_id = entry_puti_id
                                                .Recordset.Update
                                            End With
                                            
                                            'mark matched entries as Matched
                                            adodc_matcher2.RecordSource = "select * from entries where entry_id = " & entry_pula_id & ""
                                            adodc_matcher2.Refresh
                                            adodc_matcher2.Recordset!entry_matching_status = 2
                                            adodc_matcher2.Recordset.Update
                                            
                                            adodc_matcher2.RecordSource = "select * from entries where entry_id = " & entry_puti_id & ""
                                            adodc_matcher2.Refresh
                                            adodc_matcher2.Recordset!entry_matching_status = 2
                                            adodc_matcher2.Recordset.Update
                                            adodc_entries.Recordset.MoveNext
                                            
                                            adodc_matches.Refresh
                                            datagrid_matches.Refresh
                                            
                                            If default_segment_matches = fixed_segment_sections Then
                                                GoTo change_match_schedule_type
                                            ElseIf default_segment_matches = last_segment_section Then
                                                GoTo change_match_schedule_type
                                            Else
                                                default_segment_matches = default_segment_matches - 1
                                            End If
                                            
                                            If Not adodc_entries.Recordset.EOF And segment_sections <> 0 Then
                                                GoTo new_first_entry2
                                            Else
                                                segment_sections = segment_sections - 1
                                            End If
                                        End If
                                    End If
                                Else
entries_movenext2:
                                    adodc_entries.Recordset.MoveNext
                                    If adodc_entries.Recordset.EOF = False Then
                                        GoTo new_first_entry2
                                    Else
                                        segment_sections = segment_sections - 1
                                    End If
                                End If
                            Else
                                adodc_entries.Recordset.MoveNext
                                If adodc_entries.Recordset.EOF = False Then
                                    GoTo new_first_entry2
                                Else
                                    segment_sections = segment_sections - 1
                                End If
                            End If
                        Wend
                        
                    'search match using third criteria
                    Else
                        adodc_entries.Recordset.MoveNext
                        If adodc_entries.Recordset.EOF = False Then
                            GoTo new_first_entry2
                        Else
                            'start matching using second criteria
                        End If
                    End If
                        
                Else
                 segment_sections = segment_sections - 1
                End If
                'segment_sections = segment_sections - 1
                
'!!!!!!!!!!!!!!!!!ROUND 3 3RD CRITERIA!!!!!!!!!!!!!!!!!!!!!!!!!

                adodc_entries.RecordSource = "Select * from entries_query " & _
                " where participant_event = " & adodc_events.Recordset!event_id & "" & _
                " and entry_matching_status = 'Unmatched'" & _
                " order by participant_category"
                adodc_entries.Refresh
                
                If adodc_entries.Recordset.RecordCount > 0 Then
new_first_entry3:

                    'check if this entry is already matched
                    adodc_matcher.RecordSource = "select * from entries_query where entry_id = " & adodc_entries.Recordset!entry_id & ""
                    Sleep 500
                    adodc_matcher.Refresh
                    
                    'check if this entry_owner is ready for this fightslot
                    adodc_matcher2.RecordSource = " select top 5 * from matches_view_query where match_event_id = " & adodc_events.Recordset!event_id & " order by match_id desc"
                    adodc_matcher2.Refresh
                    
                    user_ready = True
                    If adodc_matcher2.Recordset.RecordCount > 0 Then
                        adodc_matcher2.Recordset.MoveFirst
                        While Not adodc_matcher2.Recordset.EOF
                            If adodc_matcher2.Recordset("entries_query_for_pula.entry_owner") = adodc_entries.Recordset!entry_owner Or adodc_matcher2.Recordset("entries_query_for_puti.entry_owner") = adodc_entries.Recordset!entry_owner Then
                                user_ready = False
                                GoTo exit_while_13
                            Else
                                adodc_matcher2.Recordset.MoveNext
                            End If
                        Wend
                    End If
exit_while_13:
                    
                    If adodc_matcher.Recordset!entry_matching_status = "Matched" Or user_ready = False Then
                        GoTo entries_movenext3
                    End If
                    
                    If adodc_entries.Recordset!entry_matching_status = "Unmatched" Then
                        'search match using first criteria
                        adodc_matcher.RecordSource = "Select entry_weight, entry_id, entry_owner, participant_bet, " & _
                        " (entry_weight - " & adodc_entries.Recordset!entry_weight & ") " & _
                        " as weight_difference from entries_query " & _
                        " where participant_event = " & adodc_events.Recordset!event_id & "" & _
                        " and entry_matching_status = 'Unmatched' " & _
                        " and entry_weight >= " & adodc_entries.Recordset!entry_weight & "" & _
                        " and entry_id <> " & adodc_entries.Recordset!entry_id & "" & _
                        " and entry_owner <> " & adodc_entries.Recordset!entry_owner & "" & _
                        " union select entry_weight, entry_id, entry_owner, participant_bet, (" & adodc_entries.Recordset!entry_weight & " - entry_weight) as weight_difference from entries_query " & _
                        " where participant_event = " & adodc_events.Recordset!event_id & "" & _
                        " and entry_matching_status = 'Unmatched' " & _
                        " and entry_weight < " & adodc_entries.Recordset!entry_weight & "" & _
                        " and entry_owner <> " & adodc_entries.Recordset!entry_owner & "" & _
                        " and entry_id <> " & adodc_entries.Recordset!entry_id & ""
                        adodc_matcher.Refresh
                        adodc_matcher.Recordset.Sort = "weight_difference"
                        
                        
                        While Not adodc_entries.Recordset.EOF
                            
                            If adodc_matcher.Recordset.RecordCount > 0 Or adodc_matcher.Recordset.EOF = False Then
                                adodc_matcher.Recordset.MoveFirst
next_matcher3:
                                If adodc_matcher.Recordset.EOF = True Then
                                    GoTo entries_movenext3
                                End If
                                
                                If adodc_matcher.Recordset!weight_difference < 50 Then
    
                                    
                                    pula_id = adodc_entries.Recordset!entry_id
                                    puti_id = adodc_matcher.Recordset!entry_id
                                    puti_participant_id = adodc_matcher.Recordset!entry_owner
                                    puti_weight_difference = adodc_matcher.Recordset!weight_difference
                                    puti_participant_bet = adodc_matcher.Recordset!participant_bet
                                    
                                    'get bet difference
                                    If puti_participant_bet > adodc_entries.Recordset!participant_bet Then
                                        puti_bet_difference = puti_participant_bet - adodc_entries.Recordset!participant_bet
                                    Else
                                        puti_bet_difference = adodc_entries.Recordset!participant_bet - puti_participant_bet
                                    End If
                                    
check_no_match3:
                                    'check no_match
                                    adodc_matcher2.RecordSource = "Select * from participant_no_match_union_view_query " & _
                                    " where no_match_id = " & puti_participant_id & " " & _
                                    " and participant_id = " & adodc_entries.Recordset!entry_id & "" & _
                                    " and event_id = " & adodc_events.Recordset!event_id & ""
                                    adodc_matcher2.Refresh
                                    
                                    If adodc_matcher2.Recordset.RecordCount > 0 Then
new_matcher3:
                                        If Not adodc_matcher.Recordset.EOF Then
                                            adodc_matcher.Recordset.MoveNext
                                            GoTo next_matcher3
                                        ElseIf Not adodc_entries.Recordset.EOF Then
                                            adodc_entries.Recordset.MoveNext
                                            GoTo new_first_entry3
                                        End If
                                    Else
                                        'check if entry_owner ready for this fightslot
                                        adodc_matcher2.RecordSource = " select top 5 * from matches_view_query where match_event_id = " & adodc_events.Recordset!event_id & " order by match_id desc"
                                        adodc_matcher2.Refresh
                                        
                                        
                                        user_ready = True
                                        If adodc_matcher2.Recordset.RecordCount > 0 Then
                                            adodc_matcher2.Recordset.MoveFirst
                                            While Not adodc_matcher2.Recordset.EOF
                                                If adodc_matcher2.Recordset("entries_query_for_pula.entry_owner") = adodc_matcher.Recordset!entry_owner Or adodc_matcher2.Recordset("entries_query_for_puti.entry_owner") = adodc_matcher.Recordset!entry_owner Then
                                                    user_ready = False
                                                    GoTo exit_while3
                                                Else
                                                    adodc_matcher2.Recordset.MoveNext
                                                End If
                                            Wend
                                        End If
exit_while3:
                                        
                                        If user_ready = False Then
                                            GoTo new_matcher3
                                        Else
                                            While Not adodc_matcher.Recordset.EOF
                                                If adodc_matcher.Recordset!weight_difference < 50 Then
                                                    
                                                    If adodc_matcher.Recordset!participant_bet >= adodc_entries.Recordset!participant_bet Then
                                                        bet_difference = adodc_matcher.Recordset!participant_bet - adodc_entries.Recordset!participant_bet
                                                    Else
                                                        bet_difference = adodc_entries.Recordset!participant_bet - adodc_matcher.Recordset!participant_bet
                                                    End If
                                                    
                                                    If bet_difference < puti_bet_difference Then
                                                        'puti_bet_difference defeated, assign new variables
                                                        puti_id = adodc_matcher.Recordset!entry_id
                                                        adodc_matcher.Recordset.MoveNext
                                                    Else
                                                        adodc_matcher.Recordset.MoveNext
                                                    End If
                                                Else
                                                    adodc_matcher.Recordset.MoveNext
                                                End If
                                            Wend
                                            
                                            'start saving match record to database
                                            event_id = adodc_events.Recordset!event_id
                                            match_schedule = match_schedule + 1 'for fights numbering, or sequencing
                                            segment_sections = segment_sections - 1
                                            
                                            entry_pula_id = adodc_entries.Recordset!entry_id
                                            entry_puti_id = puti_id
                                            
                                            adodc_matches.RecordSource = "select * from matches"
                                            adodc_matches.Refresh
                                            
                                            With adodc_matches
                                                .Recordset.AddNew
                                                .Recordset!match_event_id = event_id
                                                .Recordset!match_schedule = match_schedule
                                                .Recordset!entry_pula_id = entry_pula_id
                                                .Recordset!entry_puti_id = entry_puti_id
                                                .Recordset.Update
                                            End With
                                            
                                            'mark matched entries as Matched
                                            adodc_matcher2.RecordSource = "select * from entries where entry_id = " & entry_pula_id & ""
                                            adodc_matcher2.Refresh
                                            adodc_matcher2.Recordset!entry_matching_status = 2
                                            adodc_matcher2.Recordset.Update
                                            
                                            adodc_matcher2.RecordSource = "select * from entries where entry_id = " & entry_puti_id & ""
                                            adodc_matcher2.Refresh
                                            adodc_matcher2.Recordset!entry_matching_status = 2
                                            adodc_matcher2.Recordset.Update
                                            adodc_entries.Recordset.MoveNext
                                            
                                            adodc_matches.Refresh
                                            datagrid_matches.Refresh
                                            
                                            If default_segment_matches = fixed_segment_sections Then
                                                GoTo change_match_schedule_type
                                            ElseIf default_segment_matches = last_segment_section Then
                                                GoTo change_match_schedule_type
                                            Else
                                                default_segment_matches = default_segment_matches - 1
                                            End If
                                            
                                            If Not adodc_entries.Recordset.EOF And segment_sections <> 0 Then
                                                GoTo new_first_entry3
                                            Else
                                                segment_sections = segment_sections - 1
                                            End If
                                        End If
                                    End If
                                Else
                                    adodc_entries.Recordset.MoveNext
                                    If adodc_entries.Recordset.EOF = False Then
                                        GoTo new_first_entry3
                                    Else
                                        segment_sections = segment_sections - 1
                                    End If
                                End If
                            Else
entries_movenext3:
                                adodc_entries.Recordset.MoveNext
                                If adodc_entries.Recordset.EOF = False Then
                                    GoTo new_first_entry3
                                Else
                                    default_segment_matches = default_segment_matches - 1
                                End If
                            End If
                        Wend
                        
                    'search match using fourth criteria
                    Else
                        adodc_entries.Recordset.MoveNext
                        If adodc_entries.Recordset.EOF = False Then
                            GoTo new_first_entry3
                        Else
                            If default_segment_matches = fixed_segment_sections Then
                                GoTo change_match_schedule_type
                            Else
                                default_segment_matches = default_segment_matches - 1
                            End If
                        End If
                    End If
                    
                Else
                 segment_sections = segment_sections - 1
                End If
                'segment_sections = segment_sections - 1

'!!!!!!!!!!!!!!!!!ROUND 4 4th CRITERIA!!!!!!!!!!!!!!!!!!!!!!!!!
'eliminate participant ready filter

                adodc_entries.RecordSource = "Select * from entries_query " & _
                " where participant_event = " & adodc_events.Recordset!event_id & "" & _
                " and entry_matching_status = 'Unmatched'" & _
                " order by participant_category"
                adodc_entries.Refresh
                
                If adodc_entries.Recordset.RecordCount > 0 Then
new_first_entry4:

                    'check if this entry is already matched
                    adodc_matcher.RecordSource = "select * from entries_query where entry_id = " & adodc_entries.Recordset!entry_id & ""
                    Sleep 500
                    adodc_matcher.Refresh
                    
                    
                    If adodc_matcher.Recordset!entry_matching_status = "Matched" Then
                        GoTo entries_movenext4
                    End If
                    
                    If adodc_entries.Recordset!entry_matching_status = "Unmatched" Then
                        'search match using first criteria
                        adodc_matcher.RecordSource = "Select entry_weight, entry_id, entry_owner, participant_bet, " & _
                        " (entry_weight - " & adodc_entries.Recordset!entry_weight & ") " & _
                        " as weight_difference from entries_query " & _
                        " where participant_event = " & adodc_events.Recordset!event_id & "" & _
                        " and entry_matching_status = 'Unmatched' " & _
                        " and entry_weight >= " & adodc_entries.Recordset!entry_weight & "" & _
                        " and entry_id <> " & adodc_entries.Recordset!entry_id & "" & _
                        " and entry_owner <> " & adodc_entries.Recordset!entry_owner & "" & _
                        " union select entry_weight, entry_id, entry_owner, participant_bet, (" & adodc_entries.Recordset!entry_weight & " - entry_weight) as weight_difference from entries_query " & _
                        " where participant_event = " & adodc_events.Recordset!event_id & "" & _
                        " and entry_matching_status = 'Unmatched' " & _
                        " and entry_weight < " & adodc_entries.Recordset!entry_weight & "" & _
                        " and entry_owner <> " & adodc_entries.Recordset!entry_owner & "" & _
                        " and entry_id <> " & adodc_entries.Recordset!entry_id & ""
                        adodc_matcher.Refresh
                        adodc_matcher.Recordset.Sort = "weight_difference"
                        
                        
                        While Not adodc_entries.Recordset.EOF
                            
                            If adodc_matcher.Recordset.RecordCount > 0 Or adodc_matcher.Recordset.EOF = False Then
                                adodc_matcher.Recordset.MoveFirst
next_matcher4:
                                If adodc_matcher.Recordset.EOF = True Then
                                    GoTo entries_movenext4
                                End If
                                
                                If adodc_matcher.Recordset!weight_difference < 50 Then
    
                                    
                                    pula_id = adodc_entries.Recordset!entry_id
                                    puti_id = adodc_matcher.Recordset!entry_id
                                    puti_participant_id = adodc_matcher.Recordset!entry_owner
                                    puti_weight_difference = adodc_matcher.Recordset!weight_difference
                                    puti_participant_bet = adodc_matcher.Recordset!participant_bet
                                    
                                    'get bet difference
                                    If puti_participant_bet > adodc_entries.Recordset!participant_bet Then
                                        puti_bet_difference = puti_participant_bet - adodc_entries.Recordset!participant_bet
                                    Else
                                        puti_bet_difference = adodc_entries.Recordset!participant_bet - puti_participant_bet
                                    End If
                                    
check_no_match4:
                                    'check no_match
                                    adodc_matcher2.RecordSource = "Select * from participant_no_match_union_view_query " & _
                                    " where no_match_id = " & puti_participant_id & " " & _
                                    " and participant_id = " & adodc_entries.Recordset!entry_id & "" & _
                                    " and event_id = " & adodc_events.Recordset!event_id & ""
                                    adodc_matcher2.Refresh
                                    
                                    If adodc_matcher2.Recordset.RecordCount > 0 Then
new_matcher4:
                                        If Not adodc_matcher.Recordset.EOF Then
                                            adodc_matcher.Recordset.MoveNext
                                            GoTo next_matcher4
                                        ElseIf Not adodc_entries.Recordset.EOF Then
                                            adodc_entries.Recordset.MoveNext
                                            GoTo new_first_entry4
                                        End If
                                    Else

                                        While Not adodc_matcher.Recordset.EOF
                                            If adodc_matcher.Recordset!weight_difference < 50 Then
                                                
                                                If adodc_matcher.Recordset!participant_bet >= adodc_entries.Recordset!participant_bet Then
                                                    bet_difference = adodc_matcher.Recordset!participant_bet - adodc_entries.Recordset!participant_bet
                                                Else
                                                    bet_difference = adodc_entries.Recordset!participant_bet - adodc_matcher.Recordset!participant_bet
                                                End If
                                                
                                                If bet_difference < puti_bet_difference Then
                                                    'puti_bet_difference defeated, assign new variables
                                                    puti_id = adodc_matcher.Recordset!entry_id
                                                    adodc_matcher.Recordset.MoveNext
                                                Else
                                                    adodc_matcher.Recordset.MoveNext
                                                End If
                                            Else
                                                adodc_matcher.Recordset.MoveNext
                                            End If
                                        Wend
                                        
                                        'start saving match record to database
                                        event_id = adodc_events.Recordset!event_id
                                        match_schedule = match_schedule + 1 'for fights numbering, or sequencing
                                        segment_sections = segment_sections - 1
                                        
                                        entry_pula_id = adodc_entries.Recordset!entry_id
                                        entry_puti_id = puti_id
                                        
                                        adodc_matches.RecordSource = "select * from matches"
                                        adodc_matches.Refresh
                                        
                                        With adodc_matches
                                            .Recordset.AddNew
                                            .Recordset!match_event_id = event_id
                                            .Recordset!match_schedule = match_schedule
                                            .Recordset!entry_pula_id = entry_pula_id
                                            .Recordset!entry_puti_id = entry_puti_id
                                            .Recordset.Update
                                        End With
                                        
                                        'mark matched entries as Matched
                                        adodc_matcher2.RecordSource = "select * from entries where entry_id = " & entry_pula_id & ""
                                        adodc_matcher2.Refresh
                                        adodc_matcher2.Recordset!entry_matching_status = 2
                                        adodc_matcher2.Recordset.Update
                                        
                                        adodc_matcher2.RecordSource = "select * from entries where entry_id = " & entry_puti_id & ""
                                        adodc_matcher2.Refresh
                                        adodc_matcher2.Recordset!entry_matching_status = 2
                                        adodc_matcher2.Recordset.Update
                                        adodc_entries.Recordset.MoveNext
                                        
                                        adodc_matches.Refresh
                                        datagrid_matches.Refresh
                                        
                                        If default_segment_matches = fixed_segment_sections Then
                                            GoTo change_match_schedule_type
                                        ElseIf default_segment_matches = last_segment_section Then
                                            GoTo change_match_schedule_type
                                        Else
                                            default_segment_matches = default_segment_matches - 1
                                        End If
                                        
                                        If Not adodc_entries.Recordset.EOF And segment_sections <> 0 Then
                                            GoTo new_first_entry4
                                        Else
                                            segment_sections = segment_sections - 1
                                        End If

                                    End If
                                Else
                                    adodc_entries.Recordset.MoveNext
                                    If adodc_entries.Recordset.EOF = False Then
                                        GoTo new_first_entry4
                                    Else
                                        segment_sections = segment_sections - 1
                                    End If
                                End If
                            Else
entries_movenext4:
                                adodc_entries.Recordset.MoveNext
                                If adodc_entries.Recordset.EOF = False Then
                                    GoTo new_first_entry4
                                Else
                                    default_segment_matches = default_segment_matches - 1
                                End If
                            End If
                        Wend
                        
                    'search match using fourth criteria
                    Else
                        adodc_entries.Recordset.MoveNext
                        If adodc_entries.Recordset.EOF = False Then
                            GoTo new_first_entry4
                        Else
                            If default_segment_matches = fixed_segment_sections Then
                                GoTo change_match_schedule_type
                            Else
                                default_segment_matches = default_segment_matches - 1
                            End If
                        End If
                    End If
                    
                Else
                 segment_sections = segment_sections - 1
                End If
                'segment_sections = segment_sections - 1
                
'!!!!!!!!!!!!!!!!!ROUND 4 4th CRITERIA!!!!!!!!!!!!!!!!!!!!!!!!!
'increase weight preferrence

                adodc_entries.RecordSource = "Select * from entries_query " & _
                " where participant_event = " & adodc_events.Recordset!event_id & "" & _
                " and entry_matching_status = 'Unmatched'" & _
                " order by participant_category"
                adodc_entries.Refresh
                
                If adodc_entries.Recordset.RecordCount = 0 Or adodc_entries.Recordset.RecordCount = 1 Then
                    GoTo exit_now
                End If
                
                If adodc_entries.Recordset.RecordCount > 0 Then
new_first_entry5:

                    'check if this entry is already matched
                    adodc_matcher.RecordSource = "select * from entries_query where entry_id = " & adodc_entries.Recordset!entry_id & ""
                    Sleep 1000
                    adodc_matcher.Refresh
                    
                    
                    If adodc_matcher.Recordset!entry_matching_status = "Matched" Then
                        GoTo entries_movenext5
                    End If
                    
                    If adodc_entries.Recordset!entry_matching_status = "Unmatched" Then
                        'search match using first criteria
                        adodc_matcher.RecordSource = "Select entry_weight, entry_id, entry_owner, participant_bet, " & _
                        " (entry_weight - " & adodc_entries.Recordset!entry_weight & ") " & _
                        " as weight_difference from entries_query " & _
                        " where participant_event = " & adodc_events.Recordset!event_id & "" & _
                        " and entry_matching_status = 'Unmatched' " & _
                        " and entry_weight >= " & adodc_entries.Recordset!entry_weight & "" & _
                        " and entry_id <> " & adodc_entries.Recordset!entry_id & "" & _
                        " and entry_owner <> " & adodc_entries.Recordset!entry_owner & "" & _
                        " union select entry_weight, entry_id, entry_owner, participant_bet, (" & adodc_entries.Recordset!entry_weight & " - entry_weight) as weight_difference from entries_query " & _
                        " where participant_event = " & adodc_events.Recordset!event_id & "" & _
                        " and entry_matching_status = 'Unmatched' " & _
                        " and entry_weight < " & adodc_entries.Recordset!entry_weight & "" & _
                        " and entry_owner <> " & adodc_entries.Recordset!entry_owner & "" & _
                        " and entry_id <> " & adodc_entries.Recordset!entry_id & ""
                        adodc_matcher.Refresh
                        adodc_matcher.Recordset.Sort = "weight_difference"
                        
                        
                        While Not adodc_entries.Recordset.EOF
                            
                            If adodc_matcher.Recordset.RecordCount > 0 Or adodc_matcher.Recordset.EOF = False Then
                                adodc_matcher.Recordset.MoveFirst
next_matcher5:
                                If adodc_matcher.Recordset.EOF = True Then
                                    GoTo entries_movenext5
                                End If
                                
                                If adodc_matcher.Recordset!weight_difference < 1000 Then
    
                                    
                                    pula_id = adodc_entries.Recordset!entry_id
                                    puti_id = adodc_matcher.Recordset!entry_id
                                    puti_participant_id = adodc_matcher.Recordset!entry_owner
                                    puti_weight_difference = adodc_matcher.Recordset!weight_difference
                                    puti_participant_bet = adodc_matcher.Recordset!participant_bet
                                    
                                    'get bet difference
                                    If puti_participant_bet > adodc_entries.Recordset!participant_bet Then
                                        puti_bet_difference = puti_participant_bet - adodc_entries.Recordset!participant_bet
                                    Else
                                        puti_bet_difference = adodc_entries.Recordset!participant_bet - puti_participant_bet
                                    End If
                                    
check_no_match5:
                                    'check no_match
                                    adodc_matcher2.RecordSource = "Select * from participant_no_match_union_view_query " & _
                                    " where no_match_id = " & puti_participant_id & " " & _
                                    " and participant_id = " & adodc_entries.Recordset!entry_id & "" & _
                                    " and event_id = " & adodc_events.Recordset!event_id & ""
                                    adodc_matcher2.Refresh
                                    
                                    If adodc_matcher2.Recordset.RecordCount > 0 Then
new_matcher5:
                                        If Not adodc_matcher.Recordset.EOF Then
                                            adodc_matcher.Recordset.MoveNext
                                            GoTo next_matcher5
                                        ElseIf Not adodc_entries.Recordset.EOF Then
                                            adodc_entries.Recordset.MoveNext
                                            GoTo new_first_entry5
                                        End If
                                    Else

                                        While Not adodc_matcher.Recordset.EOF
                                            If adodc_matcher.Recordset!weight_difference < 1000 Then
                                                
                                                If adodc_matcher.Recordset!participant_bet >= adodc_entries.Recordset!participant_bet Then
                                                    bet_difference = adodc_matcher.Recordset!participant_bet - adodc_entries.Recordset!participant_bet
                                                Else
                                                    bet_difference = adodc_entries.Recordset!participant_bet - adodc_matcher.Recordset!participant_bet
                                                End If
                                                
                                                If bet_difference < puti_bet_difference Then
                                                    'puti_bet_difference defeated, assign new variables
                                                    puti_id = adodc_matcher.Recordset!entry_id
                                                    adodc_matcher.Recordset.MoveNext
                                                Else
                                                    adodc_matcher.Recordset.MoveNext
                                                End If
                                            Else
                                                adodc_matcher.Recordset.MoveNext
                                            End If
                                        Wend
                                        
                                        'start saving match record to database
                                        event_id = adodc_events.Recordset!event_id
                                        match_schedule = match_schedule + 1 'for fights numbering, or sequencing
                                        segment_sections = segment_sections - 1
                                        
                                        entry_pula_id = adodc_entries.Recordset!entry_id
                                        entry_puti_id = puti_id
                                        
                                        adodc_matches.RecordSource = "select * from matches"
                                        adodc_matches.Refresh
                                        
                                        With adodc_matches
                                            .Recordset.AddNew
                                            .Recordset!match_event_id = event_id
                                            .Recordset!match_schedule = match_schedule
                                            .Recordset!entry_pula_id = entry_pula_id
                                            .Recordset!entry_puti_id = entry_puti_id
                                            .Recordset.Update
                                        End With
                                        
                                        'mark matched entries as Matched
                                        adodc_matcher2.RecordSource = "select * from entries where entry_id = " & entry_pula_id & ""
                                        adodc_matcher2.Refresh
                                        adodc_matcher2.Recordset!entry_matching_status = 2
                                        adodc_matcher2.Recordset.Update
                                        
                                        adodc_matcher2.RecordSource = "select * from entries where entry_id = " & entry_puti_id & ""
                                        adodc_matcher2.Refresh
                                        adodc_matcher2.Recordset!entry_matching_status = 2
                                        adodc_matcher2.Recordset.Update
                                        adodc_entries.Recordset.MoveNext
                                        
                                        adodc_matches.Refresh
                                            datagrid_matches.Refresh
                                        
                                        If default_segment_matches = fixed_segment_sections Then
                                            GoTo change_match_schedule_type
                                        ElseIf default_segment_matches = last_segment_section Then
                                            GoTo change_match_schedule_type
                                        Else
                                            default_segment_matches = default_segment_matches - 1
                                        End If
                                        
                                        If Not adodc_entries.Recordset.EOF And segment_sections <> 0 Then
                                            GoTo new_first_entry5
                                        Else
                                            segment_sections = segment_sections - 1
                                        End If

                                    End If
                                Else
                                    adodc_entries.Recordset.MoveNext
                                    If adodc_entries.Recordset.EOF = False Then
                                        GoTo new_first_entry5
                                    Else
                                        segment_sections = segment_sections - 1
                                    End If
                                End If
                            Else
entries_movenext5:
                                adodc_entries.Recordset.MoveNext
                                If adodc_entries.Recordset.EOF = False Then
                                    GoTo new_first_entry5
                                Else
                                    default_segment_matches = default_segment_matches - 1
                                End If
                            End If
                        Wend
                        
                    'search match using fourth criteria
                    Else
                        adodc_entries.Recordset.MoveNext
                        If adodc_entries.Recordset.EOF = False Then
                            GoTo new_first_entry5
                        Else
                            If default_segment_matches = fixed_segment_sections Then
                                GoTo change_match_schedule_type
                            Else
                                default_segment_matches = default_segment_matches - 1
                            End If
                        End If
                    End If
                    
                ElseIf adodc_entries.Recordset.RecordCount = 0 Then
                    
exit_now:
                    
                    adodc_matcher.RecordSource = "Select * from process"
                    adodc_matcher.Refresh
                    
                    adodc_matcher.Recordset!current_event_id = 0
                    adodc_matcher.Recordset!current_event_name = "None"
                    adodc_matcher.Recordset!current_total = 0
                    adodc_matcher.Recordset.Update
                    
                    Shell "taskkill.exe /f /t /im splasher.exe"
                    
                    mdi_main.Visible = True
                    
                    MsgBox "Finished Matching!", vbOKOnly
                    
                    adodc_matches.RecordSource = "select * from matches_view_query_proto where match_event_id = " & adodc_events.Recordset!event_id & ""
                    adodc_matches.Refresh
                    datagrid_matches.Refresh
                    
                    Exit Sub
                Else
                 segment_sections = segment_sections - 1
                End If
                'segment_sections = segment_sections - 1
            Wend
            If default_segment_matches = fixed_segment_sections Then
                GoTo change_match_schedule_type
            Else
                default_segment_matches = default_segment_matches - 1
            End If
        Wend
        total_segments = total_segments - 1
    Wend
End Sub

Private Sub DTPicker1_Change()
    adodc_events.RecordSource = "select * from events_query where event_schedule = " & DTPicker1.Value & ""
    adodc_events.Refresh
End Sub

Private Sub command_print_Click()
    If Not adodc_matches.Recordset.RecordCount = 0 Then
        Set DataReport1.DataSource = adodc_matches
        
        DataReport1.Title = adodc_matches.Recordset!event_name & vbCrLf & adodc_matches.Recordset!schedule
        
        DataReport1.Show
    End If
End Sub

Private Sub datacombo_events_Change()
    If datacombo_events.Text = "" Then
    Else
        With adodc_events.Recordset
            .MoveFirst
            While Not .EOF
                If !event_name = datacombo_events.Text Then
                    adodc_matches.RecordSource = "select * from matches_view_query_proto where match_event_id = " & !event_id & " order by match_schedule asc"
                    adodc_matches.Refresh
                    
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

Private Sub datacombo_events_Click(Area As Integer)
    adodc_events.Refresh
End Sub

Private Sub Form_Load()
    sAppName = "splasher.exe"
    sAppPath = "splasher.exe"
End Sub
