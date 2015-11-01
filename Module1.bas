Attribute VB_Name = "Module1"
Public conn As New ADODB.Connection
Public rsEntry As New ADODB.Recordset

Public Sub Connect()
    Set conn = New ADODB.Connection
     conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\data\fight_db.mdb"
     conn.Open
End Sub

Public Sub Entry()
 Set rsEntry = Nothing
    Set rsEntry = New ADODB.Recordset

    With rsEntry
         .CursorType = adOpenDynamic
         .LockType = adLockOptimistic
         .ActiveConnection = conn
         .Source = "Select * from events"
         .CursorLocation = adUseClient
         .Open
    End With
End Sub
