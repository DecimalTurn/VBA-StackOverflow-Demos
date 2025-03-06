Attribute VB_Name = "MultipleQueries"
Option Explicit

Private PendingQueries As Object

Sub testM()

    ' Initialize the dictionary
    Set PendingQueries = CreateObject("Scripting.Dictionary")
    
    ' Execute multiple queries in parallel
    Call ExecuteM("Query1", "#table({""a"",""b""},{{1,2},{3,4}})")
    Call ExecuteM("Query2", "#table({""c"",""d""},{{1,2},{3,4}})")
    
    ' Start monitoring queries if not already running
    If PendingQueries.Count >= 1 Then
        Application.OnTime Now + TimeValue("00:00:01"), "CheckQueriesRefreshStatus"
    End If
    
End Sub

Public Sub ExecuteM(ByVal QueryName As String, ByVal mCode As String)
    Dim wb As Workbook: Set wb = Workbooks.Add()
    Dim query As WorkbookQuery: Set query = wb.Queries.Add("PQ", mCode)
    Dim ws As Worksheet: Set ws = wb.Sheets(1)
    Dim lo As ListObject
    Set lo = ws.ListObjects.Add(xlSrcQuery, _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=PQ;Extended Properties=""""", _
        Destination:=ws.Range("A1"))

    Dim qt As QueryTable
    Set qt = lo.QueryTable
    qt.CommandType = xlCmdSql
    qt.CommandText = Array("SELECT * FROM [PQ]")
    
    ' Store query in the dictionary
    PendingQueries.Add QueryName, qt

    ' Refresh asynchronously
    Call qt.Refresh(True)

End Sub

Public Sub CheckQueriesRefreshStatus()
    Dim i As Integer
    Dim qt As QueryTable
    Dim keysToRemove As Collection
    Set keysToRemove = New Collection

    ' Check all queries in the dictionary
    Dim Key As Variant
    For Each Key In PendingQueries.Keys
        Set qt = PendingQueries(Key)
        If Not qt.Refreshing Then
            ' Mark this query for removal
            keysToRemove.Add Key
        End If
    Next Key

    ' Remove completed queries
    For i = 1 To keysToRemove.Count
        MsgBox keysToRemove(i) & " refresh complete!", vbInformation
        PendingQueries.Remove keysToRemove(i)
    Next i

    ' If there are still queries running, check again in 1 second
    If PendingQueries.Count > 0 Then
        Application.OnTime Now + TimeValue("00:00:01"), "CheckQueriesRefreshStatus"
    Else
        MsgBox "All queries have finished refreshing!", vbInformation
    End If
End Sub




