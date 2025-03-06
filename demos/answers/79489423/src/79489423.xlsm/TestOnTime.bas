Attribute VB_Name = "TestOnTime"
Option Explicit

Private qt As QueryTable

Sub testM()
  Call ExecuteM("#table({""a"",""b""},{{1,2},{3,4}})")
End Sub

Public Sub ExecuteM(ByVal mCode As String)
    Dim wb As Workbook: Set wb = Workbooks.Add()
    Dim query As WorkbookQuery: Set query = wb.Queries.Add("PQ", mCode)
    Dim ws As Worksheet: Set ws = wb.Sheets(1)
    Dim lo As ListObject
    Set lo = ws.ListObjects.Add(xlSrcQuery, _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=PQ;Extended Properties=""""", _
        Destination:=ws.Range("A1"))
    
    Set qt = lo.QueryTable
    qt.CommandType = xlCmdSql
    qt.CommandText = Array("SELECT * FROM [PQ]")
    
    ' Refresh async
    Call qt.Refresh(True)

    ' Use OnTime to check status later
    Application.OnTime Now + TimeValue("00:00:01"), "CheckQueryRefreshStatus"
End Sub

Public Sub CheckQueryRefreshStatus()
    
    ' If still refreshing, check again in 1 second
    If qt.Refreshing Then
        Application.OnTime Now + TimeValue("00:00:01"), "CheckQueryRefreshStatus"
    Else
        MsgBox "Query refresh complete!", vbInformation
        
        'The rest of the code goes here
        
    End If
    
End Sub


