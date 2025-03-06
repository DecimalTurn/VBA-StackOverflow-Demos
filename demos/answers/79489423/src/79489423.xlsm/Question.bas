Attribute VB_Name = "Question"
Option Explicit

Public Sub ExecuteM(ByVal mCode As String)
  Dim wb As Workbook: Set wb = Workbooks.Add()
  Dim query As WorkbookQuery: Set query = wb.Queries.Add("PQ", mCode)
  Dim ws As Worksheet: Set ws = wb.Sheets(1)
  Dim lo As ListObject: Set lo = ws.ListObjects.Add(xlSrcQuery, "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=PQ;Extended Properties=""""", Destination:=ws.Range("A1"))
  Dim qt As QueryTable: Set qt = lo.QueryTable
  qt.CommandType = xlCmdSql
  qt.CommandText = Array("SELECT * FROM [PQ]")
  
  'Refresh async...
  Call qt.Refresh(True)
  
  'The data will never populate...
  While qt.Refreshing
    DoEvents
  Wend
End Sub
