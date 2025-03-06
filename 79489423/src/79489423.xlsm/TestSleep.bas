Attribute VB_Name = "TestSleep"
Option Explicit

'Declare the Sleep() method from the Windows API
#If VBA7 Then ' Excel 2010 or later
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else ' Excel 2007 or earlier
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

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
    Sleep 5000
    'Application.Wait (Now + TimeValue("00:00:05"))
    Debug.Print "Waiting..."
  Wend
End Sub

Sub testM()
  Call ExecuteM("#table({""a"",""b""},{{1,2},{3,4}})")
End Sub
