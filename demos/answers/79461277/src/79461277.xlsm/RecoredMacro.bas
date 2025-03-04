Attribute VB_Name = "RecoredMacro"
Option Explicit

Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    ActiveWorkbook.Queries.Add Name:="sample_data1", _
        Formula:= _
        "let" & Chr(13) & "" & Chr(10) & _
        "    Source = Csv.Document(Web.Contents(""https://raw.githubusercontent.com/DecimalTurn/VBA-StackOverflow-Demos/refs/heads/main/data/csv/sample_data1.csv""),[Delimiter="","", Columns=5, Encoding=65001, QuoteStyle=QuoteStyle.None])," & Chr(13) & "" & Chr(10) & _
        "    #""Promoted Headers"" = Table.PromoteHeaders(Source, [PromoteAllScalars=true])" & "," & Chr(13) & "" & Chr(10) & _
        "    #""Changed Type"" = Table.TransformColumnTypes(#""Promoted Headers"",{{""ID"", Int64.Type}, {""Name"", type text}, {""Age"", Int64.Type}, {""Email"", type text}, {""Country"", type text}})" & Chr(13) & "" & Chr(10) & _
        "in" & Chr(13) & "" & Chr(10) & _
        "    #""Changed Type"""
    ActiveWorkbook.Worksheets.Add
    With ActiveSheet.ListObjects.Add( _
            SourceType:=0, _
            Source:="OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=sample_data1;Extended Properties=""""", _
            Destination:=Range("$A$1") _
        ).QueryTable
        
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [sample_data1]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = "sample_data1"
        .Refresh BackgroundQuery:=False
        
    End With
End Sub
