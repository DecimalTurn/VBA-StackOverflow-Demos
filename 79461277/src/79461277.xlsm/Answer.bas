Attribute VB_Name = "Answer"
Option Explicit

Private Sub DataPull()
    
    Dim Wb As Workbook
    Set Wb = ThisWorkbook
    
    Dim CsvUrl As String
    CsvUrl = "https://raw.githubusercontent.com/DecimalTurn/VBA-StackOverflow-Demos/refs/heads/main/data/csv/sample_data1.csv"
    'CsvUrl = ThisWorkbook.Worksheets("Settings").Range("Y18").Value2

    Wb.Queries.Item("sample_data1").Delete
    Wb.Queries.Add Name:="sample_data1", _
        Formula:= _
        "let" & Chr(13) & "" & Chr(10) & _
        "    Source = Csv.Document(Web.Contents(""" & CsvUrl & """),[Delimiter="","", Columns=5, Encoding=65001, QuoteStyle=QuoteStyle.None])," & Chr(13) & "" & Chr(10) & _
        "    #""Promoted Headers"" = Table.PromoteHeaders(Source, [PromoteAllScalars=true])" & Chr(13) & "" & Chr(10) & _
        "in" & Chr(13) & "" & Chr(10) & _
        "    #""Promoted Headers"""
    
    'Corresponding M code:
    'let
    '    Source = Csv.Document(Web.Contents("https://raw.githubusercontent.com/DecimalTurn/VBA-StackOverflow-Demos/refs/heads/main/data/csv/sample_data1.csv"),[Delimiter=",", Columns=5, Encoding=65001, QuoteStyle=QuoteStyle.None]),
    '    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true])
    'in
    '    #"Promoted Headers"

    Wb.Sheets("Data_Pull").Select
    Wb.Sheets("Data_pull").UsedRange.ClearContents
    
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

