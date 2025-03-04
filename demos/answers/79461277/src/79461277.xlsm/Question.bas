Attribute VB_Name = "Question"
Option Explicit

Private Sub Data_Pull()
    Dim Myurl As String
    Myurl = Worksheets("Settings").Range("Y18").Value
    Sheets("Data_Pull").Visible = True

    Sheets("Data_Pull").Select
    Sheets("Data_pull").UsedRange.ClearContents
    With ActiveSheet.QueryTables.Add(Connection:= _
        "URL;" + Myurl, Destination:=Range("$A$1"))
        .Name = "287%2C321%2C362%2C632%2C379%2C426%2C720&osm_ids=&oxm_ids=445%2C442&ofm_ids=&datasource_viz=nvd3Table"
        .FieldNames = True
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
        .WebSelectionType = xlSpecifiedTables
        .WebFormatting = xlWebFormattingNone
        .WebTables = "2"
        .WebPreFormattedTextToColumns = True
        .WebConsecutiveDelimitersAsOne = False
        .WebSingleBlockTextImport = False
        .WebDisableDateRecognition = False
        .WebDisableRedirections = False
        .Refresh BackgroundQuery:=False
    End With
    'Sheets("Data_Pull").Visible = True
    
End Sub

