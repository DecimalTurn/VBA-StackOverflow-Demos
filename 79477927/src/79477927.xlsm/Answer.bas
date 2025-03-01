Attribute VB_Name = "Answer"
Option Explicit

Sub DemoGetLocalPath()


    Dim FolderName As String
    FolderName = "Subfolder"
    
    Dim FullFileName As String
    FullFileName = ThisWorkbook.Path & "\" & FolderName & "\"
          
    Dim MyFiles As Collection
    Set MyFiles = GetFiles(FullFileName)
          
    Dim File As Variant
    For Each File In MyFiles
        Debug.Print File
    Next

End Sub

