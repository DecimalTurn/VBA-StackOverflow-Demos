Attribute VB_Name = "Answer"
Option Explicit

Sub DemoGetLocalPath()

    Dim SubFolderName As String
    SubFolderName = "Subfolder"
    
    Dim LocalFolderPath As String
    LocalFolderPath = GetLocalPath(ThisWorkbook.Path)
          
    Dim MyFiles As Collection
    Set MyFiles = GetFiles(LocalFolderPath & "\" & SubFolderName)
          
    Dim FilePath As Variant
    For Each FilePath In MyFiles
        Debug.Print FilePath
    Next

End Sub


