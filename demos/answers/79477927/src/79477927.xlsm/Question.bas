Attribute VB_Name = "Question"
Option Explicit

Sub question()

Dim FolderName As String
FolderName = "Subfolder"

'folderName is identified earlier
Dim Files As FileDialog
Dim vrtSelectedItem As Variant
Dim List As String, WbName As String, cnt As Long, atmpt As Long

WbName = Left(ThisWorkbook.Path, InStrRev(ThisWorkbook.Path, "/")) & _
         "Registry/Joining Instructions/" & FolderName & "/"

Do
   Set Files = Application.FileDialog(msoFileDialogFilePicker)
   cnt = 0
   atmpt = atmpt + 1
        
   With Files
       .InitialFileName = WbName

       Application.SendKeys "+{TAB}", True
       Application.SendKeys "+{TAB}", True
       Application.SendKeys "^a", True
       Application.SendKeys "~", True
                    
       If .Show = -1 Then
            'Code continues...
       End If
       
    End With
    
    If atmpt > 3 Then
        Exit Do
    End If
    
Loop

End Sub

Sub DisplayDialog()

Dim FolderName As String
FolderName = "Subfolder"

'folderName is identified earlier
Dim Files As FileDialog
Dim vrtSelectedItem As Variant
Dim List As String, WbName As String, atmpt As Long

'Note that to make it work for a standard path, "/" was replaced by "\"
WbName = ThisWorkbook.Path & "\" & FolderName & "\"

Do

   Set Files = Application.FileDialog(msoFileDialogFilePicker)
   Dim cnt As Long
   cnt = 0
   atmpt = atmpt + 1
        
   With Files
       
       .InitialFileName = WbName

       Application.SendKeys "+{TAB}", True
       Application.SendKeys "+{TAB}", True
       Application.SendKeys "^a", True
       Application.SendKeys "~", True
                    
       If .Show = -1 Then
            Exit Do
       End If
       
    End With
    
    If atmpt > 3 Then
        Exit Do
    End If
    
Loop
       
Dim File As Variant
For Each File In Files.SelectedItems
    Debug.Print File
Next

End Sub

