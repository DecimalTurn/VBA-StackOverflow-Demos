Attribute VB_Name = "BreakingTypeName"
Option Explicit

Sub BrokenTypeNameDemo()
    
    Dim obj As CustomClass
    Set obj = New CustomClass
    
    'This line will give a compile error: Wrong Number of arguments or invalid property assignment
    'That's because it tries to use the definition of TypeName as defined in the private sub below.
    Debug.Print TypeName(obj)

End Sub

Private Sub TypeName()
    
End Sub


