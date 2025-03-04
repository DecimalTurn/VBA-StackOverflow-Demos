Attribute VB_Name = "BreakingTypeNameProperty"
Option Explicit

Sub BrokenTypeNameDemo()
    
    Dim obj As CustomClass
    Set obj = New CustomClass
    
    'This line will give a a runtime error: Type mismatch
    'That's because it tries to use the definition of TypeName as defined in the private property below.
    Debug.Print TypeName(obj)

End Sub

Private Property Get TypeName()
    
End Property

Private Property Let TypeName(a)
    
End Property



