Attribute VB_Name = "SimpleDemo"
Option Explicit

Sub SimpleDemo()

    Dim obj As CustomClass
    Set obj = New CustomClass
    
    Debug.Print TypeName(obj)

End Sub


Sub PassingVariant()

    Dim obj As CustomClass
    Set obj = New CustomClass
    
    ReceiveVariant obj

End Sub

Sub ReceiveVariant(Value As Variant)
    
    Debug.Print TypeName(Value)
    
End Sub
