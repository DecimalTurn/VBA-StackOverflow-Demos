Attribute VB_Name = "UserDefinedTypeDemo"
Option Explicit


'Shows that TypeName doesn't support User-Defined Types

Public Type DemoType
    Message  As String
End Type

Sub Demo()

    Dim obj As DemoType
    
    'The following line will give a Compile Error:
    'Only user-defined types defined in public object modules can
    'be coerced to or from a variant or passed to late-bound functions
    Debug.Print TypeName(obj)

End Sub


