Attribute VB_Name = "WrongNumberOfArguments"
Option Explicit


'This modules shows how to get the Error Wrong Number of arguments or invalid property assignment

'Objective: Show when we get "Wrong Number of arguments or invalid property assignment"

'Conclusion: It seems that we only get this error when the number of arguments supplied to the sub/function is higher that the number of argument it takes.
'Also note that a function with no arg will still try to evaluate if you try to pass it arguments

Sub DemoSub()

    SubWithNoArg
    'SubWithNoArg ("test") 'Wrong Number of arguments or invalid property assignment
    'SubWithNoArg ("test", "test") 'Doesn't compile
    
    'SubWith1Arg 'Error: Argument not optional
    SubWith1Arg ("test")
    
        
    'SubWith2Args 'Error: Argument not optional
    'SubWith2Args ("test") 'Error: Argument not optional

End Sub


Sub SubWithNoArg()
    Debug.Print "0"
End Sub

Sub SubWith1Arg(a)
    Debug.Print a
End Sub

Sub SubWith2Args(a, b)
    Debug.Print a & b
End Sub



Sub DemoFunc()
    Debug.Print Func0
    'Debug.Print Func0("test") 'Error: Type mismatch
    'Debug.Print Func0("test", "test") 'Error: Type mismatch
    
    'Debug.Print Func1 'Error: Argument not optional
    Debug.Print Func1("test")
    'Debug.Print Func1("test", "test") 'Error: Wrong Number of arguments or invalid property assignment
    
    'Debug.Print Func2 'Error: Argument not optional
    'Debug.Print Func2("test") 'Error: Argument not optional
    Debug.Print Func2("test", "test")
    
    
End Sub

Function Func0()
    Func0 = "0"
End Function

Function Func1(a)
    Func1 = a
End Function

Function Func2(a, b)
    Func2 = a & b
End Function
