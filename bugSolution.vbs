Function MyFunction(param1, param2)
  On Error Resume Next
  If IsNumeric(param1) And IsNumeric(param2) Then
    result = param1 + param2
    If Err.Number <> 0 Then
      MsgBox "Error during calculation: " & Err.Description
    End If
  Else
    MsgBox "Error: Parameters must be numeric."
  End If
  On Error GoTo 0
  MyFunction = result
End Function

' Example usage:
Dim result
result = MyFunction(5, "hello") 'Incorrect Parameter Type, triggers error handling.
result = MyFunction(10, 20) 'Correct parameters, returns 30
result = MyFunction("a", "b") ' Incorrect parameter type, triggers error handling.
MsgBox result 