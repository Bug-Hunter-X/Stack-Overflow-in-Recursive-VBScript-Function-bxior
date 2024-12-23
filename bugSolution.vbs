Function f(x)
  If x = 0 Then
    f = 1
  ElseIf x < 0 Then
    f = 0 ' Handle negative inputs appropriately
  Else
    f = x * f(x - 1)
  End If
End Function
MsgBox f(5) 'Now it should handle 5 correctly

'Further Improvement with Iteration
Function f_iterative(x)
  Dim result
  result = 1
  If x < 0 Then
      result = 0 'Handle negative inputs
  Else
      Dim i
      For i = 1 To x
          result = result * i
      Next
  End If
  f_iterative = result
End Function
MsgBox f_iterative(5) 'Using iteration for efficiency and no stack overflow risk