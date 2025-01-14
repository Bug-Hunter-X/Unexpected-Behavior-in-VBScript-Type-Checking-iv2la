Function f(a,b)
  If IsNumeric(a) Then
    If IsNumeric(b) Then
      f = a+b
    Else
      f = "Error: b is not numeric"
    End If
  Else
    f = "Error: a is not numeric"
  End If
End Function
MsgBox f(10,"hello")