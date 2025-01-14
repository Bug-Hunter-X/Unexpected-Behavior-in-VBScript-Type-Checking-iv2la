Function f(a,b)
  If IsNumeric(a) And IsNumeric(b) Then
    f = a + b
  Else
    If Not IsNumeric(a) Then
      f = "Error: a is not numeric"
    ElseIf Not IsNumeric(b) Then
      f = "Error: b is not numeric"
    End If
  End If
End Function
MsgBox f(10,"hello")
MsgBox f(10,20)
MsgBox f("hello",20)