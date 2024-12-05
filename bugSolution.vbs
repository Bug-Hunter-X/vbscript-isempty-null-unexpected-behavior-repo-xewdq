Function f(a)
  If IsNothing(a) Then
    f = 0
  ElseIf IsEmpty(a) Then
    f = 0
  Else
    f = a
  End If
End Function

x = f(Null)
MsgBox x