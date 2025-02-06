Function CheckEmail(email)
  If InStr(1, email, "@", vbTextCompare) = 0 Then
    CheckEmail = False
  ElseIf InStrRev(email, ".", vbTextCompare) < InStr(1, email, "@", vbTextCompare) Then
    CheckEmail = False
  Else
    CheckEmail = True
  End If
End Function

MsgBox CheckEmail("test@example.com")  ' Returns True
MsgBox CheckEmail("test@@example.com") ' Returns True (Incorrect)
MsgBox CheckEmail("test@example.")   ' Returns True (Incorrect)