Function CheckEmail(email)
  Set regEx = New RegExp
  regEx.Pattern = "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
  CheckEmail = regEx.Test(email)
  Set regEx = Nothing
End Function

MsgBox CheckEmail("test@example.com") ' Returns True
MsgBox CheckEmail("test@@example.com") ' Returns False (Correct)
MsgBox CheckEmail("test@example.")  ' Returns False (Correct) 