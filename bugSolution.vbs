Function MyFunction(param1)
  If param1 = vbNullString Then
    ' Handle empty parameter explicitly using vbNullString
    param1 = ""
  End If
  ' ... rest of the function
End Function