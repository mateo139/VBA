Public Sub unprotect_all()
    Dim ZakladkaNr As Integer
    For ZakladkaNr = 1 To 16
        Sheets(ZakladkaNr).Unprotect "password"
    Next ZakladkaNr
    Sheets(1).Activate
End Sub
