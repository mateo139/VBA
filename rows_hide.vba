Sub Zwin()

    ActiveSheet.Unprotect "password"
    Range("11:52,54:95,97:138,140:181,183:224,226:267,269:310,312:353").Select
    Selection.EntireRow.Hidden = True
    ActiveSheet.Protect "cal"
    
End Sub
