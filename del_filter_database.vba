Sub Odkryj_Usuń_FilerDatabase()

    Dim ZakladkaNr As Integer
    For ZakladkaNr = 1 To 16
        Sheets(ZakladkaNr).Unprotect "cal"
    Next ZakladkaNr
    Sheets(1).Activate
      

    Dim n As Name
    Dim Count As Integer
    For Each n In ActiveWorkbook.Names
        If Not n.Visible Then
           n.Visible = True
           Count = Count + 1
        End If
        Next n

    Dim sh As Worksheet
    On Error Resume Next
    For Each sh In ActiveWorkbook.Worksheets
    sh.Names("_FilterDatabase").Delete

    Next
    ActiveWorkbook.Names("_FilterDatabase").Delete
    
    Call UkryjZakladkeAdmin
 
End Sub
