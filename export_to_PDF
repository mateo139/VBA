 Sub pdf()
' pdf Makro
    ActiveSheet.Unprotect "password"
    ActiveSheet.Range("$K$3:$K$338").AutoFilter Field:=1, Criteria1:="<>"
    
    Sheets(Array("sheet1", "sheet2")).Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
    "" & Worksheets("sheet0").Range("U38").Value & ".pdf", Quality:= _
    xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
    OpenAfterPublish:=True
    Sheets(sheet2").Select
    ActiveSheet.Protect "password"
End Sub
