'procedure calls diferent macros if value in cell F3 change into specific value

Sub Worksheet_Change(ByVal Target As Range)
Set Target = Range("F3")

If Target.Value = 5 Then
    Call Makro1
End If

If Target.Value = 8 Then
    Call Makro2
End If

If Target.Value = 10 Then
    Call Makro3
End If

If Target.Value = 15 Then
    Call Makro4
End If

If Target.Value = 20 Then
    Call Makro5
End If

End Sub
