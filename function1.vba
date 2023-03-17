Public Function DOP_ALFA(T As Variant, a As Variant) As Variant
Dim stala As Variant

If T = 5 Then
stala = 1500
Else:
    If T = 10 Then
    stala = 3000
    Else:
        If T = 15 Then
        stala = 5000
        Else:
            If T = 20 Then
            stala = 6500
            Else: 'MsgBox "podaj wartość t= 5 ,10 ,15, 20 mm"
            End If
        End If
    End If
End If

DOP_ALFA = stala / a

If DOP_ALFA > 40 Then
DOP_ALFA = 40
Else:
End If

End Function
