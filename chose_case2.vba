'Function definition


Public Function WSP_KSZTAŁTU(b As Single, a As Single, T As Single, n As Single, D As Single) As Variant

Dim Pi As Single
Pi = 3.141593

WSP_KSZTAŁTU = (a * b - n * Pi * D ^ 2 / 4) / (T * (2 * a + 2 * b + n * Pi * D))

End Function

'-------------------------------------------------------------------------------------------------------------------------------

Public Function DOP_NAPR(typ As String, b As Single, a As Single, T As Single, n As Single, D As Single) As Variant

Dim S As Single
Dim wynik As Single

S = WSP_KSZTAŁTU(b, a, T, n, D)

Select Case typ
    Case Is = "A"
        wynik = (S ^ 2 + S + 1) / 2
            If wynik > 5 Then
            wynik = 5
            Else:
            wynik = (S ^ 2 + S + 1) / 2
            End If
        
   Case Is = "B"
            wynik = (S ^ 2 + S + 1) / 0.7
            If wynik > 20 Then
            wynik = 20
            Else:
            wynik = (S ^ 2 + S + 1) / 0.7
            End If
        
    Case Is = "C"
            wynik = (S ^ 2 + S + 1) / 0.95
            If wynik > 25 Then
            wynik = 25
            Else:
            wynik = (S ^ 2 + S + 1) / 0.95
            End If

        
    Case Is = "D"
            wynik = S ^ 1.16 * 4.05
            If wynik > 14 Then
            wynik = 14
            Else:
            wynik = S ^ 1.16 * 4.05
            End If
                
    Case Is = "E"
            wynik = S * 6.99
            If wynik > 21 Then
            wynik = 21
            Else:
            wynik = S * 6.99
            End If
        

                        
End Select
    
DOP_NAPR = wynik

End Function


