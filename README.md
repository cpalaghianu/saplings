# saplings

Function quadrat(suprafata As Integer, x, y, x0, y0, x_max, y_max As Single)
' face incadrarea intr-un quadrat in functie de x, y, coordonatele pct
' suprafata - suprafata din care face parte
' x0, y0 - dim quadrat
' x_max, y_max - dimensiunile maxime ale suprafetei
Dim x1, y1, a, b As Single

  If (y = y_max) Then
        y_1 = Int(y / y0) - 1
        Else
        y_1 = Int(y / y0)
        End If
        
        If (x = x_max) Then
        x_1 = Int(x / x0) - 1
        Else
        x_1 = Int(x / x0)
        End If
        b = Int(x_max / x0)
        a = Int(y_max / y0)
        nr_quad = a * b
        quadrat = b * y_1 + x_1 + 1 + (suprafata - 1) * nr_quad
     'abatere = Application.WorksheetFunction.StDev(domeniu)
End Function

Sub Calcul()

Dim i, j, k, sp1, sp2, ev_zero, col As Integer
Dim aux, jac, sok, euclid As Single
Dim zona As Variant

j = 1

While Cells(2, j).Value <> ""
j = j + 1 ' find the first colums
Wend
col = j ' the first colums - the date is printed starting from here

Cells(2, col).Value = "."
Cells(2, col + 1).Value = "Valori indici"
Cells(3, col + 1).Value = "Jaccard"
Cells(4, col + 1).Value = "Sokal"
Cells(5, col + 1).Value = "Euclid"
Cells(6, col + 1).Value = "Nr. quad"
Cells(7, col + 1).Value = "Ev. zero"
col = col + 2

'analyse for all combinations of the 7 species, taken by two
x = 2 ' the rellation between species x-1 and y-1

While x < 8

y = x + 1

While y < 9
'the calculations for Jaccard, Sokal & Mitchell indices
' and the Euclidian distance divided by the sum of the individuals by species

    i = 3: jac = 0: sok = 0: k = 0: ev_zero = 0: euclid = 0
    Cells(2, col).Value = (x - 1) * 10 + (y - 1)

    While Cells(i, x).Value <> ""  ' till it finds the date - the end of the column
    sp1 = Cells(i, x).Value
    sp2 = Cells(i, y).Value
    k = k + 1 'counting events - quadrats
    
    If (sp1 + sp2 > 0) Then 'for euclidian distance - case of division by zero
    euclid = euclid + (((sp1 - sp2) ^ 2) / (sp1 + sp2))
    End If
    
    If (sp1 * sp2 > 0) Then
    jac = jac + 1
    sok = sok + 1
    ElseIf (sp1 + sp2 = 0) Then
        sok = sok + 1
        ev_zero = ev_zero + 1 'events-quadrats in which no species are present
    End If
    i = i + 1

    Wend

Cells(3, col).Value = jac / (k - ev_zero)
Cells(4, col).Value = sok / k
Cells(5, col).Value = euclid
Cells(6, col).Value = k
Cells(7, col).Value = ev_zero
y = y + 1
col = col + 1

Wend ' ends while ptr y
x = x + 1

Wend ' ends while ptr x



End Sub
