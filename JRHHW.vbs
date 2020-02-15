#please run twice- when I ran it twice it worked, clicking
#below in the script. It works!

Sub stockdata():


Cells(1, 9) = ("Ticker")
Cells(1, 10) = ("Yearly Change")
Cells(1, 11) = ("Percent Change")
Cells(1, 12) = ("Total Stock Volume")
Cells(1, 13) = ("Year Open")
Cells(1, 14) = ("Year Close")
Cells(2, 13) = Cells(2, 3)

Dim ticker As String
Dim row As Integer
row = 1
Dim openpricenext As Double
Dim closeprice As Double

lastrow1 = Cells(Rows.Count, "A").End(xlUp).row
lastrow2 = Cells(Rows.Count, "I").End(xlUp).row

For i = 2 To lastrow1

    If Cells(i, 1) <> Cells(i + 1, 1) Then
        ticker = Cells(i, 1)
        row = row + 1
        closeprice = Cells(i, 6)
        openpricenext = Cells(i + 1, 3)
        Cells(row, 9) = ticker
        Cells(row, 14) = closeprice
        Cells(row + 1, 13) = openpricenext
        
        
End If

Next i


For i = 2 To lastrow2

Cells(i, 10) = (Cells(i, 13) - Cells(i, 14)) * (-1)

Cells(i, 11).NumberFormat = "0.00%"

If Cells(i, 13) = 0 Then
    Cells(i, 11) = ("0%")
Else
    Cells(i, 11) = Cells(i, 10) / Cells(i, 13)

End If

Next i

Dim number As Integer
Dim sum As Double

number = 1
volume = Cells(2, 7)


For i = 2 To lastrow1

If Cells(i, 1) = Cells(i + 1, 1) Then
 volume = volume + Cells(i, 7)
 
Else
    volume = volume + Cells(i, 7)
    number = number + 1
    Cells(number, 12) = volume
    volume = 0
    
End If
    
Next i

For i = 2 To lastrow2

Dim change As Double

change = Cells(i, 10).Value

If change > 0 Then
    Cells(i, 10).Interior.ColorIndex = 4
    
ElseIf change < ("0") Then
    Cells(i, 10).Interior.ColorIndex = 3
 
Else
    Cells(i, 10).Interior.ColorIndex = 6
    
End If
Next i

For i = 2 To lastrow2
changeamount = Cells(i, 14) - Cells(i, 13)
Cells(i, 15) = changeamount
Next i



Dim biggestchange As Double
Dim biggestchanger As String


biggestchange = Cells(2, 11).Value


For i = 3 To lastrow2
If Cells(i, 11) > biggestchange Then
    biggestchange = Cells(i, 11).Value
    biggestchanger = Cells(i, 9)

End If

Next i

biggestloss = Cells(2, 11)

For i = 3 To lastrow2

If Cells(i, 11) < biggestloss Then
    biggestloss = Cells(i, 11).Value
    biggestloser = Cells(i, 9)
End If
Next i

Dim greatestvolume As Double
Dim greatestvolumer As String
greatestvolume = Cells(2, 12)


For i = 3 To lastrow2

If Cells(i, 12) > greatestvolume Then
    greatestvolume = Cells(i, 12).Value
    greatestvolumer = Cells(i, 9)
    
End If
Next i

Cells(1, 17) = ("Ticker")
Cells(1, 18) = ("Value")

Cells(2, 16) = ("Greatest % Increase")
Cells(2, 17) = biggestchanger
Cells(2, 18) = biggestchange
Cells(2, 18).NumberFormat = "0.00%"

Cells(3, 16) = ("Greatest % Decrease")
Cells(3, 17) = biggestloser
Cells(3, 18) = biggestloss
Cells(3, 18).NumberFormat = "0.00%"

Cells(4, 16) = ("Greatest Total Volume")
Cells(4, 17) = greatestvolumer
Cells(4, 18) = greatestvolume




End Sub
