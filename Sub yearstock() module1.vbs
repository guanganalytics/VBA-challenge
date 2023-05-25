Sub yearstock()

Dim ticker As String
Dim yearlychange As Double
Dim percentchange As Double
Dim totalstock As Double
Dim line As Integer
Dim originalprice As Double



Dim maxrow As Double
maxrow = Range("A" & Rows.Count).End(xlUp).Row
line = 2

yearlychange = 0
totalstock = 0
originalprice = 24.44

Cells(1, 9).value = "Ticker"
Cells(1, 10).value = "Yearly Change"
Cells(1, 11).value = "Percentage Change"
Cells(1, 12).value = "Total Stock Volume"

For i = 2 To maxrow 'could use maxrow,use smaller# for testing purpose

    If Cells(i + 1, 1) <> Cells(i, 1) Then
    
    Cells(line + 1, 9).value = ticker
    ticker = Cells(i + 1, 1)
    
    originalprice = Cells(i + 1, 3)
    yearlychange = Cells(i, 6) - originalprice

    totalchange = totalchange + Cells(i, 7)
    Cells(line, 12).value = totalchange
    
    yearlychange = 0
    Cells(line + 1, 10).value = yearlychange
    percentchange = 0
    Cells(line + 1, 11) = percentchange
    totalchange = 0
    Cells(line + 1, 10) = totalchange
        
        
     line = line + 1
    

    Else
    Cells(line, 9).value = ticker
    ticker = Cells(i, 1)
    Cells(line, 9).value = ticker
    
    yearlychange = Cells(i + 1, 6) - originalprice
    Cells(line, 10).value = yearlychange
    
        If yearlychange >= 0 Then
        Cells(line, 10).Interior.ColorIndex = 4
        Else
        Cells(line, 10).Interior.ColorIndex = 3
        End If
            
    percentchange = yearlychange / originalprice
    Cells(line, 11) = percentchange
    Range("K" & line).Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.00%"
    
    totalchange = totalchange + Cells(i, 7)
    Cells(line, 12).value = totalchange
    
   End If
   
Next i

Cells(1, 16) = "Ticker"
Cells(1, 17) = "Value"
Cells(2, 15) = "Greatest % Increase"
Cells(3, 15) = "Greatest % Decrease"
Cells(4, 15) = "Greatest Total Volume"


Dim greatestincrease As String
Dim greatestdecrease As String
Dim greatesttotal As String
Dim maxrow2 As Double

maxrow2 = Range("k" & Rows.Count).End(xlUp).Row


greatestincrease = WorksheetFunction.Max(Range("k2:k" & line))
greatestdecrease = WorksheetFunction.Min(Range("k2:k" & line))
greatesttotal = WorksheetFunction.Max(Range("L2:L" & line))

Cells(2, 17) = greatestincrease
Range("q2").Select
Selection.Style = "Percent"
Selection.NumberFormat = "0.00%"

  
Cells(3, 17) = greatestdecrease
Range("q3").Select
Selection.Style = "Percent"
Selection.NumberFormat = "0.00%"
    
Cells(4, 17) = greatesttotal

For j = 2 To maxrow2
If Cells(j, 11) = greatestincrease Then
Cells(2, 16) = Cells(j, 9)
ElseIf Cells(j, 11) = greatestdecrease Then
Cells(3, 16) = Cells(j, 9)
ElseIf Cells(j, 12) = greatesttotal Then
Cells(4, 16) = Cells(j, 9)

End If
Next j


End Sub
