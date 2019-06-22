Attribute VB_Name = "Module1"
Sub stockloopmod()

Dim ws As Worksheet
    For Each ws In Worksheets
        ws.Activate

Dim ticker As String
Dim totalvolume As Double
totalvolume = 0
Dim summaryrow As Integer
summaryrow = 2
Dim reset As Integer
reset = 0
Dim yearchange As Double
Dim percentchange As Double
Dim openprice As Double
openprice = Cells(2, 3).Value
Dim closeprice As Double

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

Dim lastrow As Long
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastrow

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    ticker = Cells(i, 1).Value
    Range("I" & summaryrow).Value = ticker
    totalvolume = totalvolume + Cells(i, 7).Value
    Range("L" & summaryrow).Value = totalvolume
    closeprice = Cells(i, 6).Value
    yearlychange = (closeprice - openprice)
    Range("J" & summaryrow).Value = yearlychange
    percentchange = yearlychange / openprice
    Range("K" & summaryrow).Value = percentchange
    Range("K" & summaryrow).NumberFormat = "0.00%"
    summaryrow = summaryrow + 1
    totalvolume = reset
    
    Else
    totalvolume = totalvolume + Cells(i, 7).Value
        
    End If
Next i

For i = 2 To lastrow
    If Cells(i, 10) > 0 Then
    Cells(i, 10).Interior.ColorIndex = 4
    ElseIf Cells(i, 10) < 0 Then
    Cells(i, 10).Interior.ColorIndex = 3
    End If
Next i

Next ws

End Sub
