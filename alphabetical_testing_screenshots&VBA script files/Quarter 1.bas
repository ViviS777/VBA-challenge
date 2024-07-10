Attribute VB_Name = "模块1"
Sub Q12()
Dim sheet1 As Worksheet
Dim i, j, k As Integer
Dim m, su, su1 As Double
Dim closeQ1 As Double
Dim openQ1 As Double
Dim Volume As Double

Set sheet1 = ThisWorkbook.Worksheets("Q1")
columnNumber = 1
columnLength1 = sheet1.Cells(sheet1.Rows.Count, columnNumber).End(xlUp).Row
sheet1.Cells(1, 9) = "Ticker"
sheet1.Cells(1, 10) = "Quarterly Change"
sheet1.Cells(1, 11) = "Percentage Change"
sheet1.Cells(1, 12) = "Total Stock Volume"
ticker = AAWJDJSK
k = 2
' Extract Tickers----------------------------------

For i = 2 To columnLength1
    If sheet1.Cells(i, 1) <> ticker Then
    sheet1.Cells(k, 9) = sheet1.Cells(i, 1)
    ticker = sheet1.Cells(i, 1)
    k = k + 1
    End If
Next i
'--------------------------------------------------
' Quarterly Change/Percentage Change/Total Stock Volume----------------------------------
columnLength2 = columnLength1 + 1
openQ1 = sheet1.Cells(2, 3)
k = 2
su = 2
For i = 3 To columnLength2
    If sheet1.Cells(i, 1) <> sheet1.Cells(i - 1, 1) Then
    closeQ1 = sheet1.Cells(i - 1, 6)
    sheet1.Cells(k, 10) = closeQ1 - openQ1
    sheet1.Cells(k, 11) = (closeQ1 - openQ1) / openQ1
        Volume = 0
        su1 = i - 1
        For m = su To su1
        Volume = Volume + sheet1.Cells(m, 7)
        sheet1.Cells(k, 12) = Volume
        Next m
        su = i
    k = k + 1
    openQ1 = sheet1.Cells(i, 3)
    End If
Next i
' Format----------------------------------
sheet1.Range("K2:K5000").NumberFormat = "0.00%"

Dim threshold As Double
threshold = 0
Dim rng As Range
Set rng = sheet1.Range("J2:J" & sheet1.Cells(sheet1.Rows.Count, 10).End(xlUp).Row) ' 假设第三列是C列

For Each cell In rng
    If cell.Value > threshold Then
        cell.Interior.Color = RGB(0, 255, 0)
    End If
    If cell.Value < threshold Then
        cell.Interior.Color = RGB(255, 0, 0)
    End If
Next cell

' Max/Min----------------------------------
Dim maxValue As Double
Dim maxPosition As Long
Dim ws As Worksheet

Set rng = sheet1.Range("k2:k5000")

sheet1.Cells(2, 15) = "Greateat % Increase"
sheet1.Cells(3, 15) = "Greateat % Decrease"
sheet1.Cells(4, 15) = "Greateat Total Volume"
sheet1.Cells(1, 16) = "Ticker"
sheet1.Cells(1, 17) = "Value"

sheet1.Range("Q2:Q3").NumberFormat = "0.00%"
maxValue = Application.WorksheetFunction.Max(rng)
sheet1.Cells(2, 17) = maxValue
maxPosition = Application.WorksheetFunction.Match(maxValue, rng, 0)
sheet1.Cells(2, 16) = sheet1.Cells(maxPosition, 9)


minValue = Application.WorksheetFunction.Min(rng)
sheet1.Cells(3, 17) = minValue
minPosition = Application.WorksheetFunction.Match(minValue, rng, 0)
sheet1.Cells(3, 16) = sheet1.Cells(minPosition, 9)


Set rng = sheet1.Range("l2:l5000")
maxValue = Application.WorksheetFunction.Max(rng)
sheet1.Cells(4, 17) = maxValue
maxPosition = Application.WorksheetFunction.Match(maxValue, rng, 0)
sheet1.Cells(4, 16) = sheet1.Cells(maxPosition, 9)




End Sub



