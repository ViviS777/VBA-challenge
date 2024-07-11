Attribute VB_Name = "__2"
Sub Q2()
Dim sheet2 As Worksheet
Dim i, j, k As Integer
Dim m, su, su1 As Double
Dim closeQ2 As Double
Dim openQ2 As Double
Dim Volume As Double

Set sheet2 = ThisWorkbook.Worksheets("Q2")
columnNumber = 1
columnLength1 = sheet2.Cells(sheet2.Rows.Count, columnNumber).End(xlUp).Row
sheet2.Cells(1, 9) = "Ticker"
sheet2.Cells(1, 10) = "Quarterly Change"
sheet2.Cells(1, 11) = "Percentage Change"
sheet2.Cells(1, 12) = "Total Stock Volume"
ticker = AAWJDJSK
k = 2
' Extract Tickers----------------------------------

For i = 2 To columnLength1
    If sheet2.Cells(i, 1) <> ticker Then
    sheet2.Cells(k, 9) = sheet2.Cells(i, 1)
    ticker = sheet2.Cells(i, 1)
    k = k + 1
    End If
Next i
'--------------------------------------------------
' Quarterly Change/Percentage Change/Total Stock Volume----------------------------------
columnLength2 = columnLength1 + 1
openQ2 = sheet2.Cells(2, 3)
k = 2
su = 2
For i = 3 To columnLength2
    If sheet2.Cells(i, 1) <> sheet2.Cells(i - 1, 1) Then
    closeQ2 = sheet2.Cells(i - 1, 6)
    sheet2.Cells(k, 10) = closeQ2 - openQ2
    sheet2.Cells(k, 11) = (closeQ2 - openQ2) / openQ2
        Volume = 0
        su1 = i - 1
        For m = su To su1
        Volume = Volume + sheet2.Cells(m, 7)
        sheet2.Cells(k, 12) = Volume
        Next m
        su = i
    k = k + 1
    openQ2 = sheet2.Cells(i, 3)
    End If
Next i
' Format----------------------------------
sheet2.Range("K2:K5000").NumberFormat = "0.00%"
Dim threshold As Double
threshold = 0
Dim rng As Range
Set rng = sheet2.Range("J2:J" & sheet2.Cells(sheet2.Rows.Count, 10).End(xlUp).Row) ' ______C_ÁÐÊÇCÁÐ

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

Set rng = sheet2.Range("k2:k5000")

sheet2.Cells(2, 15) = "Greateat % Increase"
sheet2.Cells(3, 15) = "Greateat % Decrease"
sheet2.Cells(4, 15) = "Greateat Total Volume"
sheet2.Cells(1, 16) = "Ticker"
sheet2.Cells(1, 17) = "Value"

sheet2.Range("Q2:Q3").NumberFormat = "0.00%"
maxValue = Application.WorksheetFunction.Max(rng)
sheet2.Cells(2, 17) = maxValue
maxPosition = Application.WorksheetFunction.Match(maxValue, rng, 0)
sheet2.Cells(2, 16) = sheet2.Cells(maxPosition, 9)


minValue = Application.WorksheetFunction.Min(rng)
sheet2.Cells(3, 17) = minValue
minPosition = Application.WorksheetFunction.Match(minValue, rng, 0)
sheet2.Cells(3, 16) = sheet2.Cells(minPosition, 9)


Set rng = sheet2.Range("l2:l5000")
maxValue = Application.WorksheetFunction.Max(rng)
sheet2.Cells(4, 17) = maxValue
maxPosition = Application.WorksheetFunction.Match(maxValue, rng, 0)
sheet2.Cells(4, 16) = sheet2.Cells(maxPosition, 9)
End Sub



