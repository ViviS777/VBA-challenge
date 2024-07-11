Attribute VB_Name = "__4"
Sub Q4()
Dim sheet4 As Worksheet
Dim i, j, k As Integer
Dim m, su, su1 As Double
Dim closeQ4 As Double
Dim openQ4 As Double
Dim Volume As Double

Set sheet4 = ThisWorkbook.Worksheets("Q4")
columnNumber = 1
columnLength1 = sheet4.Cells(sheet4.Rows.Count, columnNumber).End(xlUp).Row
sheet4.Cells(1, 9) = "Ticker"
sheet4.Cells(1, 10) = "Quarterly Change"
sheet4.Cells(1, 11) = "Percentage Change"
sheet4.Cells(1, 12) = "Total Stock Volume"
ticker = AAWJDJSK
k = 2
' Extract Tickers----------------------------------

For i = 2 To columnLength1
    If sheet4.Cells(i, 1) <> ticker Then
    sheet4.Cells(k, 9) = sheet4.Cells(i, 1)
    ticker = sheet4.Cells(i, 1)
    k = k + 1
    End If
Next i
'--------------------------------------------------
' Quarterly Change/Percentage Change/Total Stock Volume----------------------------------
columnLength2 = columnLength1 + 1
openQ4 = sheet4.Cells(2, 3)
k = 2
su = 2
For i = 3 To columnLength2
    If sheet4.Cells(i, 1) <> sheet4.Cells(i - 1, 1) Then
    closeQ4 = sheet4.Cells(i - 1, 6)
    sheet4.Cells(k, 10) = closeQ4 - openQ4
    sheet4.Cells(k, 11) = (closeQ4 - openQ4) / openQ4
        Volume = 0
        su1 = i - 1
        For m = su To su1
        Volume = Volume + sheet4.Cells(m, 7)
        sheet4.Cells(k, 12) = Volume
        Next m
        su = i
    k = k + 1
    openQ4 = sheet4.Cells(i, 3)
    End If
Next i
' Format----------------------------------
sheet4.Range("K2:K5000").NumberFormat = "0.00%"
Dim threshold As Double
threshold = 0
Dim rng As Range
Set rng = sheet4.Range("J2:J" & sheet4.Cells(sheet4.Rows.Count, 10).End(xlUp).Row) ' ______C_ÁÐÊÇCÁÐ

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

Set rng = sheet4.Range("k2:k5000")

sheet4.Cells(2, 15) = "Greateat % Increase"
sheet4.Cells(3, 15) = "Greateat % Decrease"
sheet4.Cells(4, 15) = "Greateat Total Volume"
sheet4.Cells(1, 16) = "Ticker"
sheet4.Cells(1, 17) = "Value"

sheet4.Range("Q2:Q3").NumberFormat = "0.00%"
maxValue = Application.WorksheetFunction.Max(rng)
sheet4.Cells(2, 17) = maxValue
maxPosition = Application.WorksheetFunction.Match(maxValue, rng, 0)
sheet4.Cells(2, 16) = sheet4.Cells(maxPosition, 9)


minValue = Application.WorksheetFunction.Min(rng)
sheet4.Cells(3, 17) = minValue
minPosition = Application.WorksheetFunction.Match(minValue, rng, 0)
sheet4.Cells(3, 16) = sheet4.Cells(minPosition, 9)


Set rng = sheet4.Range("l2:l5000")
maxValue = Application.WorksheetFunction.Max(rng)
sheet4.Cells(4, 17) = maxValue
maxPosition = Application.WorksheetFunction.Match(maxValue, rng, 0)
sheet4.Cells(4, 16) = sheet4.Cells(maxPosition, 9)
End Sub
