Attribute VB_Name = "Ä£¿é10"
Sub ExtractQuarterDataF()
    Dim sheetF As Worksheet
    Dim sheet1 As Worksheet
    Dim sheet2 As Worksheet
    Dim sheet3 As Worksheet
    Dim sheet4 As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim iQ1, iQ2, iQ3, iQ4 As Long
    Dim k As Long

    Set sheetF = ThisWorkbook.Worksheets("F")
    Set sheet1 = ThisWorkbook.Worksheets("Q1")
    Set sheet2 = ThisWorkbook.Worksheets("Q2")
    Set sheet3 = ThisWorkbook.Worksheets("Q3")
    Set sheet4 = ThisWorkbook.Worksheets("Q4")
    sheet1.Cells.ClearContents
    sheet1.Cells.ClearFormats
    sheet2.Cells.ClearContents
    sheet2.Cells.ClearFormats
    sheet3.Cells.ClearContents
    sheet3.Cells.ClearFormats
    sheet4.Cells.ClearContents
    sheet4.Cells.ClearFormats
    sheet1.Range("B:B").NumberFormat = "yyyy-mm-dd"
    sheet2.Range("B:B").NumberFormat = "yyyy-mm-dd"
    sheet3.Range("B:B").NumberFormat = "yyyy-mm-dd"
    sheet4.Range("B:B").NumberFormat = "yyyy-mm-dd"
    Set rng = sheetF.Range("B2:B40000")

    ' Copy header row
    sheet1.Rows(1).Value = sheetF.Rows(1).Value

    iQ1 = 2
    iQ2 = 2
    iQ3 = 2
    iQ4 = 2
    k = 2
    sheet1.Cells(1, 1) = sheetF.Cells(1, 1)
    sheet1.Cells(1, 2) = sheetF.Cells(1, 2)
    sheet1.Cells(1, 3) = sheetF.Cells(1, 3)
    sheet1.Cells(1, 4) = sheetF.Cells(1, 4)
    sheet1.Cells(1, 5) = sheetF.Cells(1, 5)
    sheet1.Cells(1, 6) = sheetF.Cells(1, 6)
    sheet1.Cells(1, 7) = sheetF.Cells(1, 7)
    sheet2.Cells(1, 1) = sheetF.Cells(1, 1)
    sheet2.Cells(1, 2) = sheetF.Cells(1, 2)
    sheet2.Cells(1, 3) = sheetF.Cells(1, 3)
    sheet2.Cells(1, 4) = sheetF.Cells(1, 4)
    sheet2.Cells(1, 5) = sheetF.Cells(1, 5)
    sheet2.Cells(1, 6) = sheetF.Cells(1, 6)
    sheet2.Cells(1, 7) = sheetF.Cells(1, 7)
    sheet3.Cells(1, 1) = sheetF.Cells(1, 1)
    sheet3.Cells(1, 2) = sheetF.Cells(1, 2)
    sheet3.Cells(1, 3) = sheetF.Cells(1, 3)
    sheet3.Cells(1, 4) = sheetF.Cells(1, 4)
    sheet3.Cells(1, 5) = sheetF.Cells(1, 5)
    sheet3.Cells(1, 6) = sheetF.Cells(1, 6)
    sheet3.Cells(1, 7) = sheetF.Cells(1, 7)
    sheet4.Cells(1, 1) = sheetF.Cells(1, 1)
    sheet4.Cells(1, 2) = sheetF.Cells(1, 2)
    sheet4.Cells(1, 3) = sheetF.Cells(1, 3)
    sheet4.Cells(1, 4) = sheetF.Cells(1, 4)
    sheet4.Cells(1, 5) = sheetF.Cells(1, 5)
    sheet4.Cells(1, 6) = sheetF.Cells(1, 6)
    sheet4.Cells(1, 7) = sheetF.Cells(1, 7)
   
    For Each cell In rng
        If IsDate(cell.Value) Then
            cell.Value = Format(cell.Value, "yyyy-mm-dd")
        Else
            If cell.Value <> "" Then
                Dim year As Integer, month As Integer, day As Integer
                year = CInt(Left(cell.Value, 4))
                month = CInt(Mid(cell.Value, 5, 2))
                day = CInt(Right(cell.Value, 2))
                cell.Value = Format(DateSerial(year, month, day), "yyyy-mm-dd")
            End If
        End If


        If VBA.month(cell.Value) <= 3 Then
            sheet1.Cells(iQ1, 1) = sheetF.Cells(k, 1)
            sheet1.Cells(iQ1, 2) = sheetF.Cells(k, 2)
            sheet1.Cells(iQ1, 3) = sheetF.Cells(k, 3)
            sheet1.Cells(iQ1, 4) = sheetF.Cells(k, 4)
            sheet1.Cells(iQ1, 5) = sheetF.Cells(k, 5)
            sheet1.Cells(iQ1, 6) = sheetF.Cells(k, 6)
            sheet1.Cells(iQ1, 7) = sheetF.Cells(k, 7)
            iQ1 = iQ1 + 1
        End If
        If VBA.month(cell.Value) > 3 And VBA.month(cell.Value) <= 6 Then
            sheet2.Cells(iQ2, 1) = sheetF.Cells(k, 1)
            sheet2.Cells(iQ2, 2) = sheetF.Cells(k, 2)
            sheet2.Cells(iQ2, 3) = sheetF.Cells(k, 3)
            sheet2.Cells(iQ2, 4) = sheetF.Cells(k, 4)
            sheet2.Cells(iQ2, 5) = sheetF.Cells(k, 5)
            sheet2.Cells(iQ2, 6) = sheetF.Cells(k, 6)
            sheet2.Cells(iQ2, 7) = sheetF.Cells(k, 7)
            iQ2 = iQ2 + 1
        End If
        If VBA.month(cell.Value) > 6 And VBA.month(cell.Value) <= 9 Then
            sheet3.Cells(iQ3, 1) = sheetF.Cells(k, 1)
            sheet3.Cells(iQ3, 2) = sheetF.Cells(k, 2)
            sheet3.Cells(iQ3, 3) = sheetF.Cells(k, 3)
            sheet3.Cells(iQ3, 4) = sheetF.Cells(k, 4)
            sheet3.Cells(iQ3, 5) = sheetF.Cells(k, 5)
            sheet3.Cells(iQ3, 6) = sheetF.Cells(k, 6)
            sheet3.Cells(iQ3, 7) = sheetF.Cells(k, 7)
            iQ3 = iQ3 + 1
        End If
        If VBA.month(cell.Value) > 9 And VBA.month(cell.Value) <= 12 Then
            sheet4.Cells(iQ4, 1) = sheetF.Cells(k, 1)
            sheet4.Cells(iQ4, 2) = sheetF.Cells(k, 2)
            sheet4.Cells(iQ4, 3) = sheetF.Cells(k, 3)
            sheet4.Cells(iQ4, 4) = sheetF.Cells(k, 4)
            sheet4.Cells(iQ4, 5) = sheetF.Cells(k, 5)
            sheet4.Cells(iQ4, 6) = sheetF.Cells(k, 6)
            sheet4.Cells(iQ4, 7) = sheetF.Cells(k, 7)
            iQ4 = iQ4 + 1
        End If
        k = k + 1
    Next cell
End Sub






