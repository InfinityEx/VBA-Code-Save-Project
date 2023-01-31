Sub global_calculate()
x = 1
y = 0
For i = 2 To 95
    For a = 12 To 17
        If Sheet4.Cells(i, a).Value <> "" Then
            y = y + 1
        End If
    Next
    y = CInt(y / 2)
    For b = 1 To y
            Sheet5.Cells(x, 1).Value = CStr(Sheet4.Cells(i, 1).Value)
            Sheet5.Cells(x, 3).Value = Sheet4.Cells(i, 3).Value
            Sheet5.Cells(x, 5).Value = Sheet4.Cells(i, 5).Value
            Sheet5.Cells(x, 6).Value = Sheet4.Cells(i, 6).Value
            Sheet5.Cells(x, 7).Value = Sheet4.Cells(i, 7).Value
            Sheet5.Cells(x, 8).Value = Sheet4.Cells(i, 8).Value
            Sheet5.Cells(x, 9).Value = Sheet4.Cells(i, 9).Value
            Sheet5.Cells(x, 10).Value = Sheet4.Cells(i, 10).Value
            Sheet5.Cells(x, 11).Value = Sheet4.Cells(i, 11).Value
            Sheet5.Cells(x, 2).Value = Sheet4.Cells(i, 11 + 2 * b - 1).Value
            Sheet5.Cells(x, 4).Value = Sheet4.Cells(i, 12 + 2 * b - 1).Value
            x = x + 1
    Next
    y = 0
Next
End Sub
