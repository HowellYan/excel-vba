Sub test()
    Dim str
    Dim i, j, i2
    i = 1
    i2 = 1
    j = 1
    isHave = 0
    cellNum = 1
    For r = 1 To Worksheets(2).UsedRange.Rows.Count
        cellNum = 1
        For r2 = 1 To Worksheets(2).UsedRange.Rows.Count
            str = Worksheets(2).Cells(r, 1).Value
            str1 = Worksheets(2).Cells(r2, 1).Value
            isHave = 0
            If str = str1 Then
                cellNum = cellNum + 1
                For r3 = 1 To Worksheets(3).UsedRange.Rows.Count
                    str2 = Worksheets(3).Cells(r3, 1).Value
                    If str2 = str Then
                        isHave = 1
                        Worksheets(3).Cells(r3, cellNum).Value = Worksheets(2).Cells(r2, 2).Value
                    End If
                Next
                If isHave = 0 Then
                    Worksheets(3).Cells(j, 1).Value = str
                    Worksheets(3).Cells(j, cellNum).Value = Worksheets(2).Cells(r2, 2).Value
                    j = j + 1
                End If
            End If
            i2 = i2 + 1
        Next
        i = i + 1
    Next
End Sub
