Sub HighlightAlternatingCells()
    Dim i As Long
    Dim lastRow As Long

    ' Get the last row with data in Column A
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row

    ' Loop through all rows in Column A
    For i = 1 To lastRow
        ' Check if the row number is odd or even
        If i Mod 2 = 1 Then
            ' If odd, highlight in red
            Cells(i, 1).Interior.Color = RGB(255, 0, 0) ' Red
        Else
            ' If even, highlight in blue
            Cells(i, 1).Interior.Color = RGB(0, 0, 255) ' Blue
        End If
    Next i
End Sub
