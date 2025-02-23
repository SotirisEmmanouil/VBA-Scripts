Sub MergeSameCells()
    Dim i As Long
    Dim lastRow As Long
    Dim mergeStart As Long
    lastRow = Cells(Rows.Count, 2).End(xlUp).Row ' Find the last row in column 1 (A)
   
    i = 1
    While i <= lastRow
        mergeStart = i
        ' Check for consecutive cells with the same value
        While i < lastRow And Cells(i, 2).Value = Cells(i + 1, 2).Value
            i = i + 1
        Wend
       
        ' If more than 1 cell has the same value, merge them
        If i > mergeStart Then
            Range(Cells(mergeStart, 2), Cells(i, 2)).Merge
        End If
       
        ' Move to the next cell to continue checking
        i = i + 1
    Wend
End Sub
