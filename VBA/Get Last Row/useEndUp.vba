Sub useEndUp()
    'Advantages:
    'Works on a protacted sheet

    'Disadvantages:
    'Doesn't work well with jagged data
    'You need to build the range

    Dim lastrow As Long

    lastrow = Sheet1.Cells(Sheet1.Columns.Count, 1).End(xlUp).Row
    lastcolumn = Sheet1.Cells(1, Sheet1.Columns.Count).End(xlToLeft).Column

    Sheet1.Columns(lastcolumn).Select

    Dim rgFull As Range
    With Sheet1
    Set rgFull = .Range(.Cells(1, 1), .Cells(lastrow, lastcolumn))
    End With

    rgFull.Select

End Sub