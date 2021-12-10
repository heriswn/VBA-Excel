Sub useSpecialCells()
    'Adventages:
    'Works with jagged data
    'Gets the last cell(row and column)

    'Disadventages:
    'Formated cells are included
    'Doesn't work protected sheet
    'You have to build the range
    
    Dim rg As Range
    Set rg = Sheet1.Cells.SpecialCells(xlCellTypeLastCell)

    Dim rgFull As Range
    Set rgFull = Sheet1.Range(Sheet1.Cells(1, 1), rg)

    rgFull.Select

End Sub