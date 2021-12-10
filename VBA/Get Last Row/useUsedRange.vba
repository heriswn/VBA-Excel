Sub useUsedRange()
    'Advantages:
    'Gets the range
    'Work with protected sheet

    'Disadvantages:
    'Formated cell are included
    'Cann't specify a section of worksheet

    Dim rg As Range
    Set rg = Sheet1.UsedRange

    rg.Select

    Dim lastrow As Long
    lastrow = rg.Rows(rg.Rows.Count).Row

    Sheet1.Rows(lastrow).Select

End Sub
