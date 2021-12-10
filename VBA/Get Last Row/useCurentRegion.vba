Sub useCurrentRegion()
    'Advantages:
    'Get the range easily

    'Disadvantages:
    'Doesn't work on protacted sheet
    'Data must be adjacent
    
    Dim rg As Range
    Set rg = Sheet1.Range("A1").CurrentRegion

    rg.Select

    Dim lastrow As Long
    lastrow = rg.Rows(rg.Rows.Count).Row

    rg.Rows(rg.Rows.Count).Select

End Sub