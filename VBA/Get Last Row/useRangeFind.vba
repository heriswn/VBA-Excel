Sub useRangeFind()
    'The right way!

    Dim lastrow As Long, lastcolumn As Long

    lastrow = Sheet1.Cells.Find(What:="*" _
                        , Lookat:=xlPart _
                        , LookIn:=xlFormulas _
                        , searchorder:=xlByRows _
                        , searchdirection:=xlPrevious).Row

    lastcolumn = Sheet1.Cells.Find(What:="*" _
                        , Lookat:=xlPart _
                        , LookIn:=xlFormulas _
                        , searchorder:=xlByColumns _
                        , searchdirection:=xlPrevious).Column

    Dim rg As Range
    With Sheet1
        Set rg = .Range(.Cells(1, 1), .Cells(lastrow, lastcolumn))
    End With

    rg.Select

End Sub