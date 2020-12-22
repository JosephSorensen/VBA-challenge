Attribute VB_Name = "YearlyChange"
' Declare variables and arrays using Dim and Integer, Long, Single, Double, Object, etc.
' Create Loops (Iterations and Conditionals)


' Yearly Change

Sub CminusF()

' J = C - F

    Dim LastRow As Long
    LastRow = Cells(Rows.Count, "C").End(xlUp).Row
    Range("J2:J" & LastRow) = Evaluate("C2:C" & LastRow & "-F2:F" & LastRow)

End Sub
