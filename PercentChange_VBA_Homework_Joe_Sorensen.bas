Attribute VB_Name = "PercentChange"
' Percent Change


Sub PercentChange()

' K = C / F

    'Range("K2:K760192").Formula = "=C2/F2" ---- easy Excel formula, but not VBA
    
    Dim LastRow As Long
    LastRow = Cells(Rows.Count, "C").End(xlUp).Row
    Range("K2:K" & LastRow) = Evaluate("C2:C" & LastRow & "/F2:F" & LastRow)
    
' Can't seem to figure out how to divide the result by 200 in order to move the decimal place over twice.
' When I format the cells to percentage, it makes the number super inflated, so I left it as a decimal for now.
' Also can't seem to figure out how to remove the 1 if it should be a 0


End Sub
