Attribute VB_Name = "TotalVolume"
' Total Volume

Sub TotalVolume()

' L = F * G


 Dim LastRow As Long
    LastRow = Cells(Rows.Count, "F").End(xlUp).Row
    Range("L2:L" & LastRow) = Evaluate("F2:F" & LastRow & "*G2:G" & LastRow)


End Sub
