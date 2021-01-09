Attribute VB_Name = "Part4"
Sub Conditionals()

        ' ==============
        ' VBA Stock Assignment
        ' ==============

        ' IV. CONDTIONALS (Conditional Formatting for YearlyChange column)
        ' -----------------------------------------------------------
        ' The following method for looping using "Select Cell" and "ActiveCell.Offset" was found on Stack Overflow.
        ' Link: https://stackoverflow.com/questions/18654144/vba-code-to-color-cells-having-negative-values
            
            
        ' Select cell
            Range("J2").Select
            
        ' For Loop
            
        For i = 2 To LastRow
                
                If ((ActiveCell.Value) > 0) Then
                    ActiveCell.Interior.ColorIndex = 4
                ElseIf ((ActiveCell.Value) < 0) Then
                    ActiveCell.Interior.ColorIndex = 3
            
                End If
     
                    ActiveCell.Offset(1, 0).Select
     
        Next i

End Sub
