Attribute VB_Name = "Part3"
Sub Iterations()

        ' ==============
        ' VBA Stock Assignment
        ' ==============

        ' III. ITERATIONS (IN PLACE OF THE FOR LOOPS)
        ' -----------------------------------------
        ' I was having a lot of issues with my For Loops. I'm not quite sure why.
            ' I believe that if I knew, then I wouldn't have had so many issues. I've successfully made them in the past.
                ' So after a few hours of trial and error, I searched for an alternative to (at least) return the same results.
        
        ' The following "Evaluate" method was found on the Microsoft Documentation website.
        ' Link: https://docs.microsoft.com/en-us/office/vba/api/excel.application.evaluate
        
        
        ' Solve and Print for "J"
        Range("J2:J" & LastRow) = Evaluate("F2:F" & LastRow & "-C2:C" & LastRow)
        
        ' Solve and Print for "K"
        Range("K2:K" & LastRow) = Evaluate("C2:C" & LastRow & "/F2:F" & LastRow & "/100")
        
        ' Solve and Print for "L"
        Range("L2:L" & LastRow) = Evaluate("G2:G" & LastRow & "*F2:F" & LastRow)

End Sub
