Attribute VB_Name = "Part2"
Sub VariableAssignments()

        ' ==============
        ' VBA Stock Assignment
        ' ==============

 ' II. VARIABLE ASSIGNMENTS (AND FORMULAS)
        ' -----------------------------------------
        LastRow = Cells(Rows.Count, "C").End(xlUp).Row
        
        OpeningPrice = Range("C2").Value
        ClosingPrice = Range("F2").Value
        Volume = Range("G2").Value
        
        Range("C2").Value = OpeningPrice
        Range("F2").Value = ClosingPrice
        Range("G2").Value = Volume
        
        YearlyChange = Range("F2").Value - Range("C2").Value
        PercentChange = Range("C2").Value / Range("F2").Value / 100
        TotalStockVolume = Range("G2").Value * Range("F2").Value
         
        Range("J2").Value = YearlyChange
        Range("K2").Value = PercentChange
        Range("L2").Value = TotalStockVolume

End Sub

