Attribute VB_Name = "StockData"
Sub StockTesting()
        ' ==============
        ' VBA Stock Assignment
        ' ==============
        
        
        ' I. VARIABLE DECLARATIONS
        ' -----------------------------------------
        Dim i As Long
        Dim j As Long
        Dim cell As Range
    
        Dim Ticker As String
        Dim OpeningPrice As Double
        Dim ClosingPrice As Double
        Dim Volume As Long
        Dim YearlyChange As Double
        Dim PercentChange As Double
        Dim TotalStockVolume As Long
        Dim LastRow As Long
    
        
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


        ' Use ''For Each'' Function to loop through worksheets


