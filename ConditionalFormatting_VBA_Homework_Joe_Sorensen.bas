Attribute VB_Name = "ConditionalFormatting"

' Conditional Formatting for Yearly Change

' Use IF, ElseIf, and Else to Apply Color Conditional Formatting to Cell Value in Column J
    
Sub ConditionalFormatting()
    
    'Dim LastRow As Long
    LastRow = Cells(Rows.Count, "J").End(xlUp).Row
    
    
        ' If > 0 then Green, If < 0 then Red
        
        ' Color the Positive Value Green
        If Cells(r + 1, 10).Value > 0 Then
        
        Cells(r + 1, 10).Interior.ColorIndex = 4
        
        ' Color a Value of 0 with No Fill
        ElseIf Cells(r + 1, 10).Value = 0 Then
    
        Cells(r + 1, 10).Interior.ColorIndex = 0
        
        ' Color the Negative Value Red
        Else
        
        Cells(r + 1, 10).Interior.ColorIndex = 3
    
        End If
    
    
End Sub
