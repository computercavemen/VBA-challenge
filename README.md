# VBA-challenge
Please find Sam's HW in this repository. 

The first three documents are screenshots of the summary tables generated after VBS code was applied to Multiple_year_stock data. The documents are in chronological order starting with the oldest: 2014, 2015, 2016. The image mirrors the screenshot presented in the assignment instructions. 

The last document is the macros-enabled Excel sheet with should include the VBS code. If there is any concern regarding my assignment, please do not hesitate to contact me! 
            
    'Place total stock value
    ws.Range("L" & Summary_Table).Value = Total_Stock_Volume
    
    'Grab the new open price
    Open_Price = ws.Cells(i + 1, 3).Value
    
            'Color code the cell with nested conditional
            If Annual_Change > 0 Then
            ws.Range("J" & Summary_Table).Interior.ColorIndex = 4
            
            ElseIf Annual_Change < 0 Then
            ws.Range("J" & Summary_Table).Interior.ColorIndex = 3
            
            End If
            
    'Increment summary table
    Summary_Table = Summary_Table + 1
            
Else
                
    'Calculate the total stock volume
    Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
    
                
        End If
Next i

Next ws

End Sub


