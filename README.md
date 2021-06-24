# VBA-challenge
Please find Sam's HW in this repository. 

The first three documents are screenshots of the summary tables generated after VBS code was applied to Multiple_year_stock data. The documents are in chronological order starting with the oldest: 2014, 2015, 2016. The image mirrors the screenshot presented in the assignment instructions. 

The last document is the macros-enabled Excel sheet with should include the VBS code. I have included the raw code below for reference. If there is any concern regarding my assignment, please do not hesitate to contact me! 

Sub StockOutput()
 
'Set to run macros in all worksheets
 
'Apply to all worksheets
For Each ws In Worksheets

'Define location for output
ws.Range("I1") = "Ticker"
ws.Range("J1") = "Annual Change"
ws.Range("K1") = "Percentage Change"
ws.Range("L1") = "Total Stock Volume"

'Autofit the cells for the summary table
ws.Range("I1:L1").Columns.AutoFit
 
'Set variable for Ticker as string
Dim Ticker As String

'Set annual change
Dim Annual_Change As Double

'Set percentage change
Dim Percentage_Change As Double

'Set total stock volume
Dim Total_Stock_Volume As Double
 
'Set counter for opening price
Open_Price = ws.Cells(2, 3).Value

'Set counter for total stock volume
Total_Stock_Volume = 0

'Set summary table variable
Summary_Table = 2
 
'Set variable for column to mark change in ticker
Column = 1
 
'Set RowCounter to mark end of row on sheet
    Dim RowCount As Long
        RowCount = ws.Cells(Rows.Count, 1).End(xlUp).Row
 
'Loop through row
For i = 2 To RowCount
 
'Note change in column 1 with conditional
If ws.Cells(i + 1, Column).Value <> ws.Cells(i, Column).Value Then

    'Print the Ticker in summary table
    Ticker = ws.Cells(i, Column).Value
    
    'Place Ticker value
    ws.Range("I" & Summary_Table).Value = Ticker
    
    'Calculate annual change
    Annual_Change = ws.Cells(i, 6).Value - Open_Price
    
    'Place annual change in summary table
    ws.Range("J" & Summary_Table).Value = Annual_Change
    
        'Account for zero in the denominator in calculating percentage change
        If Open_Price = 0 Then
            ws.Range("K" & Summary_Table).Value = Null
            
            'Otherwise calculate it normally
            Else
            Percentage_Change = (ws.Cells(i, 6).Value - Open_Price) / Open_Price * 100
            
            End If
    
    'Place percentage value
    ws.Range("K" & Summary_Table).Value = Round(Percentage_Change, 2) & "%"
    
            
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


