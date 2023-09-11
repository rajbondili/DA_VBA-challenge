Sub stockout():

For Each ws In Worksheets

'Declare variables
Dim ticker As String
Dim total_Volume As Double
Dim final_stock As Double
Dim open_stock As Double
Dim close_stock As Double
Dim Summary_Table_Row As Integer

'Assign initial values to the variables
total_Volume = 0
final_stock = 0
Summary_Table_Row = 2

'Populate the summary table column headers
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

'Fetch the last row of the sheet
last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To last_row
    
        'when current ticker value is not equal to next ticker value
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        'capture current ticker value from the sheet cell value
        ticker = Cells(i, 1).Value
        'capture total volume value from the sheet cell value and sum with existing total volume
        total_Volume = total_Volume + Cells(i, 7).Value
        'capture close stock value
        close_stock = Cells(i, 6).Value
        
        'Populate summary table row
        ws.Range("I" & Summary_Table_Row).Value = ticker
        ws.Range("L" & Summary_Table_Row).Value = total_Volume
        ws.Range("J" & Summary_Table_Row).Value = close_stock - open_stock
        ws.Range("k" & Summary_Table_Row).Value = (close_stock - open_stock) / open_stock
        ws.Range("k" & Summary_Table_Row).NumberFormat = "0.00%"
        
        'increase summary table row count to move to next row
        Summary_Table_Row = Summary_Table_Row + 1
        'Zero the below variables, so that it will be calculated again when current ticker value is not equal to next ticker value
        total_Volume = 0
        final_stock = 0
        open_stock = 0
        close_stock = 0
        
        Else
        
        'Sum the volume for each ticker
        total_Volume = total_Volume + ws.Cells(i, 7).Value
        
        If final_stock = 0 Then
            'sum current and new final open stock
            final_stock = final_stock + ws.Cells(i, 3).Value
            'Populate current  open stock
            open_stock = ws.Cells(i, 3).Value
            
            Else
             'sum current and new final open stock if it is not zero
            final_stock = final_stock + ws.Cells(i, 3).Value
            
            End If
         'sum final open stock for each ticker
        final_stock = final_stock + ws.Cells(i, 3).Value
        
        
        End If
         ' when percentage of yearly change is greater than zero populate cell with  green Color
         If ws.Cells(i, 10).Value > 0 Then
        
            ws.Cells(i, 10).Interior.ColorIndex = 4
         ' when percentage of yearly change is greater than zzero populate cell with red Color
        ElseIf ws.Cells(i, 10).Value < 0 Then
        
            ws.Cells(i, 10).Interior.ColorIndex = 3
        
        End If
        
    Next i
    
    
''Display Ticker which has Greatest -> increase,decrease and total volume of the stock out results

'Declare variables
Dim Greatest_increase As Double
Dim Greatest_increase_tikr As String
Dim Greatest_decrease As Double
Dim Greatest_decrease_tikr As String
Dim Greatest_total_volume As Double
Dim Greatest_total_volume_tikr As String

'Initialize the values
Greatest_increase = 0
Greatest_decrease = 0
Greatest_total_volume = 0

'Create row & column headers for Greatest value
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % increase"
ws.Cells(3, 15).Value = "Greatest % decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"


        For i = 2 To last_row
            'Condition to check if the current percentage Change is higher  than next Percentage Change
            ' if yes then populate the value and store the value to check the next higher for each Ticker
            If ws.Cells(i, 11).Value > Greatest_increase Then
               Greatest_increase = ws.Cells(i, 11).Value
               Greatest_increase_tikr = ws.Cells(i, 9).Value
               ws.Cells(2, 16).Value = Greatest_increase_tikr
               ws.Cells(2, 17).Value = Greatest_increase
               ws.Cells(2, 17).NumberFormat = "0.00%"
               
            'Condition to check if the current percentage Change is lower  than next Percentage Change
            ' if yes then populate the value and store the value to check the next lower for each Ticker
            ElseIf ws.Cells(i, 11).Value < Greatest_decrease Then
                Greatest_decrease = ws.Cells(i, 11).Value
                Greatest_decrease_tikr = ws.Cells(i, 9).Value
                ws.Cells(3, 16).Value = Greatest_decrease_tikr
                ws.Cells(3, 17).Value = Greatest_decrease
                ws.Cells(3, 17).NumberFormat = "0.00%"
                
            'Condition to check if the current total volume is greater than next total volume
            ' if yes then populate the value and store the value to check the next higher for each Ticker
            ElseIf ws.Cells(i, 12).Value > Greatest_total_volume Then
                Greatest_total_volume = ws.Cells(i, 12).Value
                Greatest_total_volume_tikr = ws.Cells(i, 9).Value
                ws.Cells(4, 16).Value = Greatest_total_volume_tikr
                ws.Cells(4, 17).Value = Greatest_total_volume
            End If
            
        Next i
    
  Next ws
  
End Sub
