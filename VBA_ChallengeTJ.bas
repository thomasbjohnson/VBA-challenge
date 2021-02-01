Attribute VB_Name = "Module1"
Sub stock_tracker()

    ' Create a script that will loop through all the stocks for one year and output the following information.
    ' The ticker symbol.
    ' Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
    ' The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
    ' The total stock volume of the stock.
    ' You should also have conditional formatting that will highlight positive change in green and negative change in red.

    ' loop through the worksheets
    Dim ws As Worksheet
    For Each ws In Worksheets

        ' declare the variables
        Dim open_price As Double, close_price As Double, yearly_change As Double
        Dim percent_change As Double, stk_vol As Double
        Dim LastRow As Long ' to find the last row
        Dim output_table As Integer
        
        
        ' set headers for summary table
        ws.Range("i1").Value = ("Ticker Symbol")
        ws.Range("J1").Value = ("Yearly Change")
        ws.Range("k1").Value = ("Percent Change")
        ws.Range("l1").Value = ("Total Stock Volume")
        
         ' set the initial start values of the variables
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        output_table = 2
        
        
        ' begin the for loop to go through the rows skipping the headers
        For i = 2 To LastRow
        
            ' set openingstock price value
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1) Then
                open_price = ws.Cells(i, 3).Value
            End If
        
        ' Determine when the stock ticker changes
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
        
                'Display that ticker in the column row
                 ws.Cells(output_table, 9).Value = ws.Cells(i, 1).Value
                
                'Add the Stock Value
                stk_vol = stk_vol + ws.Cells(i, 7)
                ws.Cells(output_table, 12).Value = stk_vol
                
                
                ' Set the closing stock price values
                close_price = ws.Cells(i, 6).Value
            
                 ' Calculate the yearly change
                yearly_change = close_price - open_price
                ws.Cells(output_table, 10).Value = yearly_change
             
               
                
            'Find the percent change
            If open_price = 0 And close_price = 0 Then
                percent_change = 0
                ws.Cells(output_table, 11).Value = percent_change
                ws.Cells(output_table, 11).NumberFormat = "General%"
            'Division by 0 error
            ElseIf open_price = 0 Then
               ws.Cells(output_table, 11).Value = ("New Stock")
            Else
                percent_change = yearly_change / open_price
                ws.Cells(output_table, 11).Value = percent_change
                ws.Cells(output_table, 11).NumberFormat = "General%"
                End If
        
        ' Conditional formatting.  Green for Positive change, Red for Negative change, Yellow for no change
                If yearly_change = 0 Then
                    ws.Cells(output_table, 10).Interior.ColorIndex = 6  'Yellow
                ElseIf yearly_change > 0 Then
                    ws.Cells(output_table, 10).Interior.ColorIndex = 4  'Green
                Else
                    ws.Cells(output_table, 10).Interior.ColorIndex = 3  'Red
                End If
        
        
        
        
        
            ' Go to the next ticker and reset output table
                    output_table = output_table + 1
                   
             
             
        End If
        
        
        
     Next i
     
    Next ws
    
    


 End Sub
 
 



