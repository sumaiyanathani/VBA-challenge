Attribute VB_Name = "Module1"
Sub stock_loop()


'Defining variables
Dim yearly_change As Double
Dim ws As Worksheet


For Each ws In Worksheets
    'Finding the last row using code rather than manually and storing it in a variable
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'Making a variable for yearly change that can be reset
    yearly_change = 0

    'Making a variable for total stock volume that can be reset
    total_stock_volume = 0

    'Making a counter for the stock ticker
    ticker_counter = 2

    'Naming all column headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

    'Calculating difference looping through rows in the worksheet
    For j = 2 To 300
        If ws.Cells(j + 1, 1).Value <> ws.Cells(j, 1).Value Then
            diff = ws.Cells(j, 1).Row - 2
        End If
    Next j
    
    'Making a counter for opening price and closing price
    o = 2
    c = 2 + diff

    'Traversing through each row in column 1
    For i = 2 To lastrow

        'Storing stock ticker in a variable
        ticker_name = ws.Cells(i, 1).Value
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
          
          'Calculating yearly change
            yearly_change = ws.Cells(c, 6).Value - ws.Cells(o, 3).Value
            ws.Range("I" & ticker_counter).Value = ticker_name
            ws.Range("J" & ticker_counter).Value = Round(yearly_change, 2)

            'Calculating percent change
            percent_change = yearly_change / ws.Cells(o, 3).Value
            ws.Range("K" & ticker_counter).Value = percent_change
            ws.Range("K" & ticker_counter).NumberFormat = "0.00%"
            
            'Conditional formatting using red and green colours. Red is for values less than zero and green for those more than zero
            If yearly_change > 0 Then
                ws.Range("J" & ticker_counter).Interior.ColorIndex = 4
            Else
                ws.Range("J" & ticker_counter).Interior.ColorIndex = 3
            End If
            For k = o To c
                total_stock_volume = total_stock_volume + ws.Cells(k, 7).Value
            Next k
            ws.Range("L" & ticker_counter).Value = total_stock_volume
        
            'Incrementing the counters and resetting the values of certain variables
            ticker_counter = ticker_counter + 1
            yearly_change = 0
            o = o + diff + 1
            c = c + diff + 1
            total_stock_volume = 0
        
            Else
            yearly_change = ws.Cells(c, 6).Value - ws.Cells(o + 1, 3).Value
        
        End If
    
    Next i
    
    'This code uses the max and min worksheet functions to identify the values with the greatest increase, greatest decrease and greatest volume
    greatest_increase = WorksheetFunction.Max(ws.Range("K2", "K" & ticker_counter))
    greatest_decrease = WorksheetFunction.Min(ws.Range("K2", "K" & ticker_counter))
    greatest_volume = WorksheetFunction.Max(ws.Range("L2", "L" & ticker_counter))
    
    'Giving names to the rows and columns
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"

    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    'Populating each cell with the greatest increase, greatest decrease and greatest volume values
    ws.Range("Q2").Value = greatest_increase
    ws.Range("Q3").Value = greatest_decrease
    ws.Range("Q4").Value = greatest_volume
    
    'Making sure the formatting is in percent
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("Q3").NumberFormat = "0.00%"
    
    'This for loop loops over the summary table to indentify the ticker value associated with the greatest increase, decrease and volume values
    For x = 2 To ticker_counter
        If ws.Cells(x, 11).Value = greatest_increase Then
            ws.Range("P2").Value = ws.Cells(x, 9).Value
        ElseIf ws.Cells(x, 11).Value = greatest_decrease Then
            ws.Range("P3").Value = ws.Cells(x, 9).Value
            ws.Range("P3").Value = "0.00%"
        ElseIf ws.Cells(x, 12).Value = greatest_volume Then
             ws.Range("P4").Value = ws.Cells(x, 9).Value
       End If
    Next x
    
    'Makes sure that all the columns are showing the entire value
    ws.Columns.AutoFit

Next ws

End Sub

