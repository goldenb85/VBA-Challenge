Sub Stocks():

    Dim ws As Worksheet
    'Loop through all sheets
    For Each ws In Worksheets
        Dim i, j As Double
        
        'define an initial variable for holding the ticker
        Dim ticker As String
        'define an initial variable for holding the total stock volume per stock ticker
        Dim Volume_total As Double
        Volume_total = 0
        'Keep track of the location for each stock ticker in the summary table
        Dim summary_table_row As Double
        summary_table_row = 2
        
        'Determine the last row
        Dim lastrow As Double
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'Name I1,J1,K1,L1
        ws.Cells(1, 9) = "Ticker"
        ws.Cells(1, 10) = "Yearly Change"
        ws.Cells(1, 11) = "Percent Change"
        ws.Cells(1, 12) = "Total Stock Volume"
        'Name P1,Q1,O2,O3,O4
        ws.Range("P1") = "Ticker"
        ws.Range("Q1") = "Value"
        ws.Range("O2") = "Greatest % Increase"
        ws.Range("O3") = "Greatest % Decrease"
        ws.Range("O4") = "Greatest Total Volume"
        'define open_price and end_price data type
        Dim open_price, close_price As Long
              
        
        'loop through all stock tickers
            For i = 2 To lastrow
                'check if it is the first row of first ticker
                If i = 2 Then
                'set the value of open price for first ticker at a given year
                open_price = ws.Cells(i, 3).Value
                'put this value into range "ZY2"
                ws.Range("ZY" & summary_table_row).Value = open_price
                'check if it is the first row of the second ticker to other tick(except the first ticker)
                ElseIf i <> 2 Then
                
                
                        'check if we are still within the same stock ticker, if it is not... then
                        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                        'set the ticker of stock
                        ticker = ws.Cells(i, 1).Value
                        'add to the total stock volume per ticker
                        Volume_total = Volume_total + ws.Cells(i, 7).Value
                        'Print the ticker in the summary table
                        ws.Range("I" & summary_table_row).Value = ticker
                        'Print the total stock volume to the summary table
                        ws.Range("L" & summary_table_row).Value = Volume_total
                        'set the value of open price for the rest tickers at a given year
                        open_price = ws.Cells(i + 1, 3).Value
                        'put this value into range "ZY"coumn
                        ws.Range("ZY" & (summary_table_row + 1)).Value = open_price
                        'set the value of close price for all tickers at a given year
                        close_price = ws.Cells(i, 6).Value
                        'put this value into range "ZZ"coumn
                        ws.Range("ZZ" & summary_table_row).Value = ws.Cells(i, 6).Value
                        'calculate the yearly change
                        ws.Range("J" & summary_table_row).Value = (ws.Range("ZZ" & summary_table_row).Value - ws.Range("ZY" & summary_table_row).Value)
                        'formating the yearly change to 10 digits
                        ws.Range("J" & summary_table_row).NumberFormat = "##.00000000"
                        'calculate the percent change,format the cells from number to percentage
                        If ws.Range("ZY" & summary_table_row).Value <> 0 Then
                            ws.Range("K" & summary_table_row).Value = Format((ws.Range("ZZ" & summary_table_row).Value - ws.Range("ZY" & summary_table_row).Value) / ws.Range("ZY" & summary_table_row).Value, "Percent")
                            
                        End If
                            'check if year change is greater than 0 then fill color is green
                            If ws.Range("J" & summary_table_row).Value > 0 Then
                                 ws.Range("J" & summary_table_row).Interior.ColorIndex = 4
                                 'check if year change is less than 0 then fill color is red
                                 ElseIf ws.Range("J" & summary_table_row).Value < 0 Then
                                 ws.Range("J" & summary_table_row).Interior.ColorIndex = 3
                            End If
                                 
                        'Add one to the summary table row
                        summary_table_row = summary_table_row + 1
                        'Reset the total stock volume
                        Volume_total = 0
                                               
                        
                        'if the cell immediately following a row is the same ticker
                        
                        
                        Else
                        
                        'add to the total stock volume
                        Volume_total = Volume_total + ws.Cells(i, 7).Value
                        End If
            End If

        Next i
        'Determine the last row of column k
        Dim lastrow_k As Double
        
        
        lastrow_k = ws.Range("K" & Rows.Count).End(xlUp).Row
        'loop through all stock tickers in column I,J,K,L
        For j = 2 To lastrow_k
        If ws.Cells(j, 11).Value = WorksheetFunction.Max(ws.Range("K2:K" & lastrow_k)) Then
            'get the value for greatest% increase and format to percentage
            ws.Cells(2, 17) = Format(ws.Cells(j, 11), "percent")
            'get the ticker for greatest% increase
            ws.Cells(2, 16) = ws.Cells(j, 9)
        ElseIf ws.Cells(j, 11).Value = WorksheetFunction.Min(ws.Range("K2:K" & lastrow_k)) Then
            'get the value for greatest% decrease and format to percentage
            ws.Cells(3, 17) = Format(ws.Cells(j, 11), "percent")
            'get the ticker for greatest% decrease
            ws.Cells(3, 16) = ws.Cells(j, 9)
        ElseIf ws.Cells(j, 12).Value = WorksheetFunction.Max(ws.Range("L2:L" & lastrow_k)) Then
            'get the value for greatest total volume
            ws.Cells(4, 17) = ws.Cells(j, 12)
            'get the ticker for greatest total volume
            ws.Cells(4, 16) = ws.Cells(j, 9)
        End If
        Next j
        'delete the columns store open_price and close_price
        ws.Range("ZY2:ZZ" & lastrow_k).Delete
        
    Next ws
    
End Sub




