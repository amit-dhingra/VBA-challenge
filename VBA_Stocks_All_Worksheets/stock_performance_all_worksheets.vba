Attribute VB_Name = "Module1"
Sub calculateStockPerformance()

'Declare all the variables
Dim ws As Worksheet
Dim i As Long
Dim j As Integer
Dim openingPrice As Double
Dim closingPrice As Double
Dim volume As Double
Dim numberOfStocks As Integer
Dim tickervalue As String
Dim tickerStartingPos As Long
Dim yearly_change As Double
Dim percent_change As Double
'Dim WS_Count As Integer
Dim great_per_incr As Double
Dim great_per_decr As Double
Dim great_total_vol As Double
Dim great_per_incr_ticker As String
Dim great_per_decr_ticker As String
Dim great_total_vol_ticker As String
Dim LastRow As Long

'Loop through each worksheet
For Each ws In Worksheets
    'Activate the current Worksheet
    ws.Activate
    
    'Initialize the Number of stocks to 1
    numberOfStocks = 1
    
    'Initialize the Stock volume to 0
    volume = 0
    
    'Initialize the Ticker Starting row number to 2
    tickerStartingPos = 2
    
    'Get the number of rows/tickers into LastRow variable
    LastRow = ws.Range("B" & Rows.Count).End(xlUp).Row
    
    'Assign the Cell labels - Ticker , Yearly change,Percent Change and Total Stock Volume
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

    'Challenge code
    'Assign the Cell labels - Ticker ,Value
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    
    'Challenge code
    'Assign the Cell labels - Greatest % Increase,Greatest % Decrease and Greatest Total Volume
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"

        'Loop through each row in the worksheet till the last but one row
        For i = 2 To LastRow
            
            'Add the new volume to the previous value of volume per ticker
            volume = volume + Cells(i, 7).Value
            
            'Check if the ticker value changed
            'If the ticker value changes reinitialize some of the holding variables
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                
                'Get the closing price of the particular ticker
                closingPrice = ws.Cells(i, 6).Value
                
                'Get the ticker value
                tickervalue = ws.Cells(i, 1).Value
                
                'Holding variable to store the number of stocks
                numberOfStocks = numberOfStocks + 1
                
                'Get the opening price of the ticker starting at row 2 and going to the next
                'row if the ticker value changes
                openingPrice = ws.Cells(tickerStartingPos, 3).Value
                
                'Get the ticker value when the ticker value changes
                ws.Cells(numberOfStocks, 9).Value = tickervalue
                
                'Calculate yearly change of the stock price of the particular ticker
                yearly_change = closingPrice - openingPrice
                
                
                ws.Cells(numberOfStocks, 10).Value = yearly_change
                
                'Assign a cell color of green or red based on the yearly change value
                    If yearly_change < 0 Then
                        ws.Cells(numberOfStocks, 10).Interior.ColorIndex = 3
                    Else
                        ws.Cells(numberOfStocks, 10).Interior.ColorIndex = 4
                    End If
                    
                'Calculate the percent change of yearly change of the ticket
                'Also check for 0 opening price value and do corresponding error handling
                If openingPrice = 0 Then
                    percent_change = 0
                Else
                    percent_change = 100 * ((closingPrice - openingPrice) / openingPrice)
                End If
                
                
                ws.Cells(numberOfStocks, 11).NumberFormat = "0.00\%"
                ws.Cells(numberOfStocks, 11).Value = percent_change
                ws.Cells(numberOfStocks, 12).Value = volume
                
                'Reinitialize the volume variable to 0
                volume = 0
                
                'Get the particular ticker's starting position where the ticker value changes
                tickerStartingPos = i + 1
                
            End If
        Next i
        
    'Get the last row for all the unique ticker in column I
    LastRow = ws.Range("I" & Rows.Count).End(xlUp).Row
    
    'Get the first row values for the ticker and Yearly change and volume in corresponding variables
    great_per_incr = ws.Cells(2, 11).Value
    great_per_decr = ws.Cells(2, 11).Value
    great_total_vol = ws.Cells(2, 12).Value
    great_per_incr_ticker = ws.Cells(2, 9).Value
    great_per_decr_ticker = ws.Cells(2, 9).Value
    great_total_vol_ticker = ws.Cells(2, 9).Value
    
    'Loop through all the unique tickers and calculate greatest percent increase,
    'Greatest percent decrease and greatest volume of the corresponding ticker
    For i = 3 To LastRow
        If ws.Cells(i, 11).Value > great_per_incr Then
            great_per_incr = ws.Cells(i, 11).Value
            great_per_incr_ticker = ws.Cells(i, 9).Value
        End If
        If ws.Cells(i, 11).Value < great_per_decr Then
            great_per_decr = ws.Cells(i, 11).Value
            great_per_decr_ticker = ws.Cells(i, 9).Value
        End If
        If ws.Cells(i, 12).Value > great_total_vol Then
            great_total_vol = ws.Cells(i, 12).Value
            great_total_vol_ticker = ws.Cells(i, 9).Value
        
        End If
        
    Next i
    
    'Change the number format of Cell(2,17) and (3,17) to percentage
    ws.Cells(2, 17).NumberFormat = "0.00\%"
    ws.Cells(3, 17).NumberFormat = "0.00\%"
    
    'Assign the values of greatest percent increase and decrease to the corresponding cells
    ws.Cells(2, 17).Value = great_per_incr
    ws.Cells(3, 17).Value = great_per_decr
    
    'Assign the values of greatest total volume to the corresponding cell
    ws.Cells(4, 17).Value = great_total_vol
    
    'Assign the value of the ticker that corresponds to each of the 3 values calculated above
    ws.Cells(2, 16).Value = great_per_incr_ticker
    ws.Cells(3, 16).Value = great_per_decr_ticker
    ws.Cells(4, 16).Value = great_total_vol_ticker
    
Next
    
    
End Sub

