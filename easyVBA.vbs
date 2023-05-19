Sub refresh()


    Dim ws As Worksheet
	
    For Each ws In Worksheets
	
        ' Delete columns 9-17 (columns I-R)
        Range(ws.Columns(9), ws.Columns(17)).Delete
		
    Next ws
	
	
End Sub
Sub easyVBA()


    ' Set variables
    Dim row As Long
    Dim rowCount As Long
    Dim tickerSummaryRow As Integer
    
    Dim total As Double
    
    Dim startPrice As Double
    Dim endPrice As Double
    
    Dim priceChange As Double
    Dim percentChange As Double
    
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String
    
    Dim greatestIncreaseValue As Double
    Dim greatestDecreaseValue As Double
    Dim greatestVolumeValue As Double
    

    ' Print header and value labels
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
	Cells(1, 15).Value = "Outstanding Statistics"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"

    ' Set initial values for individual ticker data
    tickerSummaryRow = 2
    total = 0
    startPrice = 0
    
    ' Set initial values for standout ticker data
    greatestIncreaseTicker = " "
    greatestDecreaseTicker = " "
    greatestVolumeTicker = " "
    
    greatestIncreaseValue = 0
    greatestDecreaseValue = 0
    greatestVolumeValue = 0
    

    ' get the row number of the last row with data
    rowCount = Cells(Rows.Count, "A").End(xlUp).row

    For row = 2 To rowCount
    
        ' We increment the total every time regardless of the boundary
        total = total + Cells(row, 7).Value
    
        ' Find first nonzero start price to avoid division by zero errors
        If startPrice = 0 Then
            startPrice = Cells(row, 3).Value
        End If

        ' If ticker changes then calculate and print statistics
        If Cells(row + 1, 1).Value <> Cells(row, 1).Value Then
		
			' Print the ticker in the first column of the summary row
			Cells(tickerSummaryRow, 9).Value = Cells(row, 1).Value

            ' Begin summary statistics, printing, and formatting
            If total = 0 Then
			
                ' Print the results
                Cells(tickerSummaryRow, 10).Value = 0
                Cells(tickerSummaryRow, 11).Value = "%" & 0
                Cells(tickerSummaryRow, 12).Value = 0

            Else

                ' Calculate yearly changes
                endPrice = Cells(row, 6).Value
                priceChange = endPrice - startPrice
                percentChange = priceChange / startPrice

                ' Print the summary statistics for the current ticker	
                Cells(tickerSummaryRow, 10).Value = priceChange
                Cells(tickerSummaryRow, 10).NumberFormat = "0.00"
				
                Cells(tickerSummaryRow, 11).Value = percentChange
                Cells(tickerSummaryRow, 11).NumberFormat = "0.00%"
				
                Cells(tickerSummaryRow, 12).Value = total
                Cells(tickerSummaryRow, 12).NumberFormat = "#,##0"


                ' Color positives green, negatives red, and zeros white
                If priceChange > 0 Then
                    Cells(tickerSummaryRow, 10).Interior.ColorIndex = 4
                ElseIf priceChange < 0 Then
                    Cells(tickerSummaryRow, 10).Interior.ColorIndex = 3
                Else
                    Cells(tickerSummaryRow, 10).Interior.ColorIndex = 0
                End If
                
                ' Check if this ticker has the greatest total volume so far
                If total > greatestVolumeValue Then
                    greatestVolumeValue = total
                    greatestVolumeTicker = Cells(row, 1).Value
                End If
                
                ' Or if it has the greatest % increase/decrease so far
                If percentChange > greatestIncreaseValue Then
                    greatestIncreaseValue = percentChange
                    greatestIncreaseTicker = Cells(row, 1).Value
                ElseIf percentChange < greatestDecreaseValue Then
                    greatestDecreaseValue = percentChange
                    greatestDecreaseTicker = Cells(row, 1).Value
                End If
				
			' End of summary caluculations, printing, and formatting
            End If

            ' Finally, reset variables for the next stock ticker
            tickerSummaryRow = tickerSummaryRow + 1
            total = 0
            startPrice = 0

		' End of the boundary conditional data processing
        End If

    Next row

    ' Print the standout tickers after looping through all of them
    Cells(2, 16) = greatestIncreaseTicker
    Cells(2, 17) = greatestIncreaseValue
    Cells(2, 17).NumberFormat = "0.00%"
    
    Cells(3, 16) = greatestDecreaseTicker
    Cells(3, 17) = greatestDecreaseValue
    Cells(3, 17).NumberFormat = "0.00%"
    
    Cells(4, 16) = greatestVolumeTicker
    Cells(4, 17) = greatestVolumeValue
    Cells(4, 17).NumberFormat = "#,##0"


End Sub

