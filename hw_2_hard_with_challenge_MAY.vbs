Option Explicit

Sub TickerEasy()
    
    'define iterators
    Dim i As Integer 'iterator for the number of unique stocks
    Dim j As Long ' iterator for number of unique data points
    Dim k As Integer 'iterator for number of worksheets in the workbook
    
    'define variables for dates and prices
    Dim opendate As Double 'first date of year for stock trading
    Dim closedate As Double 'final date of year for stock trading
    Dim openprice As Double 'opening price stock
    Dim closeprice As Double 'closing price of stock
    Dim yearlychange As Double 'change in price of stock
    Dim percentchange As Double 'percentage change in price of stock
    Dim numtickers As Long 'number of unique stocks
    Dim numdata As Long 'number of data points
    Dim wscount As Integer 'number of worksheets in the workbook
    Dim volume As Double 'long is not large enough to prevent overflow error on volume summation
    Dim greatestgain As Double 'placeholder for greatest percent increase
    Dim greatestloss As Double ' placeholder for greatest percent decrease
    Dim greatestvolume As Double ' placeholder for greatest volume
    Dim pricerng As Range 'range to calculate min and max price changes
    Dim volumerng As Range 'range to calculate max volume
    
    'count number for worksheets in the workbook and assign as wscount
    wscount = ActiveWorkbook.Worksheets.Count
    
    'iterate same subroutine for every sheet in workbook
    For k = 1 To wscount
    Worksheets(k).Activate
         
        
    'set initial stock volume to zero
    volume = 0
    
    'Determine number of rows of data on each worksheet
    numdata = Cells(Rows.Count, 1).End(xlUp).Row
   
    'identify unique stock tickers and copy to column I
    ActiveSheet.Range("A:A").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=ActiveSheet.Range("I1"), Unique:=True
    
    'Count number of unique stock tickers from pasted values now in Column I
    numtickers = Cells(Rows.Count, 9).End(xlUp).Row
    
    'Set column headers for final data summary
    Range("I1") = "Ticker"
    Range("J1") = "Yearly Change"
    Range("K1") = "Percent Change"
    Range("L1") = "Total Stock Volume"
    Range("O2") = "Greatest % Increase"
    Range("O3") = "Greatest % Decrease"
    Range("O4") = "Greatest Volume"
    Range("P1") = "Ticker"
    Range("Q1") = "Value"
    
    'iterate through number of all stock tickers (COL I)
    For i = 2 To numtickers
        
        'iterate through all data points
        For j = 2 To numdata
            
            'logic test to sum individual stocks, essentially this is SUMIF function
            If Cells(j, 1).Value = Cells(i, 9).Value Then
            
            'calculate cumulative volume
            volume = Cells(j, 7).Value + volume
                  
            'report final volume for each ticker in column 12
            Cells(i, 12).Value = volume
            
            'Identify open date as first date value for individual ticker range and close date as last date value for individual range
                        
                If (Cells(j, 1).Value <> Cells(j - 1, 1).Value) Then
                opendate = Cells(j, 2).Value
                
                ElseIf (Cells(j, 1).Value <> Cells(j + 1, 1).Value) Then
                closedate = Cells(j, 2).Value
                
                End If
                               
                    'Assign values for close price and open price from rows that contain open dates and close dates
                    If Cells(j, 2) = closedate Then
                    closeprice = Cells(j, 6).Value
                                     
                    ElseIf Cells(j, 2) = opendate Then
                    openprice = Cells(j, 3).Value
                
                    End If
                        
                        'Determine annual and percentage changes from open and close prices for the year
                        yearlychange = closeprice - openprice
                        
                            'avoids divide by zero error for opening stock price of zero
                            If openprice = 0 Then
                            percentchange = 0
                            Else: percentchange = yearlychange / openprice
                            End If
                        
                        'Forumat yearly change to two decimal places
                        Cells(i, 10) = Round(yearlychange, 2)
                        
                        'format percent change cell as percentage
                        Cells(i, 11) = Format(percentchange, "Percent")
                
                            'Conditional formatting for annual positive change set interior cell color to green
                            If yearlychange > 0 Then
                                Cells(i, 10).Interior.ColorIndex = 4
                            
                            'Conditional formatting for annual negative change set interior cell color to red
                            ElseIf yearlychange < 0 Then
                                Cells(i, 10).Interior.ColorIndex = 3
                            
                            End If
                     
            End If
                               
        Next j
        
        'reset volume to zero between each stock SUMIF
        volume = 0
    
    Next i
            
    'Identify price and volume ranges using number of total tickers extracted from original data set
    Set pricerng = Range("K2", Cells(numtickers, 11))
    Set volumerng = Range("L2", Cells(numtickers, 12))
    
    'calculate max and min percent changes and max volume change and assign to new cells
    greatestgain = Application.WorksheetFunction.Max(pricerng)
    Range("Q2") = Format(greatestgain, "Percent")
    greatestloss = Application.WorksheetFunction.Min(pricerng)
    Range("Q3") = Format(greatestloss, "Percent")
    greatestvolume = Application.WorksheetFunction.Max(volumerng)
    Range("Q4") = greatestvolume
    
    
    'New for loop to identify tickers for greatest percent gain and loss and greatest volume - assigns to final table
    For i = 2 To numtickers
    
        If (Cells(i, 11).Value = greatestgain) Then
            Range("P2") = Cells(i, 9).Value
        ElseIf (Cells(i, 11).Value = greatestloss) Then
            Range("P3") = Cells(i, 9).Value
        End If
        
        If (Cells(i, 12).Value = greatestvolume) Then
            Range("P4") = Cells(i, 9).Value
        End If
    Next i
        
    Next k
    
End Sub



