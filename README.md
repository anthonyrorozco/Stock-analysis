## Green_Stock-analysis
Green Energy stock analysis with VBA Using financial data analysis

#### Overview of Project
## Purpose
Refactoring a Microsoft Excel VBA code to gather specific stock information in 2017 and 2018 and assess whether the stocks are worthwhile to invest in was the goal of this project.This procedure was earlier performed in a similar fashion; however, the purpose for this was to improve the original code's efficiency. Furthermore, compare the benefits and drawbacks of the refactored code to the original script.

#### Results
Two charts containing stock information on 12 distinct stocks are shown in the data. This data reveal the ticker, the annualized daily volume, and the profit it returned.
## Stock Analysis for 2017 and 2018.

### Orginial Code Run-time 2017

![VBA_Orginial_Runtime_2017](https://user-images.githubusercontent.com/105666905/175196411-9231e756-9c47-405b-8ee6-133737b3c49c.png)


### Original Code Run-Time 2018

![VBA_Orginial_Runtime_2018](https://user-images.githubusercontent.com/105666905/175196262-05729087-10e2-4535-a197-f59032a7e8db.png)

## Refractored VBA Script Run Time and Assessment
Using all of the information from Module 2, we are challenged to illustrate what we have learned and make it efficient using the code supplied. Not only does the code appear to run considerably faster, but there are now less lines to precisely comb through. We also have a lot of comments highlighting the refactor's success.

```

>>> '1a) Create a ticker Index
    
    tickerIndex = 0
    
    

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    
    '2a) Create a for loop to initialize the tickerVolumes to zero.
        
        For i = 0 To 11
            tickerVolumes(i) = 0
            
            Next i
        
        
        
    ''2b) Loop over all the rows in the spreadsheet.
    Worksheets(yearValue).Activate
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        

    
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
            End If
            
                
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

           End If
           
            
            '3d Increase the tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            
             tickerIndex = tickerIndex + 1

        End If
            
        'End If
    
    Next i
    
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        
        Worksheets("All Stocks Analysis").Activate
        tickerIndex = j
            Cells(i + 4, 1).Value = tickers(i)
            Cells(i + 4, 2).Value = tickerVolumes(i)
            Cells(i + 4, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
>>>    Next i

```

