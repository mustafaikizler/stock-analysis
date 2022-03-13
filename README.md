# STOCK ANALYSIS
## Purpose and Background of the Stock Analysis

### Main Purpose and Background

We had a dataset of multiple stocks with the year of the date, closing prices, and volumes. Our ultimate purpose was the find out the stocks that which has the hight volume and highest differences with the closing and opening prices. We wanted to find the most profitable tickers. We ran our analysis for 2017 and 2018

### VBA Analysis Background 

We used our VBA knowledge in this analysis. We had two different years and 12 different tickers to find returns and volumes according to ticker names. 

- Return represents subtraction of ending price, starting prices, and its percentage value
- Volume represents total buy/sell traffic of a certain ticker

![All Stock Analysis](https://user-images.githubusercontent.com/98247252/158040367-d09af356-1443-4297-a450-2d1a8e4effd4.png)

Our Secondary purpose was the reduce the analysis duration of the stock analysis by refactoring our main code. We needed to refactor the code because if we want to run our code worldwide and want to include thousands of stocks, the duration of the run time could be so long and we would face interruptions. Our main tool was the arrays. Arrays allowed us to run the once and gather all the data we would like to have. 

## Results

### Refactored Code ( '## xyx ## comments are added by writer.)

```
.
.
.

    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    'Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '## We created tickerIndex in order to call one of the twelwe ticker. Also Started it from zero as arrays starts from zero ##
    
    '1a) Create a ticker Index
    
    tickerIndex = 0
    '##We kept the datas in the arrays that we can call whenever want in the code. ##
    
    '1b) Create three output arrays
    Dim tickerVolume(12) As Long
    Dim tickerStartingPrices(12) As Long
    Dim tickerEndingPrices(12) As Long
     
    '##Outer for loop is created in a purpose of calling all the single tickers.##
     
    ''2a) Create a for loop to initialize the tickerVolumes to zero
    For i = 0 To 11
        tickerVolume(i) = 0
              
    Next i
     '##Inner loop created in a purpose of calling every single ticker and their data which belongs to prices, dates, volumes etc...## 
     
     ''2b) Loop over all the rows in the spreadsheet.
    
        For j = 2 To RowCount
        
        '##we sum all the volumes which belongs to a certain tickers.##
        
            '3a) Increase volume for current ticker
            If Cells(j, 1).Value = tickerIndex Then
                tickerVolume(tickerIndex) = tickerVolume(tickerIndex) + Cells(j, 8).Value
            End If
                    
             '##We Determined first and last price or the tickers. this information is according  the fiest day of the year and last days of the year.##    
             
            '3b) Check if the current row is the first row with the selected tickerIndex.
            'If  Then
             If Cells(j - 1, 1).Value <> tickers(tickerIndex) Then
                tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
             End If
                
            'End If
            
            '3c) check if the current row is the last row with the selected ticker
             'If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
            'If  Then
            If Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
                
                
                '## if the current row it the last row of the current ticker, we want tot switch to next tickers by using the tickerIndex ##
                
                '3d Increase the tickerIndex.
                               
                tickerIndex = tickerIndex + 1
                
                
            'End If
            End If
        
        Next j
           
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    
    '## Finally we wrote all the information that we gathered to the table (All Stock Analysis Sheet)  ##
    For i = 0 To 11
    Sheets("All Stock Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolume(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    Next i
    
    'Formatting
    .
    .
    .
```


  ***Duration of the process before refactoring the main code***


![VBA Challenge-NonRefactored-2017](https://user-images.githubusercontent.com/98247252/158041420-e263773a-16e3-489c-a990-cc33c2bf4711.png)
![VBA Challenge-NonRefactored-2018](https://user-images.githubusercontent.com/98247252/158041421-cede7c4b-f871-4383-ac17-19b9eede5997.png)

***Duration of the process after refactoring the main code***


![VBA_Challenge_2018](https://user-images.githubusercontent.com/98247252/158041439-3f36f4d3-860f-40b7-91de-e5bd3bb90327.png)
![VBA_Challenge_2017](https://user-images.githubusercontent.com/98247252/158041440-7ad46d2e-ecfb-447c-b232-31f1ec65f270.png)

- Duration of the process is less after running the refactored code.

## Summary

### Advantages and disadvantages of refactoring code in general 
- Running a refactored code (depends on the process) occupied less memory and ran faster in our VBA. If we can use this kind of solution in our codes we can save time, money, and labor.
- Refactoring a code can be confusing, especially if we are starting a project with the main code, we need to gather enough information about the way of coding.

### Advantages and disadvantages of the original and refactored VBA script
- Refactored script provided much more efficiency for instance by assigning "tickerIndex" we would easily reach to ticker which is located in the index. 


