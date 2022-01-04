# Green Stocks Analysis

## Project Overview

The purpose of this project is to create an Excel spreadsheet to analyze stock market results for green stocks across different tickers and years. The project utilizes VBA scripts to quickly execute the desired results based on user input. Results of each query will then indicate which stocks performed best in the user-indicated year, and those results can be used to inform decisions about stock purchases. To help make the project as efficient as possible, the original code was refactored to reduce script runtime. This analysis discusses the results of the stock analysis as well as compares the original and refactored code to show the efficiency gained. 

## Results

### Green Stocks Results

Overall, the green stocks analysis showed that 2017 was an outstanding year for the stocks analyzed. Nearly all stocks showed a return, and around 36% of the stocks analyzed showed over a 100% return. The results for 2017 are shown below:

![https://raw.githubusercontent.com/CarlS2rt/stock-analysis/main/images/VBA_Challenge_2017.png](https://raw.githubusercontent.com/CarlS2rt/stock-analysis/main/images/VBA_Challenge_2017.png)



Green stocks did not fair as well in 2018, however, as only two stocks posted a positive return. Of the remaining stocks, five of the stocks posted a negative return greater than -20%. The results for 2018 are shown below:

![https://raw.githubusercontent.com/CarlS2rt/stock-analysis/main/images/VBA_Challenge_2018.png](https://raw.githubusercontent.com/CarlS2rt/stock-analysis/main/images/VBA_Challenge_2018.png)



When looking at the year-over-year trends, the only high performing stock of those analyzed is ENPH, posting a return greater than 80% both years. 

### The Code and Refactoring

The original code for the project executed two nested For loops to cycle through all stock ticker results. The results of the code were accurate and achieved the desired result; however, the code was inefficient in its use of nest For loops. By coding the nested For loops, the code had to cycle through the data more times than necessary. The original code is below:

```visual basic
For i = 0 To 11
    ticker = tickers(i)
    totalVolume = 0
    
	Worksheets(yearValue).Activate
         For j = 2 To RowCount
           '5a) Get total volume for current ticker
          	If Cells(j, 1).Value = ticker Then
			totalVolume = totalVolume + Cells(j, 8).Value
       		End If
       
       '5b) get starting price for current ticker
       		If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
           	startingPrice = Cells(j, 6).Value
       		End If

       '5c) get ending price for current ticker
       		If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
           	endingPrice = Cells(j, 6).Value
       		End If
       
    	Next j
    
   '6) Output data for current ticker
   Worksheets("All Stocks Analysis").Activate
   Cells(4 + i, 1).Value = ticker
   Cells(4 + i, 2).Value = totalVolume
   Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
   
Next i
```

Due to the unnecessary looping, the execution time was not as fast as it could be and the code was refactored to improve that performance. To reduce the extra time created by unnecessary looping, the nested For loops were refactored into two separate For loops while attaining the same results. This was achieved by defining new variables to hold the output data during the loops. These new variables were tickerIndex to hold the ticket output, and then the arrays tickerVolumes, tickerStartingPrice, and tickerEndingPrice to get the desired ticker results. The refactored code is below:

```vbscript
'1a) Create a ticker Index
Dim tickerIndex As Single
    tickerIndex = 0
    
'1b) Create three output arrays
Dim tickerVolumes(12) As Long
Dim tickerStartingPrices(12) As Single, tickerEndingPrices(12) As Single

'2a) Create a for loop to initialize the tickerVolumes to zero.
Worksheets(yearValue).Activate

For i = 0 To 11
	tickerVolumes(i) = 0
Next i

'2b) Loop over all the rows in the spreadsheet.
For i = 2 To rowcount

    '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
          
    '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
                
    '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, set the tickerEndingPrices.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1) <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
	'3d Increase the tickerIndex.
        tickerIndex = tickerIndex + 1
        End If

Next i

'4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
For i = 0 To 11
    
    Worksheets("All Stocks Analysis").Activate
    tickerIndex = i
    Cells(i + 4, 1).Value = tickers(tickerIndex)
    Cells(i + 4, 2).Value = tickerVolumes(tickerIndex)
    Cells(i + 4, 3).Value = tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex) - 1

Next i 
```

### Execution Time by Code and yearValue

The code refactoring was successful in improving the execution time of the code. The original execution times are below:

![https://raw.githubusercontent.com/CarlS2rt/stock-analysis/main/images/VBA_Challenge_2017_Timer_Original.png](https://raw.githubusercontent.com/CarlS2rt/stock-analysis/main/images/VBA_Challenge_2017_Timer_Original.png)

![https://raw.githubusercontent.com/CarlS2rt/stock-analysis/main/images/VBA_Challenge_2018_Timer_Original.png](https://raw.githubusercontent.com/CarlS2rt/stock-analysis/main/images/VBA_Challenge_2018_Timer_Original.png)

The execution times for the refactored code are below:

![https://raw.githubusercontent.com/CarlS2rt/stock-analysis/main/images/VBA_Challenge_2017_Timer.png](https://raw.githubusercontent.com/CarlS2rt/stock-analysis/main/images/VBA_Challenge_2017_Timer.png)

![https://raw.githubusercontent.com/CarlS2rt/stock-analysis/main/images/VBA_Challenge_2018_Timer.png](https://raw.githubusercontent.com/CarlS2rt/stock-analysis/main/images/VBA_Challenge_2018_Timer.png)

The end result was that the refactored code ran nearly ten times faster than the original code, which is a substantial gain in efficiency.

## Summary

Code refactoring is an essential tool for coders. The main advantages of refactoring are that you can create a more maintainable version of your code and you can remove redundancies or inefficiencies from the code to make it run faster or in a way that allows it to be reused more readily. A clear disadvantage of refactoring is that it requires a strong understanding of the code syntax on which to improve.

As it relates to VBA, code refactoring is similarly advantageous. As displayed through the example above, refactoring is able to clean code to make it more efficient. The refactored code created above is now faster than the original, and it is also more easily expanded or added into a larger set of code. The disadvantages remain those of knowledge and experience. Without enough background knowledge and resources, the time involved in refactoring may be too great and attempts to improve the code could potentially break it. 
