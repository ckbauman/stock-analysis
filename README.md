# Stock Analysis with VBA

## Overview of Project

Our client Steve has been reviewing Stock Analysis.  Originally he was only looking at one stock and expanded to 12 stocks.  He has been reviewing Daily Volume and Returns over the years of 2017 and 2018.

Total Daily volume is calculated by adding up the total purchases of each stock over the year.  Return is calculated by determing the stock prices at the beginning of the year and the end of the year and looking for a percent increase or decrease.

Now that we have developed the information for 12 specific stocks, it was determined that Steve may want to run an analysis on all stocks and not limit himself to 12.  This will require refactoring our analysis code to run more efficiently.  Specifically this will be done by transforming the code to run only one time, vs. loooping through multiple times for each index.

## Results

### Stock Comparison between 2017 and 2018

Stock analysis for 2017 shows that all but 1 of the stocks showed a positive percent increase on it's Return over the year.  

![VBA_2017_Output](https://github.com/ckbauman/stock-analysis/blob/main/VBA_Challenge_2017_Output.png)

Specifically **DQ**, **ENPH** and **SEDG** had very large Returns.  You can see why Steve's parents decided they wanted to invest in **DQ** in 2018.

The only negative Return was for **TERP** at -7.2%.

Stock analysis for 2018 shows that most stocks showed a negative percent increase on it's Return over the year

![VBA Output 2018](https://github.com/ckbauman/stock-analysis/blob/main/VBA_Challenge_2018_Output.png)

**DQ** took a huge drop at -62.6% but **ENPH** and **RUN** were the only stocks that remained positive over the 2 years.  In general, 2018 stocks did not perform well


### Execution times between Original and Refactored script

We created 2 seperate alaysis macros.  They created the same output, but used different methods.

The first macro was:

    Sub yearValueAnalysis()

This macro ran and created output into the **"All Stocks Analysis"** tab using an input button to indicate which year to run.  When the data was created, a message box pops up to indicate how long it took to run the macro.

The time is determined using:

    Dim startTime
    Dim endTime

The following images show run time for 2017 and 2018 using our original macro.

![VBA Original Runtime 2017](https://github.com/ckbauman/stock-analysis/blob/main/VBA_Challenge_2017Original.png)
![VBA Original Runtime 2018](https://github.com/ckbauman/stock-analysis/blob/main/VBA_Challenge_2018Original.png)

As you can see, the run time for 2017 and 2018 are:

- .875 seconds (2017)
- .890625 seconds (2018)

The second refactored macro was:

    Sub AllStocksAnalysisRefactored()

This macro introduces a tickerIndex variable and only loops once.  This should speed up the run time.

- **Step 1a:**  Create a ticker Index - need to Initialize to zero at start. Create a variable named tickerIndex and then set it = to zero to initialize it

        Dim tickerIndex As String
        tickerIndex = 0

- **Step 1b:**  Create three output arrays - Same as tickers(12) array above to run 12 times

        Dim tickerVolumes(12) As Long
        Dim tickerStartingPrices(12) As Single
        Dim tickerEndingPrices(12) As Single

- **Step 2a:**  Create a for loop to initialize the tickerVolumes to zero. - As the array runs through all 12 Indexes as i - it will hold the output

        For i = 0 To 11
    
            tickerVolumes(i) = 0
    
        Next i

- **Step 2b:** Loop over all the rows in the spreadsheet. - from the input spreadsheet (already in starter code)

        For i = 2 To RowCount 

- **Step 3a:** Increase volume for current ticker - shown in hint - tickerVolume is previous plus current

        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value

- **Step 3b:** Check if the current row is the first row with the selected tickerIndex. If it is then get the current starting price

        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
        
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If

- **Step 3c:** check if the current row is the last row with the selected tickerIndex. If the next row's ticker doesn't match, increase the tickerIndex. If it matches then assign current closing price to ending price.

        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
        
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

     - **Step 3d:** Increase the tickerIndex.  Takes the current value for tickerIndex and increases by 1 to move on to next.

            tickerIndex = tickerIndex + 1
            End If


- **Step 4:** Loop through your arrays to output the Ticker, Total Daily Volume, and Return.  looping through 12 rows again for each array - the values come from the year worksheets prior and are put into the "All Stocks Analysis" worksheet in row 4 and columns 1, 2 and 3

        For i = 0 To 11
        
            Worksheets("All Stocks Analysis").Activate
            Cells(4 + i, 1).Value = tickers(i)
            Cells(4 + i, 2).Value = tickerVolumes(i)
            Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
        
        Next i

The following images show run time for 2017 and 2018 using our refactored macro.

![VBA Refactored Runtime 2017](https://github.com/ckbauman/stock-analysis/blob/main/VBA_Challenge_2017.png)
![VBA Refactored Runtime 2018](https://github.com/ckbauman/stock-analysis/blob/main/VBA_Challenge_2018.png)

As you can see, the run time for 2017 and 2018 are:

- .140625 seconds (2017)
- .1640625 seconds (2018)

Refactoring shows a significant increase in efficency for the run time.  It may not appear that different for such a small dataset, but it would be a major improvement for a larger dataset.

## Summary

### What are the advantages or disadvantages of refactoring code?

Advantages to refactoring include improving efficiency and run time of the code.  Requirements change and it is worth reviewing code to run more effectively.

Disadvantages are *"don't fix what is not broken"*.  It takes time to rework code and it might not be a priority just yet.

In this Module, the refactoring was shown to improve run time.

### How do these pros and cons apply to refactoring the original VBA script?
 
We refactored the original VBA script and it showed an improvement on run time.  The refactored code only had to run 1 time vs. running multiple times in a double loop.

See refactored code below:

![Refactored code 1](https://github.com/ckbauman/stock-analysis/blob/main/VBA_Code_1.png)

![Refactored code 2](https://github.com/ckbauman/stock-analysis/blob/main/VBA_Code_2.png)

![Refactored code 3](https://github.com/ckbauman/stock-analysis/blob/main/VBA_Code_3.png)


