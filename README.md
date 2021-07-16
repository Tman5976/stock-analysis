# Stock-Analysis

## Project Overview

### Purpose & Background

Before starting this challenge, we found the returns for 12 different stocks using VBA. We wanted to find the same information on the same stocks, but we needed to refactor our code from the module to account for a couple of changes.

We wanted to more efficiently run a subroutine that could account for a significantly higher number of stocks and make it run faster.
 
## Results
For i = 0 To 11
        tickerVolumes(i) = 0
    
Next i
    
      Worksheets(yearValue).Activate
        
      For j = 2 To RowCount
              tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
            
          If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j - 1, 1).Value <> tickers(tickerIndex) Then
               tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
            
          End If
            
          If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
            
                tickerIndex = tickerIndex + 1
            
           End If
           
       Next j
           
![VBA_Challenge_2017](https://user-images.githubusercontent.com/85756203/125705087-6ae6f6dc-e66a-4090-b2af-030e22767984.png)

Running the above code with the input year of 2017 resulted in the proper results for that year's stock times. In 2017, 11 of the 12 stocks had positive returns, including 4 stocks that had a return of over 100%.
The picture shows the assigned subroutine took about half a second to return the proper results.

![VBA_Challenge_2018](https://user-images.githubusercontent.com/85756203/125705099-3d034a42-1af9-4195-8265-a024535abeac.png)

Running the code with an input of 2018 also gave back the correct results. In 2018, only two stocks had a positve return. Those stocks being ENPH and RUN. ENPH showed a very strong return over both years, and RUN had a considerably stronger return in the second year.
Running the subroutine for 2018 took less time than running it in 2017. It took less than 1/10th of a second to run the subroutine for the year 2018.

## Summary

An advantage of refactoring code is that it can be easier than starting from scratch. If the code required for a project is similar to a previous project, it can be better to take the old code, and it, hopefully, will only need a few tweaks to match the new requirements.

Refactoring code can also require fewer resources. Taking old code and updating it will require less time and money than creating an entirely new routine.

Refactoring code can be an issue if the old code is not easily understood. Different people may look at the older code and may not come to the same conclusion. Old code that is not well described or not clean could cause difficulties for the people working on the updates.
For this module, refactoring the original code was a helpful thing to do. We were tasked with using different variables and a slightly different process to arrive at the same outcome.

A challenge refactoring our code in VBA was including the tickerIndex variable. I encountered some difficulty using other variables with the tickerIndex as the index.
