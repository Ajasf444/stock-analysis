# Stock Analysis using VBA

## Overview of Project

### Purpose

The purpose of this analysis was to analyze the investment potential for twelve different green energy stocks for our client, Steve.
This workbook takes annual stock ticker information and produces a formatted output highlighting the yearly return for each ticker and whether the stock would have appreciated in value or not.
The code utilized in this project was optimized to work for more than just the original twelve stocks analyzed. 

## Results

![2017 performance](/Resources/VBA_Challenge_2017.png)
![2018 performance](/Resources/VBA_Challenge_2018.png)

The annual return on the twelve sample stock tickers provided showed a positive return for all stocks save TERP for the year of 2017.
TERP had a loss in value of 7.2%.
For the year of 2018 all stocks save ENPH and RUN.
ENPH and RUN had positive returns of 81.9% and 84.0%, respectively.
This information can easily be gleaned from the images provided above.

In terms of the code performance one can see that all the stock information for both analyses was under 0.15 seconds.
This information can be seen in the images above.
The original code would perform the analyses in approximately 0.7 seconds.
This represents a speedup in the performance of around a factor of 5, a rather significant improvement.

## Summary

1. What are the advantages or disadvantages of refactoring code?
   - Advantages
 - The code can run more efficiently by using either less memory, or less time.
       The program does not need to necessarily improve in either of these regards either.
       The code could simply be easier to understand which would make it easier to maintain.

 - Disadvantages
 - There is a cost in time to refactor existing code which may not be necessary or worthwhile in the grand scheme of the project.
     - Sometimes finding an extremely efficient algorithm for doing some task may ultimately make the code incredibly difficult to understand which reduces the maintainability which may not be desireable if utmost performance is not a requirement. 

2. How do these pros and cons apply to refactoring the original VBA script?
   - The refactoring of the code in this challenge improved the run time of the code by a factor of around 5.
     This makes the code much easier to have implemented for larger pools of stock tickers.
     The code was changed to not loop over each ticker and have an inner loop that looped over the entire list of data.
     The new code simply looped over all the data once, and had a tickerIndex variable that would increment based on the changing ticker symbol in said loop.
   - The refactored code did not introduce any new complicated features to the code and is very easy to read and maintain.
     In terms of downsides to this change, all the information for each ticker's starting price, ending price, and transaction volumes was stored in memory.
     This memory usage while small could be reduced if the information was printed to the spreadsheet during the tickerIndex increment.
     However, the time taken to print at this stage may end up taking more time than merely printing all computations at the end.
