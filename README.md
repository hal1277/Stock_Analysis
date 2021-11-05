# VBA of Wall Street

An Analysis of stock performance across a given year

## Overview of Project

Steve wanted an efficient way to analyze stock performance over a given year.  While Steve was satisfied with the intial report and it's functionality he would like to be able to expand the data set to the entire stock market.  

### Purpose

The purpose of this project was to take the exitsing functional code and refactor it to reduce the time it takes to complete the analysis which would allow Steve to analyze far bigger data sets efficiently. 

## Results

The starting point of this project was writing a code to analyze the performance of one stock for one year for Steve.  After seeing the data Steve asked to see the same analysis but for 12 stocks and for 2 years.  Now Steve would like to have the ability to expand this analysis to the entire stock market.  While the original code does function it may not be efficient enough to analyze the entire stock market in a reasonable period of time.  

So, the original script has been refactored specifically with increasing speed of the analysis as the goal.  In the refactoring, arrays were used as a more efficient way to collect the data and an index was added to connect the data across the 4 arrays being used.  As well I added an additional variable called lineIndex to stop the next loop from reading the first row of data that was already analyszed in the prior loop to redcue processing time.  And, finally an Exit For condition was added so that once the endingPrice criteria was met the loop would be exited without the code running through every remaining row in the data, again saving some processing time.  

As a result the refactored code prdouces exactly the same data results but in 85% less processing time as can be seen in these images 

![2017 Original Script Speed](https://github.com/hal1277/Stock_Analysis/blob/ee14cc578676c7bb16d3eea886afd2ec7e77f5c3/VBA_Challenge_2017_Original_Code.png)
https://github.com/hal1277/Stock_Analysis/blob/ee14cc578676c7bb16d3eea886afd2ec7e77f5c3/VBA_Challenge_2017.png
https://github.com/hal1277/Stock_Analysis/blob/ee14cc578676c7bb16d3eea886afd2ec7e77f5c3/VBA_Challenge_2018_Original_Code.png
https://github.com/hal1277/Stock_Analysis/blob/ee14cc578676c7bb16d3eea886afd2ec7e77f5c3/VBA_Challenge_2018.png

## Summary

### Refactoring Code in General

#### Pros of Refactoring

Some advantages of refactoring include not having to start from scratch on a problem.  As well refactoring can lead to more organized code that is easier to read and understand.  It can also lead to code that runs faster.  

#### Cons of Refactoring

Refactoring is not always the best solution.  If a code is very large, very complex, and not very readable or well described it may be faster to start a new code from scratch.  Reworking a complex poorly organized code may take a lot of time and produce unreliable results.  

### Refactoring the VBA code for this challenge

In this challenge the original code was short, simple, and quite readable so it was a good candidate for refactoring.  A clear goal to be acheived was also present - the goal of speed to allow for the code to analyze larger data sets.  The original code performed reliably and so it was reasonable to use that code as a starting point.  The refactored version runs faster and is also reliable with predictable outcomes.  

The disadvantage of the original script was speed - it would take a long time if the entire stock market needed to be analyzed.  The refactored code runs much faster but it's more complex and in order to maximize speed an element of risk exists.  The code as written requires that the tickers are sorted by ticker name in the data set.  If that sort is not correct the output will be incorrect.  This could be corrected through further refactoring by adding a line of code to sort the tickers before entering the analysis loops.  
