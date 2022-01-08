# Overview of Project
In this project stock market data was analyzed. The given dataset included the starting 
and ending prices as well as the trading volume of twelve stocks for two years.
Analysis was performed,using VBA for Excel, to find the performance of each stock  
in a given year. In the first version of the code, an array and nested for-loops
were used to come up with the return for each stock.
The nested for loops find had to run through the whole dataset twice for each stock to find 
the starting and ending prices in order to calculate the return. 
The first version of the code was refactored in order to make it more efficient when running 
it on the entire stock market data over the last years.  


#Results
In the refactored code a tickerIndex variable was introduced to iterate through the tickers
array, the tickerVolumes array, the ticker startingPrice array, and the tickerEndingPrice array. 
Looping through the data once resulted in all the four outputs for one stock. 

A timer was scripted in both versions of the code so we can clearly see that the run times
were significantly reduced for both years. 

  ![Module2RunTime2017]("C:\Users\nagyf\Downloads\Module2VBA\Module2RunTime2017.png")
  ![refactoredRunTime2017]("C:\Users\nagyf\Downloads\Module2VBA\refactoredRunTime2017.png")
  ![Module2RunTime2018]("C:\Users\nagyf\Downloads\Module2VBA\Module2RunTime2018.png")
  ![refactoredRunTime2018]("C:\Users\nagyf\Downloads\Module2VBA\refactoredRunTime2018.png")

#Summary

## Advantages and disadvantages of refactoring code in general

Generally, code is refactored to reduce run time so the efficiency of 
analyzing large amount of data is increased. Refactoring code may require 
it's own economical analysis to justify time spent on changing and debugging
parts of an existing code as the process could be very time consuming.


## Detail on the advantages and disadvantages of the original and refactored VBA scrip
In this project, the refactored code resulted in a greatly decreased run time.
Removing the nested loop made reading the script much faster for the human eye
as well. The intent is use the code for the entire stock market for over the last
few years so for that large of a data set the faster run time will be much more
apparent than for the sample set used for the purpose of this project.
The disadvantages of refactoring the original code  
