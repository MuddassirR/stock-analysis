# stock-analysis

## Overview of Project
Using VBA in Excel, financial Analysis was conducted on green energy stocks to determine which stock had the best return in 2017 and 2018. Macros were created to create different variable types, such as string for a ticker's symbol and long for a ticker's starting and ending prices in a given year. Cells were accessed primarily using  Cells() objects. Range() was only used to type string characters. The correct worksheets were activated using Worksheets(" ").Activate. On the All stocks Analysis sheet, ticker symbol, total daily volume, and return were typed using VBA in cells A3:C3. Total daily and yearly volume, and prices was auto calculated using for loops, if then statements, and arrays. Visual formatting was applied using VBA such as changing text colour, font size, and autofitting. The solution code was refactored to loop through all the data one time to collect the same information quicker. 

### Purpose
The purpose of the analysis overall was to determine the the return of each green energy stock in the dataset by year. To make the analysis quicker, the code was refactored to run quicker. 


### Results 
We see from the [2017 results (not refactored)](https://github.com/MuddassirR/stock-analysis/blob/main/2017_not_refactored.png) that ENPH had the best return while TERP had the worst. It took 0.6 seconds to run the analysis. Compared to the [refactored 2017 results])(https://github.com/MuddassirR/stock-analysis/blob/main/VBA_Challenge_2017.png) DQ had the highest return while TERP still had the lowest return. This time, the code ran in 0.07 seconds, almost 10x faster. 

Furthermore, we see from the  [2018results (not refactored)](https://github.com/MuddassirR/stock-analysis/blob/main/2018_not_refactored.png) that Run had the highest return while DQ had the lowest and this was found in 0.6 sseconds. From the [refactored 2018 results](https://github.com/MuddassirR/stock-analysis/blob/main/VBA_Challenge_2018.png) we see that the code actually ran this in 0.07 seconds, almost 10x faster as well. 


## Summary
We recommend that given the companies in the dataset, ENPH seems to be a company that will yield high returns and we recommend avoiding investing in DQ or TERP.

### Pros & Cons of Refactoring
As mentioned on [Stack Exchange](https://sqa.stackexchange.com/questions/10311/pros-and-cons-of-refactoring-code-during-testing-phase#:~:text=Pros%3A%20%2D%20Your%20code%20will%20be,your%20tests%20won't%20catch.), by refactoring, it makes the code better organized and can improve the time it takes to run it. A con can be that it can potentially introduce bugs that your tests won't catch.

### Application of Pros & Cons to this Challenge
In this challenge, we saw that the time it took to run analysis on both years was shortened by almost 10 times. The immediate benefit of refactoring is that, especially when working with large datasets, it can improve efficiency and create time available for other activities. In terms of this activity, refactoring involved creating multiple different variables and arrays that had different data types such as string, long, and single. Working with multiple variables with different data types can get difficult, especially when trying to code using different variable types.

