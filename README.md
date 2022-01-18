# stock-analysis
## Overview of the Project

This project amis to analyze different stocks for the years 2017 and 2018 with the help of VBA by using loops, conditionals, formatting, and arrays. There is an illustration of runtimes when only a subset of stocks were analyzed versus the full set.

## Results
The initial subset of data analyzed was in the green_stocks excel that has the macro AllStockAnalysis. Considering 2017, there was a tremendos improvements in the runtimes from 0.5009 seconds to 0.10400 seconds in the refactored code as shown below in the images. 


![alt text](https://github.com/riteshnimmagadda/stock-analysis/blob/main/green_stocks_2017.png "Green Stocks 2017 data set run time")

![alt text](https://github.com/riteshnimmagadda/stock-analysis/blob/main/VBA_Challenge_2017.png "Full Dataset analysis run time for 2017")

Similarly for 2018, green stock analysis returned 0.49694 seconds whereas the run time for the refactored code on the whole data set was 0.10900 seconds.


![alt text](https://github.com/riteshnimmagadda/stock-analysis/blob/main/green_stocks_2018.png "Green Stocks 2018 data set run time")


![alt text](https://github.com/riteshnimmagadda/stock-analysis/blob/main/VBA_Challenge_2018.png "Full Dataset analysis run time for 2018")



### Refactoring the code:
There are couple of things that helped in looping through the stock set. one such thing is the use of tickerindex, which helped in improved run times when searching for the tickers. tickerstartingprices and tickerendingprices have helped to eliminate the nessecity of using the nested loops, thus by saving the amount of time this takes to run.


## Summary
### Advantages of refacoring in general
Changing the structure inside the code to optimize it to higher performace without changing the resulting output is refactoring. it helps in code optimization, run times, output efficiencies and also detects bugs earlier than later. Refactoring could be done just at any time before moving into production. 

### Disadvantages:
Some of the disadvantages of refactoring is that sometimes it could be more expensive to put efforts in refacotring than the actual effectiveness of the refactored code. spending too much time in refacotring could delay the delivey of the product. there is also a problem of introducing additional bugs while refactoring if not tested throughly.

### Advantages of refactored VBA code:
Major advantage of the refactored stock analysis VBA code is that the run times are considerably shortened even when using a larger data set. The shortening of run times can give us a ripple effect depending on how large the data set is.

### Disadvantages of refactored VBA code:
Introduction of additional variables in the new code such as tickerindex, tickerstartingprices and ticketendingprices makes the code look longer. As a personal example, I myself did few typos while writing down the new variables. See it is advised to be more cautious while refacotring the code.
