# stock-analysis

## Overview of Project

### The purpose of this analysis is to help in deciding which of a preselected group of stocks to purchase. To do this I looked at how the set of stocks performed over two years, 2017 and 2018. For this analysis I will be looking at the rate of return for each year. Based on this data and how it changed over time I can make some guesses as to what might be the best stock to purchase.

## Results

### Comparing the Data
  Looking at the return data for 2017, we can see that almost all of the stocks in the group showed growth during the year. Only one stock, TERP, showed a loss. Of the stocks that had gains in 2017, 4 of them (DQ, ENPH, FSLR, SEDG) showed a improvment of over 100%, making those 4 look like good options.
  
![2017_Data.png](Resources/2017_Data.PNG)

  The data for 2018 shows a much different picture then the data for 2017. All but two stocks, ENPH and RUN, showed a loss. The two stocks that showed growth both had more then 80% return. This makes them the best bet for purchase. When choosing between the two it should be noted that while both stocks showed positive returns each year, only RUN actually improved more in 2018 then 2017. 
  
![2018_Data.png](Resources/2018_Data.PNG)

### Comparing the Code
  To create the summarized view of data that supproted the above analysis we used VBA code. Two versions of that VBA code were tested and compared to decide which would run better in the final worksheet. The first version averaged around 1.69 seconds for each year
  
![VBA_Challenge_2017_OrigCode.png](Resources/VBA_Challenge_2017_OrigCode.png)![VBA_Challenge_2018_OrigCode.png](Resources/VBA_Challenge_2018_OrigCode.png)

  The second version of the code ran much quicker. This refactored version averaged around 0.2 seconds. That huge difference shows that the refactored code is far more efficiant the the original version  
  
![VBA_Challenge_2017.png](Resources/VBA_Challenge_2017.png)![VBA_Challenge_2018.png](Resources/VBA_Challenge_2018.png)


## Summary

- Adavantages/Disadvantage of refactoring code

- How do these pros and cons apply to refactoring the original VBA script?
