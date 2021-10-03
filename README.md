# stock-analysis

## Overview of Project 

### Purpose
	This Project's purpose was to build a macro to analyze certain metrics (total daily volume and return) of stocks over the course of a year. 

## Results

### 2017 and 2018 differences 
	![2017 Stock Analysis](https://github.com/wellma41/stock-analysis/blob/main/VBA_Challenge_2018.PNG?raw=true)   [2018 Stock Analysis] (VBA_Challenge_2018.PNG)  
	
	Based on the return percentages of both years 2017 seems to have been a much better year for the stocks we kept track of.
	In 2017 11 out of the 12 stocks, TERP being the exception, had a positive return percentage and thus a profit if the stocks were held on to.
	2018, by contrast only saw two stocks (RUN and ENPH) stay positive, the rest lost value over the course of the year.
	The total daily volume saw an increase of around 140 million. This means that these stocks were traded more frequently in 2018.
	In conclusion, RUN and ENPH were the only stocks to not incur a loss at some point so they might be ideal for a long-term investment.


## Summary
	

### Pros and Cons of Refactoring Code
	Pros: Refactoring the code allows you to re-work through your code and write it in a way that is more efficient that faster.
	 In addition, code allows you to gain a more robust understanding of the data set and the variables at play so if the code breaks in the future when additions are made to the sheet you will be more equipped to deal with it.

	Cons: Refactoring the code can be time consuming, arduous process and possibly redundant. The old saying "If it a'int broke don't fix it" seems applicable here.

	Conclusions to refactoring: If you have the time, it is worth it. My previous code was rather inefficient due to redundant and erroneous 'Worksheets("sheets").Activate' commands. The new code cut nearly thirty seconds off the code's run time.       
