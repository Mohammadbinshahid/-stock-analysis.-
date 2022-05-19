# Stock Analysis

##Overview of the Project

### Purpose

1.	To refactor a Microsoft Excel VBA code to collect certain stock information in the year 2017 and 2018 
2.	Determine whether or not the stocks are worth investing. 

The process was completed during our training module but the purpose was to alter the code to reduce the run timeline

#### The Data
The data presented has 2 charts with information on 12 stocks covering returns for the periods 2017 and 2018. Data includes daily stock ticker information showing market opening, closing, Adjusted Closing, high and low prices as well as the daily traded volume of each stock.

#### Goal
Our goal aimed at sorting data to segregate stock tickers and finding out the total annual traded volume to assess the liquidity of each stock; and calculate the return over two years.

### Results & Analysis
I analyzed the results of our analysis to conclude the performance of all stocks was better in 2017 as compared to 2018. This could be relevant to an overall bull run in the stock market in the year 2018.
The tickers DQ, ENPH and SEDG were the best performing stocks in 2017. While in 2018, most stocks underperformed, ENPH had a return of 81.9% and hence was the best performing stock for two years in terms of consistency.
I calculated time weighted rate of return of DQ, ENPH and SEDQ and found out that ENPH had a TWRR of 317.5% over two years as compared to SEDQ’s return of 162.3% as DQ’s return of 12.0%. ENPH was also the most liquid stock in our universe. 
Based on this analysis, one would be most inclined to invest in ENPH however information analyzed is insufficient to forecast future returns. The stock may already have achieved its target value and historical performance is not an indicator of future returns. Forward looking forecasts and fair value calculations are required in order to give a good stock recommendation.

### Advantages of Refactoring Code to our analysis
By refactoring the code I was able to reduce runtime significantly for both years as given in the attached snapshots. Our original code took 0.66 seconds to run for 2018 whereas the refactored code ran at 0.11 seconds hence reducing the macro run time by 5 times.

### Pros of Refactored VBA script
Cleaner and more organized code leads to design improvements, debugging and programming ease for future users

### Cons of Refactored VBA script
Refactoring can be time consuming, depending upon the complexity of the code you may not have an assessment of how long it will take to refactor and whether there will be a significant improvement after refactoring
