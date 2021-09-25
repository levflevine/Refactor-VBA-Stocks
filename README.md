# Refactor-VBA-Stocks

## Overview of the Project ##
The project had three goals. 
- Provide Steve with **an analysis of the green energy stock performance accross the market** to help make an investment decision for his family, 
- **Review the performance of the *DQ* stock** and generate the recommendation on whether or not invest in it, and 
- Conduct the **refactoring** of the code to make sure that the **analysis algorithm is efficient** and is able support the processing of large volumes of stock data. 

Steve has obtained the **[2017 and 2018 stock performance dataset](https://2u-data-curriculum-team.s3.amazonaws.com/dataviz-online/module_2/green_stocks.xlsx)**, which is used in this project. 

The dataset includes the following data points:

  - ***Ticker*** - a symbol used to identify publicly traded stocks
  - ***Date*** - the date of stock trade 
  - ***Open*** - the openning price of the stock at the beginnign of a trading session on a particular day
  - ***High*** - highest price of the stock during a trading session on a particular day
  - ***Low*** - lowest price of the stock during a trading session on a particular day
  - ***Close*** - the closing price of the stock at the conclusion of a trading session on a particular day
  - ***Adj.Close*** - adjusted closing price that factors in corporate actions such as dividents, stock splits, and etc. 
  - ***Volume*** - number of stock shares that have been traded in a session on a particular day
 
The following **metrics** has been generated for each stock ticker in the dataset:
1. **Total Daily Volume** = the total ***sum*** of the daily stock ***Volumes*** over the period of a year
2. **Stock Return** = (the ***Close*** price of the stock at the last trading session of a year) / (the ***Open*** price of the stock at the first trading session of a year) ***- 1***

The analysis has been conducted using **VBA** and included **2 interations**. 

The ***first iteration*** of the anaslysis was based on the following algorithm:
  - Outer loop through each stock Ticker
    - Inner loop through the entire dataset 
      - Calculation of the **Total Daily Volume** and **Stock Return** for each **Ticker**
        - Calculation of the compute time to complete the analysis

The ***second iteration*** of the anaslysis was based on a more efficient algorithm developed by **refactoring** the ***first iteration***:
  - Loop through the entire dataset 
    - Using the TickerIndex counter to keep track of each **Ticker**
      - Calculation of the **Total Daily Volume** and **Stock Return** for each **Ticker**
        - Calculation of the compute time to complete the analysis
        
## Results ##

The analysis the dataset has demonstrated that in 2017, most of the green stocks have been profitable with the exception of *TERP*. Also, most of the stocks had large trading volumes, except for *DQ*. The most profitable stocks in 2017 were: *DQ*, *ENPH*, *FSLR*, and *SEDG*.

In 2018, the performance picture has signifacantly changed. Vast majority of the stocks have demonstrated negative profitability, except for *ENPH* and *RUN*. All the volumes have been significant throughout the year.  

The compute times for the ***first*** and ***second*** iterations of the code algorithm were **significantly different**, as follows.

|             |  **First Iteration**  |  **Second Iteration**  |
| ------------|-----------------------|------------------------|
|    2017     |       **0.83 sec      |      0.13 sec**        |
|    2018     |       **0.81 sec      |      0.13 sec**        |

The results of the analysis and compute times of the **second (refactored) iteration** of the code is shown below.

### 2017 Refactored Stock Analysis ###

![Results 2017](/Resources/VBA_Challenge_2017.png)

### 2018 Refactored Stock Analysis ###

![Results 2018](/Resources/VBA_Challenge_2018.png)

## Summary Conclusion ##

**Advantages or disadvantages of refactoring code?**

  ***Pros:***
  - Reducing the compute times for the code execution
  - Reducing the memory / storage requirements to run the code
  - The code gets better organized and structured

  ***Cons:***
  - No new product features are created (the functions of the app remain the same)
  - Additional time and resources spent on refactoring (incremental costs)
  - Risk of adding more bugs to the code that may result in additional testing and debugging efforts (incr. costs)

**How do these pros and cons apply to refactoring the original VBA script in the Project?**
   1. The refactoring enabled an **83% improvement of the code execution times** in the current project
   2. The refactoring required **incremental 1 hour of additional coding and debugging efforts**
 
 
