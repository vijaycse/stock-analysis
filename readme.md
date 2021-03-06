Stock Analysis With Excel VBA

Click here to view the Excel file: https://github.com/vijaycse/stock-analysis/blob/master/VBA_Challenge.xlsm

### Overview of Project

## Purpose
The purpose of this project was to refactor a Microsoft Excel VBA code to collect certain stock information in the year 2017 and 2018 and determine whether or not the stocks are worth investing. This process was originally completed in a similar format, however, the goal for this round was to increase the efficiency of the original code.

## The Data

The data that is presented includes two charts with stock information on 12 different stocks. The stock information contains a ticker value, the date the stock was issued, the opening, closing and adjusted closing price, the highest and lowest price, and the volume of the stock. The goal is to retrieve the ticker, the total daily volume, and the return on each stock.

### Results

 ## Analysis

Before refactoring the code, I began by copying the code that was needed to create the input box, chart headers, ticker array, and to activate the appropriate worksheet. The steps were then listed out in order to set the structure for the refactoring. Below is the key instruction and code as written in the file.


### Key Refactoring points:

    - In the original code, there were two inner iterations(for loop).
As a result of refactoring, I elimianated the need and made it simple.
That saved time and resources of iterating over 3K records.

 
    
        For i = 2 To RowCount
        
        // logic to calculate
        
         Next i
  
   
   and
    
      For i = 0 To 11
      
        // assign output values into the analysis worksheet
        
      Next i
      

 
### Summary

Refactoring helps make our code cleaner and more organized. A few advantages of a cleaner code include design and software improvement, debugging, and faster programming.  

These disadvantages may range from having applications that are too large to not having the proper test cases for the existing codes, which may ultimately pose some risk if we try to refactor our code.

### Result of Refactoring Stock Analysis:

 The benefit that refactoring this macro saved runtime to 0.1 constant run from  0.6 seconds.

  ##2018 Original:
 
 ![2018 Original](https://github.com/vijaycse/stock-analysis/blob/master/resources/2018_original_run.png)
 


  ##Refactored code execution:

  ![Refactored Code](https://github.com/vijaycse/stock-analysis/blob/master/resources/VBA_Challenge_2017.png)
 
 ![Refactored Code](https://github.com/vijaycse/stock-analysis/blob/master/resources/VBA_Challenge_2018.png)
