# Stock Analysis With Excel VBA

*Data Source* :  [Stock_Analysis_Data](/VBA_Challenge.xlsm).

## Overview of the Project

### The Data

The Data presents stock information for 12 Different tickers over 3,000 rows of data.

Each stock information contain a ticker value, the date the stock was issued, the opening, highest, lowest, closing and adjusted closing price as well as the volume of the stock. The goal is to collect the ticker, the total daily volume and the return on each stock for the years 2017, 2018.


### Purpose 

The main goal of this project is to use VBA to help us analyze the stocks given in the data defined above. Our goal is to look at each stocks and see the total volume and return that are getting in each year 2017 and 2018. iso as to determine whether the Stocks worth trading and investing in or not. 

The purpose of this project is to edit a VBA code previously used (Module 2 Solution code) by using the method of refactoring. Also, this process was already completed in a similar format, however, the aim of this method is to improve the efficiency of execution of the code, that is said, we just want to make the code run faster than before, by using fewer steps, using less memory, and improving the logic of the code to make it easier for future users to read.



## Results

### Comparaison between the original script and the refactored script

1. Process

First of all, I saved my previous work done in the First Module [*the_green_stocks*](/VBA_Challenge.xlsm). that contain my previous Macros. After, I started following the guidelines provided in the Challenge.
Then, I copy the script given in the file [*VBA_Challenge.vbs*](/VBA_Challenge.vbs) to start adding the appropriate code where indicated.

 Below, the insruction and the appropirate code as wanted:
 
 > Step 1a:
 
 > Create a tickerIndex variable and set it equal to zero before iterating over all the rows. You will use this tickerIndex to access the correct index across the four different arrays you’ll be using: the tickers array and the three output arrays you’ll create in Step 1b.
 
 **tickerIndex = 0**
 
 > Step 1b:

> Create three output arrays: tickerVolumes, tickerStartingPrices, and tickerEndingPrices.
The tickerVolumes array should be a Long data type.
The tickerStartingPrices and tickerEndingPrices arrays should be a Single data type.

**Dim tickerVolumes(12) As Long**

**Dim tickerStartingPrices(12) As Single**
  
**Dim tickerEndingPrices(12) As Single**

> Step 2a:

> Create a for loop to initialize the tickerVolumes to zero. 

**For i = 0 To 11**

**tickerVolumes(i) = 0**

**tickerStartingPrices(i) = 0**

**tickerEndingPrices(i) = 0**

**Next i**

> Step 2b:

> Create a for loop that will loop over all the rows in the spreadsheet.

**For i = 2 To RowCount**

> Step 3a:

> Inside the for loop in Step 2b, write a script that increases the current tickerVolumes (stock ticker volume) variable and adds the ticker volume for the current stock ticker.
Use the tickerIndex variable as the index.

**tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value**

> Step 3b:

> Write an if-then statement to check if the current row is the first row with the selected tickerIndex. If it is, then assign the current closing price to the tickerStartingPrices variable.

**If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value**
            
  **End If**

> tep 3c:

> Write an if-then statement to check if the current row is the last row with the selected tickerIndex. If it is, then assign the current closing price to the tickerEndingPrices variable.
  
**If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value**
        
 **End If**
 
> Step 3d:

> Write a script that increases the tickerIndex if the next row’s ticker doesn’t match the previous row’s ticker.

**If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1**
                
**End If**

> Step 4:

> Use a for loop to loop through your arrays (tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices) to output the “Ticker,” “Total Daily Volume,” and “Return” columns in your spreadsheet.

 **For i = 0 To 11**
        
 **Worksheets("All Stocks Analysis").Activate**
 
**Cells(4 + i, 1).Value = tickers(i)**

**Cells(4 + i, 2).Value = tickerVolumes(i)**

**Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1**
        
 
 2. Outputs
 
 The excution of this code gave us the following outputs:
 
  - First output of this code is a message box asking *"what year would you like to run the analysis on"?* 
 
![Run_time_pop_up_message](/Resources/Pop_up_message.png)

 - We enter the years 2017 and 2018 respectively and we got the outputs below:

***Final VBA Analysis 2017***

![Final_VBA_Analysis_2017](/Resources/run_time_2017.png)

***Final VBA Analysis 2018***

![Final_VBA_Analysis_2018](/Resources/run_time_2018.png)



 - The results above match perfectly the outputs from the AllsStocksAnalysis given in this module. Running our fully 2017 and 2018 data stock analysis gave us an elapsed run time for each year.
 
 Our Finding:
 
  - For 2017 output shows that most of the tickers had a positive return and were safe investments. The Particular one that Steve's parents were looking into was DQ which had the greatest return of all. Also 2017 was a great year for energy companies, 11 out the 12 companies studied had positive returns with 4 returning at least 100%. 
  
  - For the 2018 you can see that the returns were not nearly as high and the total volumes were down as well. That'd said, 2018 was a rough year for the energy sector, only 2 of the 12 companies studied managed to end the year with positive returns. Gains were consolidated to two high volume tickers ENPH and HASI while the rest of the market bled out. DQ was hit hardest by the sector correction losing roughly 63% of its valuation from the start of the year.

## Summary

### Pros and Cons of Refactoring Code

1. Pros

 - Refactoring improves code quality .
 
 - Cleaner and more organized code that include design and software improvement, debugging, and faster programming.
 
 - helps to find bugs, makes programs run faster, easier to understand the code, improves the design of software.
 
 - In our case, using Excel flow displays program logic in a more comprehensible, more concise and straightforward  manner.
 
 - VBA interpretation (Excel) of code can reveal patterns that are not easy to see in the source.
 
2. Cons

 - A long procedure may contain the same line of code in several locations, you can change the logic to eliminate the duplicate lines.
 
 - A logical structure may be duplicated in two or more procedures (possibly via copy & paste coding). When detected, this logic is best moved to a new function and called from the other functions.
 
 - A complex unstructured code is usually best to split in several functions.
 
 - Refactoring process can affect the testing outcomes.
 
3. How do these pros and cons apply to refactoring the original VBA script?

The main benefit that occurred as a result of the refactoring was decreasing the macro run time. The original analysis took approximately one second to run, whereas our new analysis only took about a four of the time (approximately 0.25 seconds) to run.
This is a very positive outcome and means the code is much more efficient, and most likely will run quicker and better if we have more tickers.



        
