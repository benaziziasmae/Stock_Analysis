# Stock Analysis With Excel VBA

Data Source: 

## Overview of the Project

### Purpose 

The main goal of the Stocks Analysis is to collect certain Stocks data for the Year of 2017 and 2018 and determine whether the Stocks worth trading and investing in or not.

The purpose of this project is to edit a VBA code previously used (Module 2 Solution code) by using the method of refactoring. Also, this process was already completed in a similar format, however, the aim of this method is to improve the efficiency of execution of the code, that is said, we just want to make the code run faster than before, by using fewer steps, using less memory, and improving the logic of the code to make it easier for future users to read.

### The Data

The Data presents stock information for 12 Different Stocks.

Each stock information contain a ticker value, the date the stock was issued, the opening, highest, lowest, closing and adjusted closing price as well as the volume of the stock. The goal is to collect the ticker, the total daily volume and the return on each stock for the years 2017, 2018.

## Results

### Comparaison between the original script and the refactored script

1. Process

First of all, I saved my previous work done in the First Module *the green_stocks.xlsx* that contain my previous Macros. After, I started following the guidelines provided in the Challenge.
Then, I copy the script given in the file 'VBA_Challenge.vbs' to start adding the appropriate code where indicated.

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

    

