# Stocks Analysis - Virtual Basic Applications/Excel

## Overview of Project

### Background

Steve was a big fan of the wrokbook we created for him (green_stocks). At the click of a button, he can analyze an entire dataset. However, Steve would like to do more research from the last few years. Our code does the required for a dozen stocks, adding potentially thousands more could either choke it up or provide such a long run time it would inadequate.

### Directive
This challenge's primary focus is refactoring, making the solution we worked on in module 2 collect all of the data required in a more succinct, quicker fassion. Acheiving the latter would determine whether or not refactoring the module 2 code did, in fact, provide a more elegant and efficient VBA code. The written analysis below will provide details on the findings.

### Purpose
Refactoring is integral to the coding process. The first run through of determining code for any solution may succeed without issue.  The truth is that while that code may function properly, there may be a better way to acheive the same goal.  Whether it be through adding comments making the code easier to read or reworking code blocks to perform the same process with less text, refactoring is essential to providing the end user with the best experience.  Ultimately, the goal is to enhance the code's efficiency.  Less run-time means the code is acheiving the same result as the original, but more quickly.

## Deliverables

### -Deliverable 1: Refactor VBA code and measure performance
  
  -This deliverable will include an updated workbook and a folder with PNGs of the pop-ups with script run time
  
###  -Deliverable 2: A written analysis of your results (README.md)

  -Fairly self explanatory here, it's what you're reading :)

## Results

After performing the five steps that involve downloading the starter code and creating a resources folder, the following steps for deliverable 1 were acheived as follows:
  - Step 1a:
>Create a tickerIndex variable and set it equal to zero before iterating over all the rows. You will use this tickerIndex to access the correct index across the four different arrays you’ll be using: the tickers array and the three output arrays you’ll create in Step 1b.

  -After establishing the row count and before setting the variables we set our tickerIndex to start at zero

![image](https://user-images.githubusercontent.com/85717081/123566091-10a38300-d77c-11eb-8f15-e06327c1d6a9.png)

  - Step 1b:

>Create three output arrays: tickerVolumes, tickerStartingPrices, and tickerEndingPrices.
  - The tickerVolumes array should be a Long data type.
  - The tickerStartingPrices and tickerEndingPrices arrays should be a Single data type.

-Here we set the pertinent data types to our variables

![image](https://user-images.githubusercontent.com/85717081/123566363-c2db4a80-d77c-11eb-8cc8-2f9523742537.png)

Step 2a & 2b:

>Create a for loop to initialize the tickerVolumes to zero.
>Create a for loop that will loop over all the rows in the spreadsheet.

For loop created and setting the tickerVolumes to start at zero

![image](https://user-images.githubusercontent.com/85717081/123566610-51e86280-d77d-11eb-961c-ffd3c6928960.png)

Step 3a:

>Inside the for loop in Step 2b, write a script that increases the current tickerVolumes (stock ticker volume) variable and adds the ticker volume for the current stock ticker.
Use the tickerIndex variable as the index.

Here I created a nested loop to run through all of our listed variables, starting with tickerVolumes

![image](https://user-images.githubusercontent.com/85717081/123566789-d0dd9b00-d77d-11eb-9b19-7240925cd8c4.png)





