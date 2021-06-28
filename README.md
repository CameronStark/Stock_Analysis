# Stocks Analysis - Virtual Basic Applications/Excel

## Overview of Project

### Background

Steve was a big fan of the workbook we created for him (green_stocks). At the click of a button, he can analyze an entire dataset. However, Steve would like to do more research from the last few years. Our code does the required for a dozen stocks, adding potentially thousands more could either choke it up or provide such a long run time it would prove inadequate.

### Directive
This challenge's primary focus is refactoring, making the solution we worked on in module 2 collect all of the data required in a more succinct, quicker fassion. This achievement would determine whether or not refactoring the module 2 code did, in fact, provide a more elegant and efficient VBA code. The written analysis below will discuss the details on the findings.

### Purpose
Refactoring is integral to the coding process. The first run through of determining code for any solution may succeed without issue.  The truth is that while that code may function properly, there may be a better way to accomplish the same goal.  Whether it be through adding comments making the code easier to read or reworking code blocks to perform the same process with less text, refactoring is essential to providing the end user with the best experience.  Ultimately, the goal is to enhance the code's efficiency.  Less run-time means the code is returning the same result as the original, but more quickly.

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

Here I created a nested loop to run through all of our listed variables, starting with tickerVolumes.  It is important to note that j is designated as our iterator instead of i

![image](https://user-images.githubusercontent.com/85717081/123566789-d0dd9b00-d77d-11eb-9b19-7240925cd8c4.png)

Step 3b:

>Write an if-then statement to check if the current row is the first row with the selected tickerIndex. If it is, then assign the current starting price to the tickerStartingPrices variable.

-Here I begin the first If/Then statement, continuing with the j iterator.  The intent is to identify the first occurance of a cell in column 1 with the applicable tickerIndex, demarcating the starting price

![image](https://user-images.githubusercontent.com/85717081/123567018-71cc5600-d77e-11eb-9d12-81bf30e537b7.png)

Step 3c:

>Write an if-then statement to check if the current row is the last row with the selected tickerIndex. If it is, then assign the current closing price to the tickerEndingPrices variable.

-Here we are identifying the last occurance of the applicable cell to tickerIndex, demarcating the ending price

![image](https://user-images.githubusercontent.com/85717081/123567491-83fac400-d77f-11eb-91aa-d1c73eef9c42.png)

Step 3d:

>Write a script that increases the tickerIndex if the next row’s ticker doesn’t match the previous row’s ticker.

 -I must say, this instruction through me off a bit.  My thinking is this is acheived through the For loops.  I am quite uncertain, however.
 
 Below you will find the performance between the 2017 and 2018 analyses, as well as the execution times of the original script and the refactored script.
 
 #### 2017 Analysis
 
![VBA_CHallenge_2017](https://user-images.githubusercontent.com/85717081/123567904-6da13800-d780-11eb-8dcd-633ea82fd0e3.PNG)
![MsgBox2017Analysis](https://user-images.githubusercontent.com/85717081/123567931-7a259080-d780-11eb-8101-25dfad21992f.PNG)

  #### 2018 Analysis

![VBA_Challenge_2018](https://user-images.githubusercontent.com/85717081/123567976-9295ab00-d780-11eb-831b-1f2b74e5a763.PNG)
![MsgBox2018Analysis](https://user-images.githubusercontent.com/85717081/123568004-a3deb780-d780-11eb-8bdb-05f9f173de35.PNG)

#### green_stocks 2017 Analysis

![green_stocks_2017](https://user-images.githubusercontent.com/85717081/123568347-5c0c6000-d781-11eb-8d1b-ed2a9a9479cd.PNG)

#### green_stocks 2018 Analysis

![green_stocks_2018](https://user-images.githubusercontent.com/85717081/123568383-75151100-d781-11eb-9e67-2e762a2ef6f7.PNG)

## Summary

Both the 2017 and 2018 analyises had longer run times after refactoring.  This is a bit baffling and leads me to think I could have accidentally augmented the original code making it more efficient.  The easiest way to determine this is to look through my commits - my incumbent responsibilty.  Nevertheless, we can identify the pros and cons of the refactoring process.

Disadvantages:

- Blocks of code may be copied and pasted to achieve the same goal.  Best practice is to consolidate these blocks into one, making it more efficient.  If you do not know the       proper way to resolve this, you are left relying on the knowledge of others which can be time consuming in itself.
- If you are beginning with another's code, you will have to familiarize yourself with it before you can start streamlining.
- As evident in my own results, refactoring may not always lead to a more efficient solution.

Advantages:

- Adding comments to each step of code helps everyone, let alone yourself, in interpretation
- Adding white space, divvying out succinct sections/blocks of code makes interpretation easier
- Debugging and identifying precesses becomes vastly easier when reformatting the code
- Ideally, all of the above would allow you to correct the original code with something more efficient

We must always consider the refactoring process as an essential piece of the coding process.  The main intent of refactoring is to augment the code for ease of reading and correcting blocks for efficiency.  When given a new set of code written by another, refactoring also provides us the opportunity to streamline and make more efficent a code that has been historically relied upon.  While my own results in terms of run-time may not display the benefits of refactoring, it is abundantly evident that the process is advantageous in terms of interpretation and opportunity for efficiency.
  
