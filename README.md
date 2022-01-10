# Stock Analysis

## Overview

This analysis is being done on the stock market over the years to hel Steve fine the best stock for his parents. Using VBA I created a code to show us which stocks did the best over which years. This will help Steve figure out which stock his parents should put their money into. 

## Results

When we run the code it will show us stocks for whichever year you enter into the popup box. This will calculate weather the stocks lost or gained monoey and by how much. We can see that the stocks in 2017 did much better than 2018 in the images below. 

<img width="321" alt="2017" src="https://user-images.githubusercontent.com/94948877/148708898-6e35bb29-2fd8-4533-a057-ad8938b21054.png">
<img width="314" alt="2018" src="https://user-images.githubusercontent.com/94948877/148708903-eaece22f-519e-4358-95b7-2ec77da4f59b.png">

His parents were wanting to invest in DQ stock but in 2018 the stock did not do well so they will probably need to look into other stocks. The images below are the execution times of the code. 

![VBA_Challenge_2017](https://user-images.githubusercontent.com/94948877/148712346-1ec2ff68-b8ed-4d4e-b950-b6874325c61c.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/94948877/148712359-fee3b68c-532e-4103-9945-0ff3f01571ea.png)

When I first used the refactored code the runtime was over 40 seconds, I noticed I had the "Worksheets(yearValue).Activate" inside the for loop so I believe it was activating the sheet every loop and that is what was taking so long. Shown below:

<img width="245" alt="Screen Shot 2022-01-09 at 7 01 31 PM" src="https://user-images.githubusercontent.com/94948877/148713666-0d6184c9-6697-48a3-a536-cf09bcf79910.png">

When I fixed that the time was about the same as the other code. Fixed code below:

<img width="236" alt="Screen Shot 2022-01-09 at 7 00 19 PM" src="https://user-images.githubusercontent.com/94948877/148713589-24b33195-7074-4815-96c8-0c85b0273de8.png">

## Summary

Refactoring code can help you figure out how you may be able to combine macros to help things run smoother. It can also help clean the code up to be easier to read. When refactoring you also have a chance to create an error in the code and would have to troubleshot it. 

In my code the analysis ran but then the formatting code needed to be run separateley to get the boxes colored based on the percentage. I created two buttons to run the analysis and then to format it. Running the refactored code you can attach that macro to the button and it runs the analysis and formats the cells all in one.
