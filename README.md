# Stocks Project with VBA
## Overview of Project
After providing Steve with a means to analyze his dataset of stocks at the click of a button, he would now like to expand that analysis to safely include thousands of stocks in the stock market opposed to dozens. To that end, my intention is to take the code already being used, and refactor it in hopes of making it more efficient.

## Results
The results of my attempt at refactoring were mixed. First of all, the refactored code was supposed to run faster than the original code. Mine did not, as shown below, where we have run times from 2017 and 2018 both, with the refactored code on the left and the original on the right.

![Resources/VBA_Challenge_2017_Comparison.png](Resources/VBA_Challenge_2017_Comparison.png) 

![Resources/VBA_Challenge_2018_Comparison.png](Resources/VBA_Challenge_2018_Comparison.png) 

Rather, my refactored code ended up running roughly .25 seconds slower than the original code. However,  even noting this, it is important to keep in mind that this script is formatting the output of the stocks in addition to giving the volume output and return percentage. This is something the original code did not do. With the original script, the resulting, default output was not visually pleasing and difficult to read. The cells had to be formatted after the fact, and after each year is run to keep it consistent. The refactored script avoids this by incorporating the formatting at the end of the script. Naturally, this code is doing more, so it runs longer, but I would argue it is not by a significant amount, and that the code succeeds in being efficient in that sense.

The exception to this efficiency is that I was unable to get the tickers to increase in the first column. This creates work for Steve, or anyone else who wishes to utilize the code, but less work than having to reformat the code with each run.

It is difficult for me to assess which parts of my code are efficient and which parts are bogging the entire process. down. I made a point to refer back to the original script when writing the code to increase the volume and when assigning the starting and closing prices, but when it came to increasing the tickers in the first column of output the answer was not so apparent. 

I believe most of the issues in my script can be linked to the tickerIndex. In my code I have:
>For i = 0 To 12

>tickerIndex = tickers(i)

>tickerVolumes(i) = 0

This is what I ended up doing to connect the newly created tickerIndex variable back to the tickers array that was originally present. This code, and it being used to access the original tickers array, is possibly cumbersome in the execution of the later code to increase volume and find the starting and ending prices.

A more minor note is that the code provided used an index of 12 in the tickers array, and I made the choice to maintain that in other parts of the script, such as the arrays created, with the thought of, "what if another ticker is added in the future."

## Summary
Overall, this was a challenge I struggled with for a variety of reasons. I feel like I understand the individual concepts of what I was meant to be doing, but perhaps not the finer points of refactoring and how I need to be going about it.

In terms of advantages and disadvantages on the whole of refactoring, I would say that it is an advantage that when refactoring a code you already know exactly what it is you want the code to do. You already know how it runs, and from that can break down how the individual lines of code run. There may be more complicated bits to figure out, such as what to add where and what is extraneous, but the outcome being worked towards is fully fledged out to start with. This is an undeniable advantage and gives something to work with if you are struggling with how to proceed.

The disadvantage would be that, since it is someone else's code that is being worked with, it can be easier to overlook the code's most important lines or miss them entirely. If you delete something, whether intentionally or by mistake, and that absence leads to the code no longer working, it is not something you will necessarily recognize right away, making it harder to identify and fix later on.

Refactoring is constrained, and that is a help as much as it hinders.

As for pros and cons when it comes to this specific VBA script, I exemplify both of them quite well probably. I know what the script is meant to do and I know what each column is meant to contain. I can recognize when the output doesn't look right and where I should be looking to fix that. I know what I am meant to be doing because the code tells me that, but the best way to execute that is another matter. My script got bungled at several points through my attempts at troubleshooting, and since I did not notice those errors in the moment, it created more work. My code was running but I left a next i in the wrong spot without realizing, so it was not running at all and I tried to fix something else that was probably running fine, making a bigger, more time consuming mess to trace back and make sense of.

All in all, it was an effort with mixed success, but with problem areas that are easy to idenitfy, and therefore fix.
