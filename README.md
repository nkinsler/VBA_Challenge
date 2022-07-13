# VBA Stock Analysis

## Overview of Project

### Purpose

The purpose of this project is to take the code prepared during module 2 and refactor it to loop through all the data one time.  Upon refactoring, I’m to determine whether refactoring the code successfully made the VBA script run faster.

[VBA_Challenge](https://github.com/nkinsler/VBA_Challenge/blob/main/VBA_Challenge.xlsm)

# Results

The Stock Analysis originally prepared was refactored.  The original time run for the 2017 and 2018 was 0.46875.

![2017_Original](https://github.com/nkinsler/VBA_Challenge/blob/main/Resources/2017%20Time%20-%20Original%20Code.png)!

![2018_Original](https://github.com/nkinsler/VBA_Challenge/blob/main/Resources/2018%20Time%20-%20Original%20Code.png)!

![VBA_Challenge_Code](https://github.com/nkinsler/VBA_Challenge/blob/main/Resources/VBA_Challenge_Code.png)!

The original code was refactored by creating a ticker index to run instead of running a nested loop.  With the ticker index created, the code was modified to accommodate the change by replacing the loop with tickerIndex.

After refactoring the code, an improvement on the time was noted.  The new run time for 2017 and 2018 is 0.078125 seconds and 0.0859375 seconds, respectively.

![2017_Refactored](https://github.com/nkinsler/VBA_Challenge/blob/main/Resources/2017%20Time%20-%20Original%20Code.png)!

![2018_Refactored](https://github.com/nkinsler/VBA_Challenge/blob/main/Resources/VBA_Challenge_2018.png)!

# Summary

## Advantages and disadvantages of refactoring code

The primary advantage of refactoring code is to simplify the code.  A simplified code is easier review and make changes to.  In addition, a refactored code should lead to improved faster run time.  This is important when running the macro through larger data file.
The primary disadvantage of refactoring code is the process is time consuming.  The process requires a time commitment that the user may not be able to fulfill.

## How do these pros and cons apply to refactoring the original VBA script?

The advantage to the refactoring the stock analysis code was the improved run time.  The improved run time was due to the simplified code and use of tickerIndex.  The disadvantage being the time commitment required to refactor the original code.

The advantage to the original code was the ability to fully follow through the code.  The original code followed the lessons learned in the VBA module.  The disadvantage to the original code was the length of the code allowing for more room for errors.
