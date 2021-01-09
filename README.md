# Stock Analysis using VBA

## Overview of Project
The project entails refactoring of the code to run the stock analysis on the dataset provided. The objective of the project is to determine if refactoring the code successfully enables us to run the VBA script faster. 

### Purpose
Steve is looking to run the analysis of the dataset containing information about stocks to make recommendations. We wrote the VBA script to run a macro on the dataset to calculate the yearly return as well as the total volume the stock traded during that year. We are refactoring the original VBA script to ensure that the code runs faster if the dataset size grows. 

### Dataset provided
We were provided with an excel dataset with stock information across 2017 and 2018 in two tabs for each year. The data included columns to indicate the stock, the day of trading, the stock prices on that day along with the volume traded. 


## Results
In the refactored code, we were able to calculate the different metrics for every stock in the dataset in a single for loop and stored the outputs in arrays. Once the analysis was completed, we used a single for loop to output the results contained in the arrays on the worksheet "All Stocks Analysis" followed by formatting to make the output readable. A summary of the refactored code is shown below. 

### Refactored code snapshot
![Refactored code snapshot](https://github.com/dkatragadda/stock-analysis/blob/main/Resources/Refactored%20code.png)

### 2017 Stocks Analysis Summary
The 2017 stock analysis is as follows.

![2017 Stocks Analysis](https://github.com/dkatragadda/stock-analysis/blob/main/Resources/2017%20Stock%20Analysis%20snapshot.png)

### 2018 Stocks Analysis Summary
The 2018 stock analysis is as follows.

![2018 Stocks Analysis](https://github.com/dkatragadda/stock-analysis/blob/main/Resources/2018%20Stock%20Analysis%20snapshot.png)

### Refactored code runtimes
The refactored code ran in about 1/7th of the time it took to run the original code. The original code ran in ~1 second and the refactored code ran in ~0.14 seconds. 

![Code run time for 2017 data](https://github.com/dkatragadda/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

![Code run time for 2018 data](https://github.com/dkatragadda/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

## Summary

### Advantages and disadvantages of code refactoring 
The main advantage of refactoring code comes from improving the efficiency of the code. The goal is to run your code in the most optimal manner by reducing the number of steps taken if possible, reusing chunks of code if possible and making it easier to read. Refactoring the code also helps in collaboration and code reviews as we work in teams to solve problems. 
The disadvantage of refactoring usually involves the size of the code that is being refactored. If the code is too big, it sometimes poses as a challenge to refactor the code. The second disadvantage of refactoring is the amount of time taken to refactor the code. 

### Pros and Cons of refactoring our VBA code
When I considered refactoring the code in this challenge, we simplified the code by running all the calculations in one big for loop and storing all the outputs in arrays which could subsequently be used while displaying the results. This reduced the run time of the code as shown in the screenshots above. My original file (green_stocks.xlsm) also has another efficient step to store the initial array of stocks used in the dataset. As opposed to hard coding the stock tickers in the tickers array, I loaded the array using a for loop and reading from the actual dataset. The primary benefit from doing this is to avoid the need of changing the code if the dataset is altered. The time saved by the refactoring can help Steve run his analysis on much larger datasets if necessary. 
