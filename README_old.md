# Stock Analysis

## Overview of Project

### Steve, who is a good friend, needed to help his parents pick a stock that performed well on the market. He did a good job providing us with a spreadsheet that narrowed the scope to a subset of stocks. Steve's parents' goals were to identify stocks that traded at a high volume, and stocks with a high return on investment. We built a stock analyzer using Excel's Visual Basic language. This analyzer will help Steve pick a high performing stock he can recommend to them.

## Results

### Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

### We used VBA Scripts to build a ticker that summarized each stock's total daily volume, and the return. We used the array variable to build our custom stock ticker as outlined in the code quoted below:

```
    'Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
```

### In order to allow Steve to choose the year he wanted to analyze, we entered an input variable with a prompt:

```
    yearValue = InputBox("What year would you like to run the analysis on?")
```
### This variable then generated a window in Excel that allowed Steve to enter the year he wanted to analyze, which was 2017 followed by 2018 in a separate instance
![Text Prompt](https://github.com/rivas-j/stock-analysis/blob/4a0563b137d8957645c4057297d97e788e6b343a/Resources/VBA_Challenge_textprompt.png)

### **Stock Tickers**

### Below you will see the results of our VBA code. We built these helpful charts using for loops with conditional statements. We also want to highlight the effect refactoring our code had on the execution time. In the ticker below, we have one column listing the total volume, and another column calculating the return for each stock:
![2017 Stock Analysis](https://github.com/rivas-j/stock-analysis/blob/568bf21329df5e139e59684e1c3e78cbd085732a/Resources/VBA_Challenge_2017.png)

### Below we can see the execution time of our script before refactoring:
![2017 Before Refactor](https://github.com/rivas-j/stock-analysis/blob/26e7f0b6e665f4f40ee42b83d67d93dc684adc11/Resources/2017%20Before%20Refactor.png)

### Here is our ticker listing total volume and return for each stock in 2018: 
![2018 Stock Analysis](https://github.com/rivas-j/stock-analysis/blob/4a0563b137d8957645c4057297d97e788e6b343a/Resources/VBA_Challenge_2018.png)

### Please also take note of the execution time of our script before refactoring:
![2018 Before Refactor](https://github.com/rivas-j/stock-analysis/blob/26e7f0b6e665f4f40ee42b83d67d93dc684adc11/Resources/2018%20Before%20Refactor.png)

### Based on these charts, if I am Steve, I suggest to Mom and Dad to invest in ENPH. It yielded positive returns in both years, and increased in trade volume after 2017

## Summary

### Stock market analysts tend to deal with large amounts of data. When designing macros in Visual Basic, we want to keep efficiency in mind to mitigate performance issues wich such large data sets. We accomplished this by refactoring our code, reducing the time it took for our script to run. This is very beneficial because it allows us to analyze large data sets quickly, without using too many CPU resources.

### One of the drawbacks of refactoring code is that it might be too time consuming, depending on how the previous developer formatted the script. If the previous developer didn't use whitespace effectively, or include helpful comments throughout the script, it might be better to build a new macro.
