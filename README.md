# An Analysis of Ticker Performance

## Overview
The goal of this project was to create an automated VBA-based process to provide an overview of the performance of a selection of 11 tickers. The guidance given for this project noted that potential users of this tool needed the ability to expand the dataset to include additional tickers or yearly data in the future. The outcome of this effort was a more flexible VBA macro that can be run more efficently. The VBA macro provides a snapshot of total trading volume and yearly return for each of the 11 tickers in the intial dataset.

## Results
Throughout this project, two approaches were used to create a macro that would collect the appropriate information for each ticker and format a table for the selected year automattically.

### Method One: Single Array Approach
The first approached used for this analysis utilized a single array that held the value of each ticker label. This method relied on a nexted For Loop that would loop through each ticker then through earch row in the yearly data set and collect information based on whether the first column matched the ticker: (For i = 0 to 11 > ticker=tickers(i) > For j rowStart to rowcount --> If Cells (j,1)=ticker Then...)

![Module-2018wFormatting](https://user-images.githubusercontent.com/81983110/116316994-c05e7380-a780-11eb-968b-08110a3b6f47.png)
![Refactor-2017wFormatting](https://user-images.githubusercontent.com/81983110/116316998-c18fa080-a780-11eb-952d-2cad2d7d91c7.png)
![Refactor-2018wFormatting](https://user-images.githubusercontent.com/81983110/116316999-c18fa080-a780-11eb-88dd-c7c6622809d2.png)



The written analysis contains the following structure, organization, and formatting:

There is a title, and there are multiple paragraphs (2 pt).
Each paragraph has a heading (2 pt).
There are subheadings to break up text (2 pt).
Links are working, and images are formatted and displayed where appropriate (2 pt).
Analysis Requirements (12 points)
The written analysis has the following:


Results
The analysis is well described with screenshots and code (4 pt).
Summary
There is a detailed statement on the advantages and disadvantages of refactoring code in general (3 pt).
There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script (3 p
