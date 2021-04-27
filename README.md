# An Analysis of Ticker Performance

## Overview
The goal of this project was to create an automated VBA-based process to provide an overview of the performance of a selection of 11 tickers. The guidance given for this project noted that potential users of this tool needed the ability to expand the dataset to include additional tickers or yearly data in the future. The outcome of this effort was a more flexible VBA macro that can be run more efficently. The VBA macro provides a snapshot of total trading volume and yearly return for each of the 11 tickers in the intial dataset.

## Results
Throughout this project, two approaches were used to create a macro that would collect the appropriate information for each ticker and format a table for the selected year automattically.

### Method One: Single Array Approach
The first approach used for this analysis utilized a single array that held the value of each ticker label. This method relied on a nexted For Loop that would loop through each ticker then through earch row in the yearly data set and collect information based on whether the first column matched the ticker: (For i = 0 to 11 > ticker=tickers(i) > For j rowStart to rowcount --> If Cells (j,1)=ticker Then...)

The screenshots below show the output of this approach for each year (2017 and 2018) and the duration that it took to run this code. For both years it took at least a full second to loop through all of the data.

#### Method 1: 2017
![Module-2017wFormatting](https://user-images.githubusercontent.com/81983110/116317134-f13ea880-a780-11eb-877d-8a2660b6914a.png)
#### Method 1: 2018
![Module-2018wFormatting](https://user-images.githubusercontent.com/81983110/116316994-c05e7380-a780-11eb-968b-08110a3b6f47.png)

### Method Two: Multi-Array Appraoch
The second approach used for this analysis utilized four arrays (tickers, totalVolumes, tickerStartingPrices, and tickerEndingPrices) and the creation of a tickerindex counter that would increase the index within each of the last three arrays as the script looped through the rows. The goal of this approach was to loop over all of the rows in the sheet a single time, rather than 11 times (once for each of the tickers): For i = rowStart to rowEnd > 'increase tickerVolumes(tickerindex) ... 'find tickerstartingprices(tickerindex) ... 'find tickerendingprices(tickerindex) > If Cells(i+1).Value <> tickers(tickerindex) Then tickerindex = tickerindex + 1

The screenshots below show the output of this second approach for each year (2017 and 2018) and the duration that it took to run this code. Note that the output is identical but the duration for running the macro is noticably shorter (from 1+ s to 0.2s). 

#### Method 2: 2017
![Refactor-2017wFormatting](https://user-images.githubusercontent.com/81983110/116316998-c18fa080-a780-11eb-952d-2cad2d7d91c7.png)
#### Method 2: 2018
![Refactor-2018wFormatting](https://user-images.githubusercontent.com/81983110/116316999-c18fa080-a780-11eb-88dd-c7c6622809d2.png)

## Summary and Conclusions
Both methods are able to deliver results to users and for the scale of the data initially provided, users may not appreciate the difference in approaches. However there are some advantages and disadvantages to both methods. The first method could face potential issues if the number of tickers being analyzed was increased (eg. if users were looking at 50 or 100 tickers) because of the nexted for loops. However, one advantage the first method has over the second is the flexibility around data input. The second method assumes that relevant values for each ticker are in rows chronologically (ie. all of the values for ticker 1 then all of the values for ticker 2) without deviation. If additional information was added for ticker 1 at the bottom of the sheet, the macro could not account for this. However, method 2 is much more scalable and would be the prefered use with two improvements:

1) Before starting the arrays and looping through rows, sort ticker column alphabetically so that we can ensure the macro logic applies
2) Create an additional index so that the tickers array does not need to be populated manually so the code can be flexible to a variety of data sets 
