# VBA of Wallstreet

## Overview of Project
To help Steve's parents make decisions about which stocks to invest in, stock market data was collected from years 2017 and 2018. Then using VBA, total stocks bought and the percentage of returns was calculated. 

### Purpose
The purpose of this challenge was to refactor VBA code to make the program more efficient and versatile for larger sets of data. Using VBA, the data was colour coded to see separate which stocks have a negative return and which stocks investments are making gains. 


## Analysis of Results 
In 2017, DQ has the highest return of 199%, while TERP has the lowest return percentage of -7.2%. In 2018, DQ's return percentatage drops to -62.6%, while ENPH returns grew from 5.5% in 2017 to 84% in 2018. Since this data is only limited to two years, it's not sufficient to conclude which stock to invest in, as it might be more useful to observe trends over years. All in all, DQ is not the best stock to invest in like Steve's parents had thought. 

## Discussion About Refactoring
Some advantages of refactoring are that it made our code more concise and therefore easier to debug if issues do arise. Also, by making the totalVolumes datatype long, we are able to store larger integers of data, making it more useful for larger data sets and values. Also, the time elapsed to run the program was significantly reduced with refactoring. 

Some disadvantages of refactoring include that such solutions for complex programs may be very hard to come by, especially if there's code throughout the program that relies on each other. This may potentially lead to bugs. Also, having more efficient solutions by making code succinct, may make it more difficult to customize variables, since we're relying on such few loops and if-then statements. 
