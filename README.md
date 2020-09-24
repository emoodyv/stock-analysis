# Green Stock Anaylsis

## Overview
This project was started in order to help Steve, who is a professional financial advisor, assist his parents in making a smart investment decision. Steve’s parents are avid about protect the environment. As such, they would like to invest in a “green” stock. Therefore this project uses the powerful tools of Excel and VBA (Visual Basic for Applications) in order to analyze the performances of 12 green stocks.

## Anaylsis and Results
In order to accomplish our goal of analyzing the 12 stocks data was needed. So, data was obtained for the following tickers: AY, CSIQ, DQ, ENPH, FSLR, HASI, JKS, RUN, SEDG, SPWR, TERP, VSLR. For each trading day in the years 2017 and 2018 the open, high, close, and volume was complied. Here is an example of what the data looked like:

![Sample Data](https://user-images.githubusercontent.com/71234992/94104841-dbcd5e80-fdec-11ea-83d1-671ee034845a.PNG)

From this raw data a yearly analysis on the stocks’ annual return and annual volume was able to be made. This was made with a VBA macro. The macro was first coded to process one stock at a time. Once this was done the macro proceeded to sum up all the daily volumes. This returned the annual volume. Then, the first open (the start of the year) and the last close (the end of the year) were obtained from the raw data. Subsequently, the ending price was divided by the starting price. This returned the annual return. Here is a snippet of what that code looked like (You can see all the code in the file VBA_Challenge.xlsm):

![Macro](https://user-images.githubusercontent.com/71234992/94105899-25b74400-fdef-11ea-89ad-c4d9f6488e78.PNG)

Now that all the data was processed and the program produced results, the macro was then tasked to clean up and format the results. You can see the results for 2017 and 2018 below. As you can see in 2017 the best performing stock was ENPH and the worst was TERP. Overall, 2017 was a particularly good year for green stocks. In 2018 the best performing stock was ENPH again and the worst was DQ. Unlike 2017, 2018 was disastrous for green stocks.

![2017](https://user-images.githubusercontent.com/71234992/94107455-5ba9f780-fdf2-11ea-9b6f-2d8ac2f945c5.PNG) ![2018](https://user-images.githubusercontent.com/71234992/94107456-5ba9f780-fdf2-11ea-81a5-fc8e2b9799bb.PNG)

## Refactoring Summary
### Refactoring in General
One of the most important steps that was taken in this process was to refactor the code. What that means is to go back through the original code and manipulate it in order to optimize performance, readability, and runtime. In general, and not just in VBA, refactoring code can be a positive process because as aforementioned it can result in optimization. However, one disadvantage of refactoring is that is consumes a copious amount of time, especially in large coding projects.

### Refactoring the VBA code
When refactoring the code in this project several things were changed. Those changes included putting the formatting in the same macro and assigning the total volume, starting price, and ending price to their own arrays. Both these big changes, among others, decreased the time of execution and made the code more readable. Below you can see the pre-refactored code execution times:

![2017](https://user-images.githubusercontent.com/71234992/94109120-62863980-fdf5-11ea-91fc-4c1d4d929022.PNG) ![2018](https://user-images.githubusercontent.com/71234992/94109122-63b76680-fdf5-11ea-9a30-4d0d3f6cea04.PNG)

This is a stark contrast with the following results which are from the refactored code:

![2017 old](https://user-images.githubusercontent.com/71234992/94109121-631ed000-fdf5-11ea-95ef-42e663686075.PNG) ![2018 old](https://user-images.githubusercontent.com/71234992/94109123-63b76680-fdf5-11ea-969a-6cf9822bfe92.PNG)

Overall refactoring the code turned out to have a greatly beneficial effect. If Steve ever desires to use this project to analyze more than 12 stocks or more data the refactored code could possible save him hours of computation.
