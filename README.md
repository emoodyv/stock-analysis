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
