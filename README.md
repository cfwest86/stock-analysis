# 2017-2018 Green Eneregy Stock Performances

## Overview of Project
A client has requested we automate analysis of stock data performance of 12 green energy companies for the years 2017 and 2018. Utilizing VBA in Excel, macros were created that allowed our client to easily summarize the Ticker, Totaly Daily Volume, and Yearly Return for each companies stocks, and output and format the results in a new worskheet. Subseqeuently, we refactored our code to make the analysis more efficient, and enable larger data sets to be analyzed if desired in the future.

## Results

### Stock Performance 2017 vs 2018

![2017eprformance](https://user-images.githubusercontent.com/81761879/116826937-992eea00-ab64-11eb-9de5-5cb823390c50.PNG)
![2018performance](https://user-images.githubusercontent.com/81761879/116826939-9d5b0780-ab64-11eb-862f-78efd5813f5f.PNG)

As illustrated in the above images, 2017 was a much better year in terms of returns for the tracked energy stocks versus 2018. Everall, only 2 stocks, ENPH and RUN saw gains for both years, indicating these 2 stocks seem to be fairly profitable year over year. We can also see these two stocks had among the highest volumes of daily trading of the stocks listed. 

### Refactoring Performance 

![VBA_Challenge_2017_Original_Code](https://user-images.githubusercontent.com/81761879/116827089-60dbdb80-ab65-11eb-930a-db104b8b1f05.PNG)
![VBA_Challenge_2018_Original_Code](https://user-images.githubusercontent.com/81761879/116827110-6c2f0700-ab65-11eb-8b4d-6d84e811882b.PNG)

As evident in the above images, the run time of the code before refactoring was roughly between 0.63 and 0.64 seconds for each year, which on the surface seems fairly quick.  

![VBA_Challenge_2017](https://user-images.githubusercontent.com/81761879/116827150-9f719600-ab65-11eb-86bc-cb3bd9df2584.PNG)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/81761879/116827153-a3051d00-ab65-11eb-96b1-47c0ed34fe4d.PNG)

However, as  shown in the above images, by refactoring the code we were able to increase the speed of our code by roughly 440%, down to roughly between 0.140 and 0.145 seconds for each year. This may not seem especially important for this relatively small data set, however if our client wished to analyze the stock performance of 100's of green energy companys vs. the 12 analyzed above, the performance difference could be quite pronounced. 

We were able to achieve this performance increase, by including a stock index in our code which allowed our macro to reduce the number of times it needed to run through rows of data.

## Summary

Refactoring code has both benifits and draw backs. Obviously, refactoring code can allow for greatly increased performance which frees up resources to analyze large data sets. In adition, refactoring code can simplify the structure allowing for an easier understanding of what is occuring in said code. However, refactoring also has its downsides. To begin, it can take additional human resources and time to run through existing code. In addition, unexpected bugs may be introduced if the individual/s refactoring do not have a complete understanding of what each line of code is accommplishing. 

Specific to the base and refactored code for this assignment, while the original code was fiarly striaghtforward and took little time to write, refactoring required a significant understanding of each step of the code, and resulkted in a number of errors and subsequent debugging and researtch in order to finally arrive at a working refactored code. 
