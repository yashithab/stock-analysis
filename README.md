# stock-analysis
VBA Project
#Stock Analysis using VBA

## Overview of Project

Performed stock analysis on the dataset. Developed macro code to loop through all the data one time in order
to identify which stock is performing well and which is not. We wanted to see the total volume of each stock 
from the beginning to the end of the years for 2017 and 2018. Additonally, we wanted to see the percentage return for
each of the 12 stocks over a year. By performing this analysis Steve will be able to recommend which (if any) of those
stocks his parents should invest in. To do this analysis, two macros were used- the original script from the module,
and a refactored version of that script. While the output of the macros was the same, the efficiency of the two versions 
differed, which will be discussed below

## Results for the Perfomance for 2017 and 2018

2017 was a better year in terms of valuation for 11 stocks except for the company "RUN".
DQ" price increased 199.4% but in 2018, decreaesd 62.6%. Overall, this seems like a very volatile sector, and for Steve's 
parents, who can be assumed to be close to retirement, I would suggest that they look at either
index funds, vanguard 20XX funds, or opt for fixed income options. 

There doesn't seem to be a relationship between the stock volume (the amount of which that 
specific stock was traded during a given year) and it's valuation. For example, "HASI" which in 
2017 increased in value by 25.8% was traded a total of 80,949,300 times. In 2018, "HASI"'s value 
decreased by 20.7% while being traded 104,340,600 times (nearly 24,000,000 more trades than in 
2017). Perhaps, if you just looked at that single stock, You could squint and say that there 
might be a very weak, negative relationship between a stocks value and the amount of times it is 
traded. However, "CSIQ" tells a different story. In 2017, "CSIQ" increased by 33.1% while being 
traded 310,592,800 times. Then, in 2018, "CSIQ" stock decreased in value by 16.3% while being 
traded 200,879,900 times (which is about 110,000,000 less trades than in 2017). Here for "CSIQ" 
(and unlike "HASI") the stock was traded about 33% less in 2018 than in 2017, but still posted a 
loss of value like "HASI" did. Overall, there just doesn't seem to be any correlation between a 
stocks volume and it's performance.

## Advantages and Disadvantages of refactoring code

#Advantages:
- The code is more concise, clean and easier to read
- The integirty of the information remains, however, the run-time is more effecient 

#Disadvantages: 
- More scope to introduce bugs into the code, these bugs will be addressed however can be time 
  consuming before making the run-time more effecient. 

## Advantages and Disadvantages of the Original and Refactored VBA Script
The refactored code produced a faster run-time for both years, and in the long run can save a lot of
time.  However, in the process of refactoring the code, I have experienced many bugs to derive a
faster running code. Given the time I had to put into writting it correctly, compared to the small difference
in efficiency gained, (the refactored code was about 28 hundreths of a second faster). 


