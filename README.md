# Stocks Analysis with VBA

## Overview of Project
Perform analysis on a handful of green energy stocks for Steve; Use Visual Basic for Applications (VBA) to automate calculation tasks on stocks' Total Daily Volume and yearly return in percentage and generate output on Excel sheets in order to determine how well the stocks of interests are performing by year. This way will help Steve determine which are the better stocks he can offer to his parents.

### Purpose
The focus of this particular activity is to refactor the original code wrote before. The main purpose is to see whether the successfully refactored code would make the VBA script run faster, therefore, more efficient for users to read.

## Refactoring Code
First, a ticker Index was created and set as zero. Three output arrays tickerVolumes, tickerStartingPrices, and tickerEndingPrices are made and assigned the appropriate data type.

![First](https://user-images.githubusercontent.com/84931545/124359474-3c4ac100-dbf3-11eb-870c-be5c3329d59f.PNG)

A for loop is create and initialize the tickerVolumes to zero.

![Initialize](https://user-images.githubusercontent.com/84931545/124359489-4ec4fa80-dbf3-11eb-8856-ba2a4cdaab82.PNG)


Next, a for loop looping over all the rows in the worksheet is created with code increase the current tickerVolumes using tickerIndex variable as the index. Inside the for loop, an if-then statement is created setting the tickerStartingPrices if it is the first row with the selected tickerIndex. a second if-then statement is created setting the tickerEndingPrices if it is the last row with the selected tickerIndex. Furthermore, a script is needed to increase the tickerIndex if the next row's ticker doesn't match the previous row's ticker. 

![Next](https://user-images.githubusercontent.com/84931545/124359498-55537200-dbf3-11eb-80a8-296a9a7ac87c.PNG)


Finally, another for loop is created to generate the appropriate output on the worksheet using tickers, tickerVolumes, tickerEndingPrices, and tickerStartingPrices with tickerIndex as the variable.

![Final](https://user-images.githubusercontent.com/84931545/124359506-5b495300-dbf3-11eb-9125-6feb8bbcf76c.PNG)


### Compare with the original code
For the original script of this section, a nested for loop was created, which might take more time to execute. 

![Original_Code](https://user-images.githubusercontent.com/84931545/124359615-b9763600-dbf3-11eb-9def-64a89f887909.PNG)


## Results
After analyzing the results from the refactored code, it is clear that overall, the performances of the stocks in 2017 are better than in 2018. 

### 2017 Results
In 2017, except TerraForm Power (ticker: TERP) had a negative return, other stocks all had a positive return. Among those stocks, Daqo New Energy Corp (DQ), Enphase Energy Inc (ENPH), First Solar Inc (FSLR), and Solaredge Technologies Inc (SEDG) had returns above 100%. 

![Result_2017](https://user-images.githubusercontent.com/84931545/124359554-8cc21e80-dbf3-11eb-8721-123bf0d4e2d9.PNG)


### 2018 Results
On the other hand, in 2018, beside Enphase Energy Inc (ENPH) and Sunrun Inc (RUN) which made positive returns 81,9% and 84.0% respectively. Others are all negative. 

![Result_2018](https://user-images.githubusercontent.com/84931545/124359557-8f247880-dbf3-11eb-9f8d-11b2f86ae3a7.PNG)


### Execution time 
Comparing the execution time between the original and refactored script. 

For 2017, the original script took 0.5019531 seconds to run while it only took 0.0625 seconds to run the refactored script. this means that the new code is more efficient. 

![VBA 2017](https://user-images.githubusercontent.com/84931545/124359565-96e41d00-dbf3-11eb-856a-6a09dadacfe1.PNG)
![VBA_Challenge_2017](https://user-images.githubusercontent.com/84931545/124359571-98ade080-dbf3-11eb-8ce2-216fc7020c93.PNG)

For 2018, a similar result is found. The original script took 0.5019531 seconds to run while it only took 0.0625 seconds to run the refactored script. The new script took a much shorter time to execute. 

![VBA 2018](https://user-images.githubusercontent.com/84931545/124359575-9c416780-dbf3-11eb-8174-56ac46e59c83.PNG)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/84931545/124359581-9fd4ee80-dbf3-11eb-93c2-284ac16e51fa.PNG)

As we can see, the new script is able to shorten the time of analysis by 10 fold. Making it much more convenient to run for users.


## Summary

- What are the advantages or disadvantages of refactoring code?

One obvious advantage of refactoring code is that by testing the refactored code, one can see whether it runs faster than the original one. When it does like in this case, refactored code would be used to provide efficient executions. One disadvantage of refactoring code is that it could be quite time-consuming and for a person who is new to coding this process could be particularly challenging. 

- How do there pros and cons apply to refactoring the original script?

The original script has fewer Arrays and variables in it; it is short; therefore it is easier to see the code. However, because it requires a nested for loop it takes more time to process so it is time-consuming. The refactored script has more Arrays and variables in it, therefore it is long. However, it's not only easier to understand since there is no nested for loops but also faster and easier to execute. 
