# stock-analysis for 2017 & 2018
## Overview of Project
The purpose of this project is to refractor previous code to be able to have the code run more efficiently.   This is done because the client, Steve, will be expanding the current dataset to include the entire stock market.    Without the refractorization of the code, the larger dataset would take an extremly long time to execute.    

## Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

By refractoring the code, the code was able to execute in significantly faster.  

-2017:  Original code ran in:  1.1799 seconds and the refractored code ran in:  .1953 seconds
[VBA_Original_2017.png](/Resources/VBA_Original_2017.png)
[VBA_Challenge_2017.png](/Resources/VBA_Challenge_2017.png)

-2018:  Original code ran in:  1.1885 seconds and the refractored code ran in:  .2227 seconds
[VBA_Original_2018.png](/Resources/VBA_Original_2018.png)
[VBA_Challenge_2018.png](/Resources/VBA_Challenge_2018.png)

The original code consisted of nexted for statements to run through the data.    For the refractored code, a stock ticker index array was complied.    This led to a clearner and more efficient run of the data.




## Summary: 
-What are the advantages or disadvantages of refactoring code?
The advantages of refractoring code is that the code is more efficient and written cleaner.
The disadvantes of refractoring code is that the code can be more complicated to read and more notes are necessary to document the steps.

-How do these pros and cons apply to refactoring the original VBA script?
The refractored code was significantly quicker at processing the data.   The ability to write an array based on the ticker index variable was harder to understand and took more time to write.    
