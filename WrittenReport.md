# Stock Analysis
## Overview of Project
The purpose of this project is to refractor previous code to be able to have the code run more efficiently.   This is done because the client, Steve, will be expanding the current dataset to include the entire stock market.    Without the refractorization of the code, the larger dataset would take an extremely long time to execute.    

## Results: 

By refractoring the code, the code was able to execute significantly faster.  

- 2017:  Original code ran in:  1.1797 seconds and the refractored code ran in:  0.1953 seconds.

  ![VBA_Original_2017.png](/Resources/VBA_Original_2017.png)
  [VBA_Original_2017.png](/Resources/VBA_Original_2017.png)

  ![VBA_Challenge_2017.png|width=300px](/Resources/VBA_Challenge_2017.png)
  [VBA_Challenge_2017.png](/Resources/VBA_Challenge_2017.png)


- 2018:  Original code ran in:  1.1875 seconds and the refractored code ran in:  0.2227 seconds.

  ![VBA_Original_2018.png](/Resources/VBA_Original_2018.png)
  [VBA_Original_2018.png](/Resources/VBA_Original_2018.png)

  ![VBA_Challenge_2018.png](/Resources/VBA_Challenge_2018.png)
  [VBA_Challenge_2018.png](/Resources/VBA_Challenge_2018.png)

The original code consisted of nested `for` statements to run through the data.    
![OriginalCode.png](/Resources/OriginalCode.png)

For the refractored code, a stock ticker index array was utilized.    This led to a cleaner and more efficient run of the data.
![RefractorCode.png](/Resources/RefractorCode.png)




## Summary: 
- What are the advantages or disadvantages of refactoring code?
  - The advantages of refractoring code is that the code is more efficient and written cleaner.
  - The disadvantages of refractoring code is that the code can be more complicated to read and more comments are necessary to document the steps.

- How do these pros and cons apply to refactoring the original VBA script?
  - The refractored code was significantly quicker at processing the data.   
  - The ability to write an array based on the ticker index variable was more difficult to understand and took more time to write.    
