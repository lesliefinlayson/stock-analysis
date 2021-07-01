# Analysis:  Can Refractoring Make VBA Script Run Faster

## Project and Purpose Overview:  
The initial project was to look at annual stock performance by analyzing the total daily total volume and return of a stock. This led to the creation of a workbook with a quick and easy interface (buttons) to run these numbers on 12 tickers(stocks) with two years of data. This workbook has proven to be useful and there is interest to expand it to analyze more tickers over more years. With more data requirements, the current code may be inadequete and slow. Refractoring the code may have it running optimally.  The purpose of this project is to refract the VBA code to see if it can run the VBA script faster. 

### Analysis and Results
Refactoring the project with the goal of running VBA script faster included:

    A.  Decreasing the number of loops in the code
    B.  Streamlining formatting
    
 A. _Decreasing number of loops_
 
 The original code for the VBA project has 2 loops:
 
        •	i to loop through the tickers and J to find the total volume, starting price, and ending volume
        
<img width="324" alt="code with 2 loops" src="https://user-images.githubusercontent.com/84471904/124057435-53639600-d9dc-11eb-8b53-7cfcb3d8a4a2.png">

        
The refractored code for the VBA project adjusted and simplified the code to 1 loop:

        •	accomplished by using a tickerIndex and nesting.  No longer needed the second loop.
    
  <img width="324" alt="code with 2 loops" src="https://user-images.githubusercontent.com/84471904/124057709-d84eaf80-d9dc-11eb-87ff-842761062e18.png">
  
 B.  _Streamlining Formatting_
 
 The original code for formatting is an If/Elself/Else command:
 
![image](https://user-images.githubusercontent.com/84471904/124059981-e1da1680-d9e0-11eb-8fd0-445934fc6ac5.png)

The refractored code simplifies the code to an If/Else command:

![image](https://user-images.githubusercontent.com/84471904/124060265-8f4d2a00-d9e1-11eb-9feb-439f39fac1ab.png)

### __Comparison:   Times for original code and refracted code:__

_Original Code for year 2017_

<img width="636" alt="2017_analysis" src="https://user-images.githubusercontent.com/84471904/124060985-f4554f80-d9e2-11eb-906d-f45b6c14a7d7.png">


_Refracted Code for year 2017_

<img width="655" alt="2017_refractored_analysis" src="https://user-images.githubusercontent.com/84471904/124061114-2b2b6580-d9e3-11eb-8cf7-2fcaf608862b.png">









 
 
    
    


      


        
    
        

    



