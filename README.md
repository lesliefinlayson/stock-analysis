# Analysis:  Can Refractoring Make VBA Script Run Faster

## Project and Purpose Overview:  
The initial project was to look at annual stock performance by analyzing the total daily total volume and return of a selected group of stocks. This led to the creation of a workbook with a quick and easy interface (buttons) to run these numbers on 12 tickers(stocks) with two years of data. This workbook has proven to be useful and there is interest to expand it to analyze more tickers over more years. With more data requirements, however, the current code may be inadequete and slow. The solution may be refractoring so that the code runs optimally.  The purpose of this project is to determine if the VBA script will run faster by refracting the original VBA code.

### Analysis and Results
Refactoring the project with the goal of running VBA script faster included:

    A.  Decreasing the number of loops in the code
    B.  Streamlining formatting
    
 A. _Decreasing number of loops_
 
 The original code for the VBA project has 2 loops:
 
        •	i to loop through the tickers and j to find the total volume, starting price, and ending volume
        
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

Results for refracting 2017 code:

  The original time to run the code was 2.383813 seconds.  The refracted time to run the code was .3125 seconds.  Approximately 2 minutes is saved by refracting the code for the data in 2017.
    
    
_Original Code for year 2018_

<img width="654" alt="2018_analysis" src="https://user-images.githubusercontent.com/84471904/124061630-16030680-d9e4-11eb-9fa4-d2372c564b67.png">


_Refracted Code for year 2018_

<img width="674" alt="2018_refracted_analysis)" src="https://user-images.githubusercontent.com/84471904/124061752-4fd40d00-d9e4-11eb-8956-329c608ddd23.png">

Results for refracting 2018 code:

  The original time to run the code was 1.984375 seconds.  The refracted time to run the code was .28125 seconds.  Approximately 1.7 minutes is saved by refracting the code for the data in 2018.

## Summary

There are advantages and disadvantages to refactoring code in general.

Some advantages are:
    
         •	may significantly boost system performance by simplifying the code
  
         •	opportunity to make the code more readable
  
Some disadvantages are:

         •	time consuming
  
         •	may introduce bugs

In the case of this project and determining if refractoring the data is beneficial.  It was indeed time-consuming to refractor the VBA script.  However, there are noticeably significant increases in performance by decreasing the amount of time needed to process the data.  Because of the goal to significantly increase the efficiency of the code to have the ability to analyze a greater number of stocks over a greater span of years, the results of refractory justifies the time spent making these changes.  





 
 
    
    


      


        
    
        

    



