# Analysis:  Can Refractoring Make VBA Script Run Faster

## Project and Purpose Overview:  
The initial project was to look at annual stock performance by analyzing the total daily total volume and return of a stock. This led to the creation of a workbook with a quick and easy interface (buttons) to run these numbers on 12 tickers(stocks) with two years of data. This workbook has proven to be useful and there is interest to expand it to analyze more tickers over more years. With more data requirements, the current code may be inadequete and slow. Refractoring the code may have it running optimally.  The purpose of this project is to refract the VBA code to see if it can run the VBA script faster. 

### Results
Refactoring the project with the goal of running VBA script faster included:

    A.  Decreasing the number of loops in the code
    B.  Streamlining formatting
    
 A. _Decreasing number of loops_
 
 The original code for the VBA project had 2 loops:
 
        •	i to loop through the tickers and J to find the total volume, starting price, and ending volume
        
        ![image](https://user-images.githubusercontent.com/84471904/124055902-9a9c5780-d9d9-11eb-93c0-c7fc3158760f.png)

        
    
        

    



