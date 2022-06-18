# Stock Analysis

## Overview of Project

Stock market data from 2017 and 2018 was analyzed to determine which green energy stocks performed the best, so that Steves' parents could make an informed decision on which stocks to invest in.

### Purpose 

Steves parents are planning to invest in DQ stock becuase "DQ" has sentimental meaning for them, but Steve does not think that feelings should determine which stock his parents invest in. The purpose of this project was to help Steve compare DQ stock performance to other green energy stocks. To do this, I refactored the original code to loop through all the data to collect the same information. I wanted to see if the refactored code increased the effieceincy of the original code to help with Steves analysis.

## Results

### Refactoring Code

The first step to refactoring the code was to copy and paste parts of the original code that set up my Worksheet. I copied the code that created the input box, column headers, the ticker array, and the code that activated the correct worksheets needed. Then, I used the template of code given and followed the instructions provided. For this step, I frequently went back to the lesson Modules to help write the code correctly. Here are two screenshots of the refactored code. 
![Screenshot 1](https://github.com/EmLanc2715/Stock-Analysis/blob/main/Code_for_Results1.png)
![Screenshot 2](https://github.com/EmLanc2715/Stock-Analysis/blob/main/Code_for_Results2.png)

### Stock Performance Comparison 

After the code was refactored and executed, I could see that the tables of stock data were the same as the original code. The tables explained that in 2018, most of the stocks performed very poorly. All of the stocks except ENPH and RUN had negative return percentages. DQ had a -63% return for this year. However, the 2017 stock analysis table showed that all stocks except for TERP had a positive return value. DQ actually performed the best at a 199.4% return in 2017.

### Execution Times

The execution times of the refactored code was signigifcantly less than the execution time in the original code. The original code took just over 1 second when I ran the origianl script for both worksheets. However, for the refactored 2017 code the execution time was 0.14066 seconds and for the refactored 2018 code the execution time was 0.11719 seconds.
![Screenshot of execution time for 2017](https://github.com/EmLanc2715/Stock-Analysis/blob/main/VBA_Challenge_2017.png)
![Screenshot of execution time for 2018](https://github.com/EmLanc2715/Stock-Analysis/blob/main/VBA_Challenge_2018.png)

#### Summary
- What are the advantages or disadvantages of refactoring code?

    The advantage of refactoring code is that it helps make the code be more organized. Also, it helps code run a lot faster and more efficiently. A disadvantage of refacotring code is that, if refactored code isn't explained properly through comments, it could become very confusing to someone who is trying to add to or decipher it. Also, if the code isn't refactored properly, new erros can occur, whill take time to debug.

- How do these pros and cons apply to refactoring the original VBA script?

    The pros apply to refactoring the original script because the set up was already done for me. I took the outline of code and adjusted what I needed to in order to follow the instructions. Starting from scratch would have been more time consuming. Also, an important pro was that the refactored code ran so much faster than the original. The execution time didn't matter as much with this data, but would matter a great deal with high volumes of data.

    The cons apply to refactoring the original script because I did run into compiling errors while working to refactor the code. Most of the errors were due to errors in my original code, but it did take some time to figure out howto fix them.