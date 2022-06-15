# Stocks Analysis with VBA Code
# Module 2 Challenge

## Overview of project
Analysis of past performance of stocks with VBA code to select appropriate stocks for investing and refactoring the code to improve time efficiency to perform analysis.

## Results

### Stocks performance
Following tables show annual return of stocks for years 2017 and 2018. In 2017 majority of stocks has positive return while in 2018 majority of stocks has negative return. Two stocks ENPH and RUN delivered positive return in both years. The ENPH stock delivered >75% return in both years and it has high trading volume in both years showing interest of investors in the stock. The ENPH stock is a good pick for consistent positive return based on past performance.

<figure>
    <img src="/Resources/stocks_performacne_2017.png" width="200" height="600"
         alt="Stocks Performance 2017">
    <figcaption>Stocks Performance 2017</figcaption>
</figure>

<figure>
    <img src="/Resources/stocks_performacne_2018.png" width="200" height="600"
         alt="Stocks Performance 2018">
    <figcaption>Stocks Performance 2018</figcaption>
</figure>

### Performance of the code before and after refactoring

The original code is written without arrays while refactored code is written with arrays. In the original code calculation and output are looped in same “For” and “If” loops. In the refactored code with arrays calculation and output are in separate loops. After refactoring the time efficiency of the code is improved ~20-40%.---
The links to the codes.---
[Code before refactoring](/Resources/all_stocks_analysis_before_refactor.vbs)---
[Code after refactoring](/Resources/vbs_challenge.vbs)---

<figure>
    <img src="/Resources/run_time_green_stocks_before_refactor_2017.png" width="600" height="100"
         alt="Code Run Time Before Refactor 2017">
    <figcaption>Code Run Time Before Refactor 2017</figcaption>
</figure>

<figure>
    <img src="/Resources/vba_challenge_2017.png" width="600" height="100"
         alt="Code Run Time After Refactor 2017">
    <figcaption>Code Run Time After Refactor 2017</figcaption>
</figure>

<figure>
    <img src="/Resources/run_time_green_stocks_before_refactor_2018.png" width="600" height="100"
         alt="Code Run Time Before Refactor 2018">
    <figcaption>Code Run Time Before Refactor 2018</figcaption>
</figure>

<figure>
    <img src="/Resources/vba_challenge_2018.png" width="600" height="100"
         alt="Code Run Time After Refactor 2018">
    <figcaption>Code Run Time After Refactor 2018</figcaption>
</figure>

## Summary

- General advantages and disadvantages of refactoring code.

    Key advantages of refactoring code are easy to read code, code simplification, efficient knowledge transfer, error reduction and code efficient improvement. Key disadvantages are time and effort to refactor the code. 

- Advantages and disadvantages of the original and refactored VBA script.

	Main advantage of the refactored VBA script is time efficiency. It reduced the code run time by ~20-40%. Main disadvantage is that for small data set of stocks time reduction is not significant considering the time and effort required to refactor the code. 
