# stock-analysis
### Submission by Omar Zu'bi

## Project Overview:

### Whats in the repository?
1) The VBA_Challenge.xlsm workbook
2) The Resources folder with VBA_Challenge_2017.png and VBA_Challenge_2018.png

### Purpose and Background:
Steve is looking to help his parents invest in certain stocks and wants to make use of VBA to help him automate the process of analyzing an entire dataset at a click of a button. He hopes that with a little more research, he will be able to effectively consult his parents with what stocks they should invest into. The goal of this project is to create a Macro that can assist Steve with analyzing the entire stock market over the last few years. In addition, a refactored code was created to help improve the script runtime and improve the reusability of a previous created macro to analyze the entire stock market data for any selected year. Data collected from the stock market in 2017 and 2018 were used to help sample the script runtime.  


## Results: 

### Stock Market Data (2017 & 2018)
<p align="center">
  <img width="1000" src= https://github.com/DrZubi/stock-analysis/blob/main/Resources/VBA_Challenge_2017.PNG>
  *Figure 1: 2017 Stock Market Analysis with the refactored macro runtime* 
</p>

As shown in figure 1 above, DQ had highest yearly return in 2017 at 199.4% return, and TERP had the lowest yearly and negative return at -7.2%. 

<p align="center">
  <img width="1000" src=https://github.com/DrZubi/stock-analysis/blob/main/Resources/VBA_Challenge_2018.PNG>
  *Figure 2: 2018 Stock Market Analysis with the refactored macro runtime* 
</p>

As shown in figure 2 above, RUN had the highest yearly return at 84% in 2018, and DQ had the lowest yearly and negative return at -62.6%. Both Run and ENPH and RUN presented a positive yearly return for 2017 and 2018. ENPH alone showed a yearly return of over 80% for both years. 

### Refactored Code Runtime
The refactored code I've created helped improve the runtime of the previous macro by improving the logic flow of the script. The script created used less memory and allowed the code to run in a fraction of a second for each year's analysis. Before refactoring the code, the script runtime for the stock market analysis for both years were around 0.604 and 0.578 seconds respectively. As shown in both figures 1 and 2, the script runtime for each year after refactoring the code improved four times over (0.09 and 0.06 respectively)! 

## Summary Questions: 

1) What are the advantages or disadvantages of refactoring code? 
2) How do those pros and cons apply to refactoring the original VBA script?
