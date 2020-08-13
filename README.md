# **VBA-Challenge**

### Task

* Create a script that will loop through all the stocks for one year for each run and take the following information:
  * The ticker symbol.
  * Yearly change from opening price at the beginning of the year to the closing price at the end of that year.
  * Percent change from opening price at the beginning of the year to the closing price at the end of that year.
  * The total stock volume of the stock.
* Incorporate confitional formatting that will highlight positive change in green and negative change in red.
* Challenge tasks: 
  * Return the stocks witht the "Greatest % Increase", "Greatest % Decrease", and "Greatest Total Volume.
  * Adjust your script so that it will run on each worksheet by running the code only once.
  
  
#### Screenshots of results

![Results](https://github.com/michaellegg16/VBA-Challenge/blob/master/Screenshots/Output_Results.PNG)

### Instructions

1. Open the excel sheet you wish to run the VBA script in. (The one provided for testing is "alphabetical_testing.xlsm")
1. Make sure that macros are enabled in Excel.
1. Select the view tab.
1. Open the dropdown menu under macros.
1. Select view macros.
1. To edit the VBA code or take a look at it, select the "StockMarketAnalysis" macro and click "edit" on the right.
1. Run the StockMarketAnalysis macro to see the results in each worksheet.

### Conclusion

As each of the seven worksheets has unique results, going through each result would be redundant. After completing this analysis, even though the numbers are fabricated and for experimental purposes only, it is clear how volatile individual stocks can be in the short term. It appears as if some have unlimited potential for gains in the term of a year, with some stocks increasing over 1000% in a year. What I find most interesting is that the potential for losses seems much more limited with the greatest % decrease over a year being less than -100%. This value is a fraction of the greatest % increase, but this model does not have the ability to provide insight as to why without speculation. Further analysis is needed.
