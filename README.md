# **VBA Stock Analysis Project**

## **Overview**

### This analysis was performed upon request of an investment client who wanted an easy way to evaluate the yearly performance of multiple stocks across various years. We used VBA to refactor a dataset including stocks from 2017 and 2018, in order to quickly and efficiently loop through very large data sets including thousands of stocks, producing an easy to read table of measures. Formatting such as color coding will also make it easy for users to determine whether specific stocks would be suitable investments. 

## **Results**

### A previously utilized code served as the template for refactoring so that it would be capable of executing functions on large spreadsheets with a higher volume of data. Our refactored code is capable of looping through the data once and returning measures such as total daily volume and yearly return for each stock ticker, with a significantly improved run time. The code is annotated in a way that users can easily understand, allowing individuals with minimal VBA experience to make minor edits and adjustments, in order to use it on other data sets if desired. 

### The edits we made in refactoring the original code include the following:
	 
##### * Creating a tickerIndex variable and setting it at zero

##### * Creating tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices variables, and setting appropriate output arrays for them
	
##### * Creating a nested For loop which initializes the tickerVolumes variable at zero, then increases the tickerVolumes variable and adds the current volume for each ticker

##### * Constructing a conditional If...Then statement in order to assign either the current starting or ending price to the respective tickerStartingPrices and tickerEndingPrices variables, depending on whether the current row is the first or last to have the tickerIndex variable selected; this statement also increases the tickerIndex variable when the ticker for the following row is for a different stock than the previous one. 
	
##### * A second For loop was created in order to quickly loop through the arrays and provide output measures to a spreadsheet, such as “Ticker,” “Total Daily Volume,” and “Return”
	 
##### * Finally, formatting code was used to make the output in the previously mentioned spreadsheet easy for users to reference and understand: well-labeled columns were constructed and the yearly return value cells appear in red or green, depending on whether the final values are positive or negative 

### An initial analysis of 12 tickers from 2017-2018 revealed that, with two exceptions, all other stocks which had performed well in [2017](https://github.com/elissewright/stock-analysis/blob/main/VBA_Challenge_2017.png), finished with negative annual returns in [2018](https://github.com/elissewright/stock-analysis/blob/main/VBA_Challenge_2018.png). Analysis for both years took around 0.07 of a second to run, a negligible time compared to the near-minute long run time of the original code. 
    
### Compared to the average stock market return of 10%, using S&P as a [benchmark](https://www.nerdwallet.com/article/investing/average-stock-market-return), many of these stocks showed exceptionally bad performance in 2018 and may be very risky investments. ENPH and RUN, which showed smaller but still positive gains in 2018 as compared to 2017, may be less risky options for prospective investors. 

### It must be acknowledged that an analysis of two years is very limited and not capable of revealing trends over long periods of time. There is a chance that some of these stocks which performed poorly across 2017-2018 may have a regular history of volatility and significant gains as well as losses from year to year. An analysis of a larger spreadsheet incorporating at least 10 years of history may show some of these long-term trends, and formal evaluations such as Mann-Kendall tests could be used to assess statistical significance of any fluctuations over time. 
## **Summary**

### While our refactoring enhanced the performance and capabilities of this program, refactoring existing code can come with both advantages and disadvantages depending on the circumstances. In addition to helping save costs and time, another advantage of refactoring existing code is that the template allows us to identify and fix bugs that may have been present. And the process can also simplify a code that was messy or difficult to understand. 

### But just as we can repair bugs or clean up messy code, refactoring can also introduce these problems into an adequately working code if the programmer is not careful. Refactoring is also not always an appropriate solution, depending on the code and specific needs in question: it may make a program too long or slow to run, or fail to improve upon an outdated code that needs to be replaced from scratch. 

### In the case of our analysis, refactoring a pre-existing code was an appropriate course of action, as the code worked and was not outdated for the purposes it was initially created for, but simply slow and limited in usefulness. Our refactoring made it more efficient and widened the applications and datasets it may be used for, as well as making the code easier to read and edit by other programmers in the future. 
