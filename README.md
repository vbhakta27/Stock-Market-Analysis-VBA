# Stock Market Analysis - VBA

Created a script that looped through all the stocks for every worksheet, every year and took the following information:
- The ticker symbol
- Yearly change from opening price at the beginning of a given year to the closing price at the end of that year
- The percent change from opening price at the beginning of a given year to the closing price at the end of that year
- The total stock volume of the stock

Each stock worksheet initially contained the ticker symbol, date, open price,	high price,	low price, close price, and volume for that day. 

![](/VBAMarket/Original_Stock_Data.png)

The script has conditional formatting that will highlight positive change in green and negative change in red. It also returns the stock with the "Greatest % increase", "Greatest % Decrease", and "Greatest total volume".

Below are the results of the VBA script.

![](/VBAMarket/2014.png)
![](/VBAMarket/2015.png)
![](/VBAMarket/2016.png)

