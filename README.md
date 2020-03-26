# Stock Market Analysis using VBA

## Final Output

In this project, I used VBA to calculate the annual performance and trading volume of stocks traded in the United States from 2014 to 2016. In addition to tracking individual stock performance, I also found each year's best and worst performance in percentage terms from the first day each stock traded that year, and the stock with the highest trading volume each year. For ease of viewing, the format of the cell containing each stock's annual performance was changed based on if the stock increased or decreased for the year. 

## Data

The data I used contained the opening, closing, high and low prices, and the trading volume for each stock for each weekday of the year. The data was first sorted by year, and then by stock ticker symbol. 

## Limitations to data

My data does not contain each stock market capitalization. Stocks with very small market capitalization can be subject to large price swings, and market manipulation by insiders. Some investment portfolios also focus on stocks with market capitalizations within a pre-defined range. The lack of market capitalization data prevented me from filtering out stocks with low market capitalization, and from filtering by a market capitalization range that my end-user may be interested in. 

My data also does not contain industry membership of each stock. This prevents me from measuring the performance of each stock compared to other stocks in similar industries, and to compare industry performance as a whole. 
