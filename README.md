# Challenge

So in order to refactor our code, we needed to forgo the nested loops. As an alternative, I have created multiple arrays for our for loops to loop through. These arrays are TickersVolume, TickersStartPrice, and TickersEndprice. I have also created a TickersIndex which is now used to hold which Ticker the former arrays will be looped for.

So now, as the code loops through the rows, it will update the ticker volume for the volume of each row for that ticker as long as the row matches that ticker. **TickersVolume(tickersindex) = TickersVolume(tickersindex)+ Cells(i,6).Value**

While looping through the rows, the code will update the starting price for each ticker as long as the ticker matches that of the previous row.

Lastly, if the the next ticker row in the loop does not match the current ticker, then two things will happen. First, the code will log the endprice for that ticker in the end price row. Second, the Tickerindex will go up by one: effectively advancing to the next ticker so that the code can continue to evaluate for the next stock.

Lastly, we have modified the outputs to display our outputs for each iterator so that we can visualize all of our analysis. **Example: Cells(4 + i, 2).Value = tickersVolume(i)**
