# VBA-challenge

## Assignment Description

### Data

This assignement involved creating a VBA script to loop through stock data and deliver a summary per "Ticker".

In the summary table I had to deliver each individual "Ticker" along with:
  - The yearly change (difference between year closing price and opening price)
  - Percent Change (Perecent difference between the year closing and opening price)
  - Total Stock Volume (Accumulation of the volume of stock across a year for each ticker)

Once the summary table had been populated, I then had to find some additonal statistics:
  - Greatest % Increase (Which ticker experienced the greatest % increase)
  - Greatest % Decrease (Which ticker experienced the greatest % decrease)
  - Greatest Total Volume (Which ticker has the greatest total volume)

### Formatting

For ease of viewing; formatting was applied to the Yearly Change and Percent Change columns:
  - Positive change cells coloured Green
  - Negative change cells coloured Red

And number formatting applied:
  - Yearly change formatted as a currency in USD ($)
  - Percent change formatted as perecentage to 2 decimal places (0.00%)

I also researched and added some additonal formatting to a range of cells:
  - Headers for the Summary table Bold and Horizontal Central Alignment
  - Headers for the Summary table also contained the year which was retrieved from the sheet name
  - Headers for the Greatest % Increase/Decrease and Total Volume Bold and Horiztonal Central Alignment
  - Total Stock Volume formatted as comma seperated to 0 decimal places

## References
Example Data supplied by: edX Boot Camps LLC

Percentage change formula checked on: Investopedia (https://www.investopedia.com/terms/p/percentage-change.asp#:~:text=How%20Do%20I%20Calculate%20Percentage,multiply%20that%20number%20by%20100.)

Additional formatting: Microsoft VBA Documnetation (https://learn.microsoft.com/en-us/office/vba/api/overview/)
