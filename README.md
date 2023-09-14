# monthly-lost-contarcts

The purpose of this script is to extract data from individual regions of the monthly report of construction machinery sold across the country and create individual files for each regional sales representative.

![data screenshot](../assets/data_sheet_screenshot.png)

The report we receive only includes voivodeships and the number of machines in a given category that have been delivered.
Based on these individual files, salespeople report what machine was delivered and where (model, brand, customer - as far as their market knowledge allows), thanks to which we obtain a more accurate picture of the market. 

To facilitate their work, the script creates new lines in place of those where there is more than one machine of a given class.
