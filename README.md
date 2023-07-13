# VBA-Challenge
Multiple_Year_Stock_Data
In this VBA challenge, raw data that has already been formatted is provided for analysis as shown below.

<img width="397" alt="Screenshot 2023-07-13 at 12 41 10" src="https://github.com/rmatsui6289/VBA-Challenge/assets/137141385/22c5da0e-75f0-4da4-b54c-d4e88e82f9d2">

Given this data, our task was to create two summary tables, the first of which ranges from column I to column L. 
Looping through the data, we obtain the opening and closing for each ticker value at the beginning and end of a given year.
Using these two values, we calculate the 'yearly change' by subtracting the opening price from the closing price. 
Following this, the 'percent change' is calculated by taking the 'yearly change' value and dividing it by the opening price. 
Finally, the 'total stock volume' is calculated by summing up all the volume values for a given ticker. The second summary table loops through the first summary table looking for the 'greatest percent increase,' 'greatest percent decrease,' and 'greatest total volume' and displays the ticker symbol for each. 

<img width="1038" alt="Screenshot 2023-07-13 at 10 54 18" src="https://github.com/rmatsui6289/VBA-Challenge/assets/137141385/141ccfed-3fe2-464d-b4cc-44ab6561a991">


The macro runs through all the worksheets in a given workbook.
