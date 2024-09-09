# VBA-Challenge

  working on Multiple_year_stock_data file

## instruction
step 1:  we need to assign the variable .
step 2:we need to find the opening and closing date of that Quarter dynamicaly.
in this case :openingdate = Range("B2").Value(any thing on value B2 is the opening date)
closingdate = Range("B2").End(xlDown) to find the date of the end quarter(any thing at the end of range("B") is the closing date)
based on the opening and closing date we can find the opening and closing price.

# how to find the Value of ticker?

step 1: we have to determine the data set of the ticker and loop through the range that associated to the "ticker.
the range for the ticker is range("A2":A96001). 

based on the opening date we can find the value of ticker and put on range("I")  if the condition met

# how to find Quarterly Change.

step 1: we have to find the price of the opening date,the opening date is the value in range"B" and it is the date when the market was  open
we can find the opening price if the condition is met based on the opening date.
step 2: finding closing Price.the closing price is the value in range "f".it is the price when the market is closed.so,we can find the closing price based on the closing date if the condition is met.

step 2:the Quarterly change is closing price-opening price.


# how to find the value of Percent Change

it is the value of quarterly change divided by opening price *100
since we format our range:which is range("K")as "percent" we can only calculate Quarterly change by opening price to get percent change

# how to get the Total stock Volume
step1: we need to determine the range of the total stock volume,which is range ("G")
the total stock volume is the sum value of  from the opening date to closing date.
we need to create variable to store the sum of the total stock volume based on the condition met.
StockVolume=range("G" & I+1).value.i is the counter when looping inside the row.start from one.the header index start 0 so we skip the header count from 1.
step 2:once the condition met: StockVolume=stockVolume + range("G" & I+1).value. which hold the sum of the value until the condition is not met.

# how to find the % increase,decrease and the max value
step 1: we need to assign the variable.we need to "dim" to store the value as variant or longlong
step 2: we can find the max value or min value of the range using Worksheetfunction.
step 3:based on the condition we can find the ticker to the associated value,by looping.

# how to rund the each Quarter.
step 1: we can create the module associated to that quarter.
step 2: inside subroutine we need to select the associated worksheet.for example.sheets("Q1").select.
step 3.we can run the script from the module one we initialy created. 
"VBAProject.Module1.StockAnalysis"

# screen Shot of the Result:
![alt text](<Screenshot 2024-09-08 211023.jpg>)

# formating.
it is formated based on the value on the data set which is if it is greater than 0 or less than 0.
if it is greater than o fill in with green and if it is less than 0 fill in with red.


# Note:
if run the script all the value will show on the first worksheet which is sheets("Q1").it will clear the existing data and it will start running on requested sheet.
buttons are created to run accordingly.