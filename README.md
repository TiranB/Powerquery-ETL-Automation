# Power Query Automation
This project is about Automating the 'data importing and cleaning process' using Power Query in Excel.

## Power Query's Purpose
 - Automate importing and cleaning data
 - Replaces need for VBA Programming for Automation
 - Reduces the time consumption of traditional methods

## About the Project:


## Meta Data
Import From - 'Sales' Folder
Number of Files - 3
Year of Sales - 2017,2018,2019

## Data Importing Process
- In Excel choose get data from a folder
- choose the folder
- Select 'Combine and transform data'

  Now the three files of sales data were combined into  one file and imported into the Power Query Editor

## Data Cleaning Process
This is how messy our data is

### Issues in the data
 1. Headers are spread over two rows
 2. Shipping Mode and the Container type stuffed into one column
 3. 'Order Quantity', 'Unit Sell Price', 'Discount Percent' columns are available but a column for 'Sales Amount' is missing
 4. Missing a column to find out how many days it takes to ship the product after ordered
 5. SalesPerson data is too much formal ( having Mr, Mrs infront of the names)

### Data Cleaning Steps
 1. Selected the 'Transform Sample File' in 'Helping Queries' window
 2. Fixed the 1st issue in the dataset (Fix the Headers)
    - In the 'Applied steps' window remove the 'promoted Headers' step which was done automatically
    - Transpose the data to separate the two headers in two columns
    - Merge the first two columns which have the headers into  one column
    - Transpose the data back
    - Apply Use first raw as Header
 3. Fixed the 2nd issue in the dataset (Separation of shipping mode and container type)
    - Split columns by delimiter
    - Rename the new columns
 4. Fixed the 3rd issue (Added a Sales Amount column)
    - Select the 'Order Quantity', 'Unit Sell Price', 'Discount Percent' columns
    - In 'Add Columns' tab select the 'Standard' and then 'Multiply' options
    - Modify the formula to ' 1 - discount percent' to get the correct sales amount
    - Rename the column into 'Sale Amount'
    - Round the sale amounts into two decimal points. ( transform --- rounding --- round)
 5. Fixed the 4th issue (Added a column to show the day difference between orderd and ship date)
    - Select the 'ship date' column first and then 'order date'
    - In 'Add column' tab goto Date and select 'substract days'
    - Rename the new column 'Days to Ship'
 7. Fixed the 5th issue (Keep the first name of the sales person)
    - In 'Add Column' tab goto 'Columns from Example' then 'from selection'
    - In new column in the first raw, give an example of reqired format
    - check other raws seem to be correct and then apply the rule
       
      
