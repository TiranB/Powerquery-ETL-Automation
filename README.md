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
    - In 'Add column' tab go to Date and select 'substract days'
    - Rename the new column 'Days to Ship'
 7. Fixed the 5th issue (Keep the first name of the sales person)
    - In 'Add Column' tab go to 'Columns from Example' then 'from selection'
    - In new column in the first raw, give an example of reqired format
    - check other raws seem to be correct and then apply the rule
 8. Removed 'Not Specified' data in OrderPriority column

To add the cleaning steps to the original data set from the sample data, the last step of the 'data' file has to be removed. This is because the column names are what before the column headers are changed. So the 
'Change type' step has to be removed from the 'data' query. Then delete the 'Source Name' column which indicates the original source of the data before they were merged. Finally, select all the columns and apply 'Detect Data type' to ensure the data types of each column.

## Load Data
 - Select 'Close & Load To' option
 - Select the 'PivotTable report' option to load the data into the pivot cache
 - Group the columns by ' Order priority' and see how sale Amount changes with the time
 - Add a line chart to see the trend

## Adding New data files
Add the 2020 sales data to the 'sale' folder. Then refresh the queries to grab the new file from the folder, run through all the transformation steps and load into Pivot cache. At the sametime this will automatically update the pivot table and the pivot chart. 

    
       
      
