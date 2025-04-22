# 🔄⚙️ ETL Process Automation Using Power Query 📊
This project is about Automating the ETL process using Power Query in Excel.

## Power Query's Purpose
 - Automate importing and cleaning data
 - Replaces need for VBA Programming for Automation
 - Reduces the time consumption of traditional methods

## About the Project:
In this project, I built an automated ETL (Extract, Transform, Load) process using Power Query in Excel to simplify and accelerate sales data analysis. The system imports multiple sales data files from a designated folder, dynamically combines them, and applies a series of transformation steps — such as filtering, merging, cleaning, and reshaping — to prepare the data for reporting.
The transformed data is loaded into the Pivot Table Cache for efficient and interactive analysis. To update the analysis, users only need to add new sales data files into the folder and refresh the queries. Once refreshed, Power Query automatically performs all transformation steps and updates the Pivot Tables accordingly.
This semi-automated solution reduces manual work, improves consistency, and enables timely insights with minimal effort.

## Meta Data
Data source - Imported from 'Sales' Folder
Number of Files Imported - 3 Files
Sales Data Coverage -Years 2017,2018,2019

## 🔍 Extracting Data (Data Importing Process)
- In Excel choose get data from a folder
- choose the folder
- Select 'Combine and transform data'

  Now the three files of sales data were combined into  one file and imported into the Power Query Editor

## 🛠️ Transforming Data (Data Cleaning Process)
This is how messy our data is

### Issues in the data
 1. Header Formatting: Column headers are split across two rows, requiring consolidation.
 2. Combined Attributes: "Shipping Mode" and "Container Type" are stored in a single column, making it difficult to analyze them separately.
 3. Missing Derived Metric: While "Order Quantity", "Unit Sell Price", and "Discount Percent" are available, the "Sales Amount" column is missing and needs to be calculated.
 4. Lack of Shipping Lead Time: There's no column indicating the number of days between the order date and Shipping date, which is critical for logistics analysis.
 5. Inconsistent Naming Convention: The "SalesPerson" names include formal prefixes (e.g., Mr., Mrs.), which should be standardized for better data clarity and consistency.
   
### Data Cleaning Steps
 1. Selected the 'Transform Sample File' in 'Helping Queries' window
 2. Fixed Header Formatting Issue
    - Removed the automatically applied 'Promoted Headers' step from the Applied Steps panel.
    - Transposed the table to separate the two header rows into distinct columns.
    - Merge the first two columns(header components) into a single column
    - Transpose the data back
    - Apply Use first raw as Header
 3. Separated Combined Attributes: Shipping Mode and Container Type
    - Split columns by delimiter
    - Rename the new columns
 4. Calculated Sales Amount
    - Selected the Order Quantity, Unit Sell Price, and Discount Percent columns.
    - Used the Add Column > Standard > Multiply option to create a calculated column.
    - Adjusted the formula to apply the discount correctly:
      (Sales Amount = Order Quantity × Unit Sell Price × (1 - Discount Percent))
    - Rename the column into 'Sale Amount'
    - Round the sale amounts into two decimal points. ( transform --- rounding --- round)
 5. Added Days to Ship Column
    - Select the 'ship date' column first and then 'order date'
    - Used Add Column > Date > Subtract Days to calculate the time taken to ship.
    - Rename the new column 'Days to Ship'
 6. Standardized SalesPerson Names
    - Used Add Column > Column from Examples > From Selection.
    - Provided an example (e.g., just the first name without titles like "Mr", "Mrs") in the first row.
    - Confirmed the suggested pattern across rows and applied the transformation.
 8. Filtered out rows where Order Priority was listed as "Not Specified", improving data quality for priority-based analysis.

To add the cleaning steps to the original data set from the sample data, the last step of the 'data' file has to be removed. This is because the column names are what before the column headers are changed. So the 
'Change type' step has to be removed from the 'data' query. Then delete the 'Source Name' column which indicates the original source of the data before they were merged. Finally, select all the columns and apply 'Detect Data type' to ensure the data types of each column.

## 📤 Load Data
 - Select 'Close & Load To' option
 - Select the 'PivotTable report' option to load the data into the pivot cache
 - Group the columns by ' Order priority' and see how sale Amount changes with the time
 - Add a line chart to see the trend

## Adding New data files
Add the 2020 sales data to the 'sale' folder. Then refresh the queries to grab the new file from the folder, run through all the transformation steps and load into Pivot cache. At the sametime this will automatically update the pivot table and the pivot chart. 

    
       
      
