# Sales Commission Report Automation with Python and Excel

## Soft and Hardskills used:
Python Programming  
Data Analysis  
Excel Skills  
Pivot Tables  
Date Handling  
Financial Analysis  
Automation  
Problem-Solving  
Attention to Detail  
Analytical Thinking  
Communication  
Time Management  
Business Acumen  
## Problem Statement:
The manual calculation of sales commissions is currently a time-intensive task for our financial analysts, requiring significant monthly effort. Errors in these calculations have led to financial losses for the company.

## Introduction:
In the dynamic world of SaaS businesses, accurately tracking and calculating sales commissions can be a complex and time-consuming task. To streamline this process, I developed a Python script that automates the calculation of sales commissions from an Excel file containing client payment data. This project is a great example of how Python can be used to improve Data Analysts workflow. 

## Project Overview
The project begins with an Excel file downloaded from a proprietary system, containing a time series client payments. The dataset includes the following columns:

* Client Code
* Client's Username
* Seller's Name
* Subscription Plan
* Payment Date
* Payment Value
 

## Business Rules
#### The commission structure is defined as follows:
Salespersons earn a commission fee of 10% on every payment made by their clients within the first three months of the client's relationship with the company.
The start of the client relationship is marked by the first payment.
The commission fee is paid in the month following the client's payment.
#### Example Scenario
To illustrate, let's consider a client who subscribes to the software in March 2024 and makes a USD100 payment for one month of use. The salesperson responsible for this client will earn a USD10 commission (10% of 100) for this payment. If the client continues to make payments in April and May 2024, the salesperson will earn commissions for these payments as well, provided they fall within the first three months of the client relationship."

## Implementation Details
The script performs the following steps to calculate the commissions:

### Step 1: Data Loading:

Load the Excel file into a Pandas DataFrame for easy manipulation.

### Step 2: Data Transformation:

Several transformations are applied to clean and structure the data:

Add Seller's Name: A new column is created to store the seller's name extracted from the first row.  
Add Evaluation Period: A new column is created to store the evaluation period extracted from the fourth row.  
Set Column Names: The sixth row is used as the column names, and the preceding rows are dropped.  
Drop Useless Column: The column 'VALOR COMISSÃO' is removed as it is not needed for the analysis.  
Rename Columns: The second column is renamed to 'USERNAME'.  
Reset Index: The index is reset and the index name is removed.  
Convert Date Format: The 'DATA PAGTO' column is converted to a 'YYYY/MM' format.

### Step 3: Data Aggregation:
A pivot table is created to aggregate the payment data by client and payment date.

### Step 4: Calculate Number of Payemnts used in Business Rule:
New columns are added to count the total number of payments and the number of payments in the last three months.

Total Payments: Counts the number of months each client has made a payment.
Payments in Last Three Months: Calculates the number of payments made in the last three months.

### Step 5: Determine Commission Eligibility:

A new column is added to indicate whether a client is eligible for a commission based on the defined rules.

### Step 6: Calculate Due Commission Value:
### Step 7: Export the Analysis to a new Excel Sheet for reporting




## Conclusion
This project exemplifyes how to leverage Python for data manipulation and business logic implementation, automating a crucial financial process for a SaaS company. By efficiently calculating sales commissions, the script not only saves time but also ensures accuracy and consistency in commission payouts. This kind of project underscores technical skills and business operations knowledge.

Feel free to reach out if you have any questions or would like to see a demonstration of the project in action.
