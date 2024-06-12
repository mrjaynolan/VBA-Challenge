# Multiple Year Stoack Data Analysis
Module 2 VBA Challenge assignment

Purpose

This VBA script is designed to analyze multiple years of stock data across various worksheets within an Excel workbook. It performs the following tasks for each worksheet:

Summarizes ticker symbols and calculates quarterly changes.
Computes percent changes and total stock volumes.
Identifies the ticker with the greatest percent increase, greatest percent decrease, and the greatest total volume.
Highlights positive and negative changes in quarterly data.
Outputs summary data for easy reference.
Instructions for Use

Requirements
Microsoft Excel with VBA support.
An Excel workbook with multiple worksheets, each containing stock data in a specific format.
Expected Data Format
Each worksheet should contain the following columns in this order:

Column A: Ticker
Column B: Date
Column C: Open
Column D: High
Column E: Low
Column F: Close
Column G: Volume
Steps to Run the Script
Open the Workbook:

Open the Excel workbook containing the stock data.
Access the VBA Editor:

Press ALT + F11 to open the VBA editor.
Insert a New Module:

In the VBA editor, go to Insert > Module to create a new module.
Copy and Paste the Script:

Copy the provided VBA script and paste it into the new module.
Run the Script:

Close the VBA editor.
Press ALT + F8, select Multiple_year_stock_data, and click Run.
Script Details
Initialization: Initializes variables for storing ticker symbols, totals, and column indexes.
Loop through Worksheets: Loops through each worksheet in the workbook.
Setup Columns: Defines the columns for summary data (Ticker, Quarterly Change, Percent Change, Total Stock Volume).
Data Calculation:
Loops through each ticker symbol to calculate the quarterly change, percent change, and total stock volume.
Color codes the quarterly change based on whether it is positive, negative, or zero.
Summary Calculation:
Identifies the greatest percent increase, greatest percent decrease, and greatest total volume.
Outputs these values along with the corresponding ticker symbols to a summary table.
By following these steps, the script will analyze the stock data across multiple worksheets, perform the necessary calculations, and highlight significant changes and volumes. This provides a clear and concise summary of the stock performance over the given period.
