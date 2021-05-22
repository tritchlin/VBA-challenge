# VBA-challenge
VBA Homework 2 for UCSD Data Bootcamp

Script file: UCSD_Data_VBA_Challenge.bas, uploaded 05/21/2021

VBA Script and applied to Excel Dataset, goals for project as follows:

    Output aggregated data for each tab:
    - Ticker Symbol
    - Yearly change from opening price of each given year, vs closing price at end of same year
    - Percent change with same data as Yearly change
    - Total stock volume for each ticker
    
    Additional:
    - Conditional formatting for Yearly Change column, with positive changes highlighted green and negative changes highlighted red
    - Create additional mini-table that shows, per worksheet:
      - Ticker with greatest % increase
      - Ticker with greatest % decrease
      - Ticker with greatest total volume
      
    VBA Script made to be as dynamic as possible to account for addition/removal of any data to existing data set

Commands and Syntax Used:
- Looping across all worksheets in a workbook
- Setting variables and variable types
- Entering data into cells
- Nested For Loops
- "Last Column" and "Last Row" identifiers
- Conditional formatting with colors
- Cell value formatting (percentage)
- Evaluating Max and Min of a given column
- Comparing data from different rows and columns
- Excel Worksheet Functions in VBA
- Autofitting columns to data
