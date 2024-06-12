# VBA Challenge: Stock Data Analysis

## Overview

This VBA script is designed to analyze stock data, specifically calculating yearly change, percentage change, and total volume for each stock ticker. The challenge tested our understanding of proper syntax usage, index definition, and looping in VBA. Through this assignment, we gained insights into effectively manipulating large datasets with VBA scripting. This challenge was especially enriching because it provided an unique understanding of excel functions, and executing them under-the-hood for more advanced capabilities. The logic used in this code has laid some of the ground work for pythonic control flow.

## Technologies Utilized

Microsoft Excel: For managing and visualizing the stock data. VBA (Visual Basic for Applications): For automating data analysis within Excel.

## Major Functions and Concepts

### Key Variables and Data Types

Double: Used to store values with decimal points, such as percentages and changes in stock prices.
Long: Used for loop counters and storing row counts to handle large datasets efficiently.

total: Stores the cumulative total volume for each stock ticker.
change: Calculates the yearly change in stock price.
percentChange: Calculates the percentage change in stock price.
rowCount: Determines the number of rows in the dataset.
start: Tracks the starting row for each stock ticker.

### Core Logic

**Initializing and Formatting:**

Set up column titles for "Ticker", "Yearly Change", "Percent Change", and "Total Volume".

**Looping Through Data:**

A nested loop iterated through each row of the dataset, checking for changes in stock ticker symbols to calculate required values.

**Conditions and Calculations:**

Conditions ensured correct calculation of yearly change and percentage change.
Handled cases where starting stock prices are zero to avoid division errors.

**Output Results:**

Populated the results in the respective columns for each stock ticker.
Applied designated formatting for numerical values and percentages.

**Conclusion**

This VBA script effectively processed stock data, providing a clear overview of yearly changes, percentage changes, and total volumes for each stock ticker. Through the use of loops, conditions, and appropriate data types, we learned to efficiently handle large datasets, demonstrating the accuracy and agility of VBA in data analysis tasks. 






