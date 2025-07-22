# Excel Formula Reference Guide

This document provides a quick overview and explanation of various Excel formulas used in data analysis. Each formula includes a description, example, and explanation of what it does in the context of a dataset.

# Logical Functions

# IF  
**Description:**  
Tests a condition and returns one value if TRUE and another if FALSE.  
**Example Formula:**  
`=IF(D2 > 30, "Old", "Young")`  
**What it does:**  
Checks if the value in D2 is greater than 30. If yes, returns "Old"; otherwise, returns "Young".

# IFS  
**Description:**  
Tests multiple conditions and returns the value for the first TRUE condition.  
**Example Formula:**  
`=IFS(F2 = "Salesman", "Sales", F2 = "HR", "Fire Immediately", F2 = "Regional Manager", "Give Christmas Bonus")`  
**What it does:**  
Returns "Sales" for Salesman, "Fire Immediately" for HR, and "Give Christmas Bonus" for Regional Manager.


# Math Functions

# MAX  
**Description:**  
Returns the highest numeric value in a range.  
**Example:** `65000`  
**Formula:** `=MAX(G2:G10)`  
**What it does:**  
Finds the maximum salary from G2 to G10.

# MIN  
**Description:**  
Returns the lowest numeric value in a range.  
**Example:** `36000`  
**Formula:** `=MIN(G2:G10)`  
**What it does:**  
Finds the minimum salary from G2 to G10.

# SUM  
**Description:**  
Adds up all numeric values in a range.  
**Example:** `437000`  
**Formula:** `=SUM(G2:G10)`  
**What it does:**  
Adds all salary values in G2 to G10.


# SUMIF  
**Description:**  
Adds values that meet a single condition.  
**Example:** `128000`  
**Formula:** `=SUMIF(G2:G10, ">50000")`  
**What it does:**  
Sums salaries greater than 50,000.

# SUMIFS  
**Description:**  
Adds values that meet multiple conditions.  
**Example:** `88000`  
**Formula:** `=SUMIFS(G2:G10, E2:E10, "Female", D2:D10, ">30")`  
**What it does:**  
Sums salaries for females older than 30.


# Counting Functions

# COUNT  
**Description:**  
Counts numeric values in a range.  
**Example:** `9`  
**Formula:** `=COUNT(G2:G10)`  
**What it does:**  
Counts numeric entries in G2 to G10.

# COUNTIF  
**Description:**  
Counts values that meet a condition.  
**Example:** `5`  
**Formula:** `=COUNTIF(G2:G10, ">45000")`  
**What it does:**  
Counts how many salaries are greater than 45,000.

# COUNTIFS  
**Description:**  
Counts values meeting multiple conditions.  
**Example:** `3`  
**Formula:** `=COUNTIFS(A2:A10, ">1005", E2:E10, "Male")`  
**What it does:**  
Counts how many employees have EmployeeID > 1005 and are Male.

# Date Functions

# DAYS  
**Description:**  
Returns the number of days between two dates.  
**Example:** `5056`  
**Formula:** `=DAYS(I2, H2)`  
**What it does:**  
Calculates total days between EndDate (I2) and StartDate (H2).

# NETWORKDAYS  
**Description:**  
Returns working days between two dates (excludes weekends).  
**Example:** `3636`  
**Formula:** `=NETWORKDAYS(H2, I3)`  
**What it does:**  
Calculates working days between H2 and I3, excluding weekends.


# Text Functions

# LEN  
**Description:**  
Returns the number of characters in a text string.  
**Formula:** `=LEN(C2:C10)`  
**What it does:**  
Counts characters in each cell from C2 to C10 (e.g., last names).


# LEFT  
**Description:**  
Returns the first N characters from a text string.  
**Formula:** `=LEFT(B2:B10, 3)`  
**What it does:**  
Extracts first 3 characters from each name in B2 to B10.

# RIGHT  
**Description:**  
Returns the last N characters from a text string.  
**Formula:** `=RIGHT(H2:H10, 4)`  
**What it does:**  
Extracts last 4 digits from dates in H2 to H10.

# TEXT  
**Description:**  
Formats a value using a custom number format.  
**Formula:** `=TEXT(H2:H10, "dd/mm/yyyy")`  
**What it does:**  
Converts dates into day/month/year format.

# TRIM  
**Description:**  
Removes extra spaces from text.  
**Formula:** `=TRIM(C2:C10)`  
**What it does:**  
Cleans text in C2:C10 by removing leading, trailing, and double spaces.

# CONCATENATE  
**Description:**  
Joins two or more text strings into one.  
**Formula:** `=CONCATENATE(B2, " ", C2)`  
**What it does:**  
Combines first and last names into a full name.

# SUBSTITUTE  
**Description:**  
Replaces specific text in a string with new text.  

**Example 1:** `=SUBSTITUTE(H2:H10, "/", "-", 1)`  
Replaces the **first** `/` with `-` in each date.

**Example 2:** `=SUBSTITUTE(H2:H10, "/", "-", 2)`  
Replaces the **second** `/` with `-`.

**Example 3:** `=SUBSTITUTE(H2:H10, "/", "-")`  
Replaces **all** `/` characters with `-`.

# Usage

These formulas are useful for:

- Cleaning and transforming data
- Creating helper columns
- Summarizing key metrics
- Enhancing dashboards and reports


#  Repository Structure

- `Excel_Sales_Analysis.xlsx` – Full Excel file with data, formulas, pivot tables, charts, and dashboard
- `README.md` – Formula documentation and usage guide



Created by Winnie- Ngamau 

