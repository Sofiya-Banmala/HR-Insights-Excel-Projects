# HR-Insights-Excel-Projects

## Project Overview  
This project analyzes **HR data** using Excel to provide key insights into employee salaries, department-wise distributions, gender-based earnings, and workforce trends. The Excel file includes **various formulas and functions** used to extract meaningful business insights.  

# 2. **Essential Excel Functions for Data Analysis**

This is an **Excel file** where I practiced some **essential Excel functions and formulas**. It includes common formulas used for basic calculations, lookup and reference, conditional counting, summing, filtering, and data organization and visualization required for data analysis tasks.  

**Download the file here:** [Download Excel File](https://github.com/Sofiya-Banmala/Excel-Project/blob/main/Essential_Excel_Functions_for_Data_Analysis_Practise.xlsx)  

Here are some of the business questions that i solved using these functions and formuas:

![image alt](https://github.com/Sofiya-Banmala/Excel-Project/blob/main/excel%20func.JPG?raw=true)

This business questions helps to find the techniques covered that enables efficient calculation of salaries, headcounts, and averages, while also supporting dynamic data filtering, sorting, and gender-based analysis. They facilitate the generation of reports, advanced lookups, error handling, and statistical or time-based analysis, providing valuable insights for decision-making.

Below, we describe each business question in detail, along with the functions used to achieve the desired results.

**1. Total Salary and Headcount by Department**

### Description:
This analysis calculates the total number of employees (HeadCount) and their total salary across different departments. Additionally, it separately calculates the headcount and total salary of permanent employees within each department.

### Table Structure:

![image alt](https://github.com/Sofiya-Banmala/Excel-Project/blob/main/1an%202.JPG?raw=true)


### Functions Used:

- `=COUNTIF(staff[Department], A4)`  
  - This function counts the number of employees in a given department.

- `=SUMIF(staff[Department], A5, staff[Salary])`  
  - This function calculates the total salary of all employees in a given department.

- `=COUNTIFS(staff[Department], A5, staff[Employee type], "Permanent")`  
  - This function counts the number of permanent employees in a given department.

- `=SUMIFS(staff[Salary], staff[Department], A4, staff[Employee type], "Permanent")`  
  - This function calculates the total salary of permanent employees in a given department.

**2. Average Salary by Department**

### Description:
This analysis calculates the average salary of employees in each department. The average salary provides insights into compensation distribution across departments.

### Function Used:

- `=AVERAGEIF(staff[Department], A4, staff[Salary])`  
  - This function calculates the average salary of employees in each department.

**3. All Employees with More Than $100K Salary**

### Description:
This analysis identifies employees earning more than $100,000 annually across different departments. By filtering employees based on salary, we can gain insights into high-income earners within the organization. This data is useful for workforce planning, salary benchmarking, and identifying top-earning employees across locations.

### Table Structure:

![image alt](https://github.com/Sofiya-Banmala/Excel-Project/blob/main/3.JPG?raw=true)

### Functions Used:
- `=FILTER(staff, staff[Salary] > D2)`: This function is used to filter and display all employees whose salary exceeds $100,000. The `FILTER` function dynamically retrieves records based on the salary condition.
- `=staff[#Headers]`: This function is used to reference the column headers dynamically, ensuring that the extracted data includes appropriate labels.

**4. All Female Employees with More Than $100K Salary**

### Description:
This analysis identifies all female employees who earn more than $100,000 annually. The purpose of this analysis is to examine gender-based salary distribution and identify high-earning female employees within the organization.

### Table Structure:

![image alt](https://github.com/Sofiya-Banmala/Excel-Project/blob/main/4.JPG?raw=true)

### Functions Used:
- `=CHOOSECOLS(FILTER(staff, staff[Gender] = "Female", staff[Salary] > 100000), 1,2,3,4,5,6)`:  
  - `FILTER(staff, staff[Gender] = "Female", staff[Salary] > 100000)`: Filters employees based on gender (Female) and salary greater than $100,000.
  - `CHOOSECOLS(..., 1,2,3,4,5,6)`: Selects specific columns (Emp ID, First Name, Last Name, Gender, Department, Salary) from the filtered data.
 
**5. All Female Employees with More Than $100K Salary Who Joined in 2020 or After**

### Description:
This analysis identifies all female employees who earn more than $100,000 annually and joined the organization in 2020 or later. The purpose of this analysis is to understand the impact of recent hires with high salaries and assess trends related to high-earning female employees in the company.

### Table Structure:

![image alt](https://github.com/Sofiya-Banmala/Excel-Project/blob/main/5.JPG?raw=true)

### Functions Used:
- `=FILTER(staff, (staff[Gender]="Female") * (staff[Salary]>100000) * (YEAR(staff[Start Date])>=2020))`:
  - `FILTER(staff, ...)`: Filters the dataset based on the specified conditions.
  - `(staff[Gender]="Female")`: Filters for female employees.
  - `(staff[Salary]>100000)`: Filters for employees with a salary greater than $100,000.
  - `(YEAR(staff[Start Date])>=2020)`: Filters for employees who joined in 2020 or after.

**6 & 7. Salary Analysis: Lowest, Highest, and Top 5 Salary Values (Overall and by Gender)**

### Description:
This analysis identifies the lowest, highest, and top 5 salary values within the organization as well as by gender (Male and Female). The goal is to assess salary distribution across the workforce, highlight trends within different genders, and identify the range of salaries, including the highest earners. This will help evaluate overall salary equity and gender-based disparities, if any, in compensation within the company.

### Table Structure:

![image alt](https://github.com/Sofiya-Banmala/Excel-Project/blob/main/67.JPG?raw=true)

### Functions Used:
- `=MIN(staff[Salary])`: Identifies the lowest salary across all employees.
- `=MAX(staff[Salary])`: Identifies the highest salary across all employees.
- `=LARGE(staff[Salary], E6)`: Returns the top 5 highest salary values from the dataset based on the rank number (from 1 for the highest, 2 for second highest, etc.).
- `=MINIFS(staff[Salary], staff[Gender], "Male")`: Identifies the lowest salary for male employees.
- `=MAXIFS(staff[Salary], staff[Gender], "Male")`: Identifies the highest salary for male employees.
- `=MINIFS(staff[Salary], staff[Gender], "Female")`: Identifies the lowest salary for female employees.
- `=MAXIFS(staff[Salary], staff[Gender], "Female")`: Identifies the highest salary for female employees.
- `=TAKE(SORT(staff[Salary], , -1), 5)`: Returns the top 5 highest salaries after sorting the dataset in descending order.

**8 & 9. Department List Analysis: All Departments and Comma-Separated List**

### Description:
This analysis generates a list of all unique departments within the organization, as well as a comma-separated list of these departments. The goal is to provide an overview of the organizational structure by highlighting the various departments and creating a consolidated, easily-readable list for reporting purposes.

### Table Structure:

![image alt](https://github.com/Sofiya-Banmala/Excel-Project/blob/main/8.JPG?raw=true)

![image alt](https://github.com/Sofiya-Banmala/Excel-Project/blob/main/9.JPG?raw=true)

### Functions Used:
- `=UNIQUE(staff[Department])`: Extracts a unique list of all departments from the `Department` column, removing duplicates.
- `=TEXTJOIN(", ", TRUE, UNIQUE(staff[Department]))`: Combines all unique departments into a single cell, with each department name separated by a comma and a space. The `TRUE` argument ensures that any empty cells are ignored.

**10. Employee Details Lookup**

### Description:
This analysis provides a lookup of employee details based on specific identifiers such as Employee ID or Last Name. The purpose of this analysis is to retrieve and display information about a particular employee, including their first name, last name, department, and salary, based on a given search criterion. This can be useful for HR departments, payroll teams, or any role requiring quick access to employee-specific data.

### Table Structure:

![image alt](https://github.com/Sofiya-Banmala/Excel-Project/blob/main/10.JPG?raw=true)

### Functions Used:
- `=VLOOKUP(B3, staff, 2, 0)`: This function looks up the Employee ID (provided in cell `B3`) in the employee data (`staff`), retrieving the corresponding value from the second column (First Name in this case). The `0` argument ensures an exact match.
- `=INDEX(staff[Emp ID], B15)`: This function uses the index to retrieve the Employee ID from the dataset based on the row number provided in `B15`.
- `=MATCH(B14, staff[Last Name], 0)`: This function searches for the Last Name (given in `B14`) in the `staff` table and returns the row number where the last name is found. The `0` ensures that only an exact match is returned.

**11. Employee Details Lookup**

### Description:
This analysis provides a lookup for employee details based on the Employee ID. The `XLOOKUP` function is used to search for a specific employee's ID in the dataset and return corresponding employee details, such as their first name, last name, department, and salary.

### Table Structure:

![image alt](https://github.com/Sofiya-Banmala/Excel-Project/blob/main/11.JPG?raw=true)

### Formula Used:
- `=XLOOKUP(C3, staff[Emp ID], staff[First Name], "NA")`: This function looks up the Employee ID provided in `C3` in the dataset `staff`, returning the corresponding first name. If the ID is not found, it returns "NA".


**12. Complex Formula: Highest Salary Person**

### Description:
This analysis identifies the person with the highest salary in the dataset. It uses the `MAX` function to find the highest salary and `XLOOKUP` and `FILTER` to return the name(s) of the employee(s) earning that salary.

### Table Structure:

![image alt](https://github.com/Sofiya-Banmala/Excel-Project/blob/main/12.JPG?raw=true)

### Formula Used:
- `=MAX(staff[Salary])`: Finds the maximum salary in the `staff` dataset.
- `=XLOOKUP(C3, staff[Salary], staff[First Name] & " " & staff[Last Name], "NA", 0)`: Looks up the highest salary and returns the corresponding employee's full name.
- `=FILTER(staff[First Name], staff[Salary] >= MAX(staff[Salary]))`: Filters and lists the first names of employees who have the highest salary.

**13. Complex Formula: All Employees Joined in March**

### Description:
This formula filters and displays all employees who joined in the month of March. It uses the `FILTER` function in combination with `MONTH` to isolate employees who have a start date in March.

### Table Structure:

![image alt](https://github.com/Sofiya-Banmala/Excel-Project/blob/main/13.JPG?raw=true)

### Formula Used:
- `=CHOOSECOLS(FILTER(staff, MONTH(staff[Start Date]) = 3), 1, 2, 3)`: Filters the employees who started in March and returns the selected columns (Emp ID, First Name, Last Name).


**14. Complex Formula: Female Employees with Monday Start**

### Description:
This formula filters female employees who started their employment on a Monday. It uses the `FILTER` function in combination with `TEXT` to check the start date and gender.

### Table Structure:

![image alt](https://github.com/Sofiya-Banmala/Excel-Project/blob/main/14.JPG?raw=true)

### Formula Used:
- `=FILTER(staff, (staff[Gender] = "Female") * (TEXT(staff[Start Date], "dddd") = "Monday"))`: Filters female employees whose start date falls on a Monday.


**15. Complex Formula: Department Report of Headcounts, Salaries, and % Diff from Overall Average**

### Description:
This analysis calculates key metrics for each department, including headcount, average salary, percentage difference from the overall average salary, highest salary, median salary, and female ratio. The `UNIQUE`, `COUNTIF`, `AVERAGEIF`, `MAXIFS`, `MEDIAN`, and `COUNTIFS` functions are used to generate the required data.

### Table Structure:
 
![image alt](https://github.com/Sofiya-Banmala/Excel-Project/blob/main/1516.JPG?raw=true) |

### Formula Used:
- `=UNIQUE(staff[Department])`: Lists all unique departments.
- `=COUNTIF(staff[Department], A6)`: Counts the number of employees in the specified department.
- `=AVERAGEIF(staff[Department], A6, staff[Salary])`: Calculates the average salary for a department.
- `=(C3 - C6) / C3`: Calculates the percentage difference from the overall average salary.
- `=MAXIFS(staff[Salary], staff[Department], A7)`: Finds the highest salary in a given department.
- `=MEDIAN(FILTER(staff[Salary], staff[Department] = A6))`: Calculates the median salary for a specific department.
- `=COUNTIFS(staff[Department], A6, staff[Gender], "Female") / COUNTIF(staff[Department], A6)`: Calculates the female ratio in a department.


**16. Calculate Median Salary and Female Ratio**

### Description:
This analysis calculates the median salary and the female ratio across the entire staff or by department. It uses the `MEDIAN` function and `COUNTIFS` to analyze salary and gender data.

### Formula Used:
- `=MEDIAN(staff[Salary])`: Calculates the median salary for all employees.
- `=COUNTIFS(staff[Gender], "Female") / COUNTA(staff[Gender])`: Calculates the ratio of female employees in the entire dataset.

**Functions Used in the File**

- **COUNTA** – Counts all non-empty rows in a given range.  
- **COUNT** – Counts only numeric values in a range.  
- **COUNTIF** – Counts how many times a specific value appears in a range.  
- **SUMIF** – Adds up values based on a condition.  
- **COUNTIFS** – Counts values based on multiple conditions.  
- **SUMIFS** – Adds up values based on multiple conditions.  
- **AVERAGEIF** – Calculates the average of values that meet a condition.  
- **MINIF / MAXIF** – Finds the smallest or largest value based on a condition.  
- **MINIFS / MAXIFS** – Finds the smallest or largest value based on multiple conditions.  
- **FILTER** – Filters data based on a condition.  
- **SORT & TAKE** – Sorts data and extracts a subset of rows.  
- **VLOOKUP** – Searches for a value in a column and returns a result from another column.  
- **XLOOKUP** – A more flexible version of VLOOKUP.  
- **CHOOSECOLS** – Selects specific columns from a dataset.  
- **MEDIAN** – Finds the middle value of a dataset.  

This is useful for anyone learning **Data Analysis in Excel**.

# 3. **Excel Case Study Questions Solving**

This file [Download File](https://github.com/Sofiya-Banmala/Excel-Project/blob/main/Excel_Interview_Practice.xlsx) contains multiple case studies and challenges designed to practice Excel formulas and data analysis techniques using sample datasets.

## Datasets Involved:

![image alt](https://github.com/Sofiya-Banmala/Excel-Project/blob/main/sales.JPG?raw=true)

![image alt](https://github.com/Sofiya-Banmala/Excel-Project/blob/main/churn.JPG?raw=true)

![image alt](https://github.com/Sofiya-Banmala/Excel-Project/blob/main/emp.JPG?raw=true)

### **Case Study 1: Sales Data Analysis**

![image alt](https://github.com/Sofiya-Banmala/Excel-Project/blob/main/salesdata.JPG?raw=true)
     
### **Case Study 2: Customer Churn Analysis**

![image alt](https://github.com/Sofiya-Banmala/Excel-Project/blob/main/custdata.JPG?raw=true)

### **Case Study 3: Employee Performance Dashboard**

![image alt](https://github.com/Sofiya-Banmala/Excel-Project/blob/main/empdata.JPG?raw=true)

### **Excel Formula Challenge**

![image alt](https://github.com/Sofiya-Banmala/Excel-Project/blob/main/excchallenge.JPG?raw=true)

### **Data Cleaning Challenge**

![image alt](https://github.com/Sofiya-Banmala/Excel-Project/blob/main/dataclean.JPG?raw=true)

## Key Learnings  

### 1. **Data Analysis & Business Insights**  
- Understand how to analyze sales, customer churn, and employee performance data.  
- Use **Pivot Tables**, **SUMIFS()**, **COUNTIFS()**, and other Excel formulas for insights.  
- Identify trends in revenue, customer retention, and workforce efficiency.  

### 2. **Excel Formula Mastery**  
- Learn advanced Excel functions such as **LARGE()**, **INDEX()**, and **IF()** for decision-making.  
- Use **COUNTIFS()** and **AVERAGEIFS()** to filter and summarize data.  
- Apply **DATEDIF()** to track inactivity in customer churn analysis.  

### 3. **Data Visualization & Reporting**  
- Create **bar charts, pie charts, and conditional formatting** for insights.  
- Use **Pivot Tables and Slicers** to build dynamic dashboards.  
- Highlight key metrics such as top sales transactions, active customers, and high-performing employees.  

### 4. **Data Cleaning & Standardization**  
- Use **PROPER(), TRIM(), and SUBSTITUTE()** to clean and standardize data.  
- Remove inconsistencies in text formatting and phone numbers.  
- Ensure clean datasets for accurate reporting and analysis.  

### 5. **Real-World Business Applications**  
- **Sales Analysis**: Identify high-revenue regions and top-performing sales reps.  
- **Customer Retention**: Predict churn and improve customer engagement strategies.  
- **Workforce Efficiency**: Measure employee performance and optimize resource allocation.  

## Why This is Important?  
This project is valuable for aspiring **Data Analysts, Business Analysts, and Excel Experts**. It enhances skills in:  
**Data-Driven Decision Making**  
**Excel Automation & Analysis**  
**Visualization & Dashboarding**  
