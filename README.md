# Weekly Report Project (VBA)
## 🗺️ Context
A company tasked me with automating the weekly creation of reports on employee performance. Additionally, they wanted to automate the process of sending these reports to each employee individually. The template for the reports was already created in an Excel file, and the KPIs were predefined. They had been manually exporting and sending these reports for several weeks, which was highly time-consuming due to the number of employees.

The template already contains all the necessary formulas and connections, so every time you select an employee's name, the correct information is displayed. There is also another sheet with the names of all employees and their email addresses.

The manager emphasized the need to have each report exported as a PDF on his computer before emailing them, allowing for a review.

## 🎯 Objectives
The main goal of the project is to create two VBA scripts:
- One to export a report for each employee in the company with their data.
- One to send each report to the corresponding employee by email.

## ✅ Solution steps
Steps for the first script:
- Count the number of employees.
- Create a loop that repeats for each employee, performing the following actions:
  - Select the name of each employee in the template to update the results of all KPI formulas.
  - Export the results to a PDF, naming the file after the employee, to a specified filepath.

That script can be found [HERE](https://github.com/Pablojox/weekly-report-vba/blob/main/export-weekly-report.bas)

Steps for the second script:
- Count the number of employees.
- Create a loop that repeats for each employee, performing the following actions:
  - Locate the report for each employee using a filepath on the computer.
  - Retrieve the email address of that employee from the Excel file.
  - Send the report to that employee via email.
 
That script can be found [HERE](https://github.com/Pablojox/weekly-report-vba/blob/main/send-email.bas)
