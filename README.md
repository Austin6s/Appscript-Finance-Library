# Appscript for Google Sheets Financial Tracking System

This code was used in Google Appscript to provide a suite of tools in a library called: ProcessesforMonthlysales.js to help automate the components of a company's financial tracking system in Google Sheets. The financial system requires users to enter sales and expenses data then calculates all the major financial documents: balance sheet, income statement, cash-flow statement. The financial spreadsheets are separated by month and are then aggregated and imported to a yearly financial document. The Appscript code (which is actually written in Javascript) is meant to help automate the integrations. There are also a few helper functions that make user data entry more streamlined and less prone to human error as well as simplify tedious tasks. There is also a script called: Trigger_script.js which have two functions: onEdit & onOpen that are used as triggers for the code using the functions from the ProcessesforMonthlysales library.

## Features

- Data import from monthly spreadsheet into yearly spreadsheet;
- User input collection for data import using HTML & CSS;
- Data reset -> clears data in manual entered cells only, while keeping sheet structure and formulas;
- Currency converter for all accounting data;
- Dependent dropdowns;
- Automated inventory tracking that collects sales data from monthly spreadsheets to update.
