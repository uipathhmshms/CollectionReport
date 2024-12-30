# Collection Report Executer

## Project Overview
This UiPath automation project processes data received from a queue, <br>
generates an Excel and PDF report,  <br>
applies custom formatting, <br>
and sends the reports via email using Microsoft Office 365. 

### Quick Overview
![Sample Data Structure](/Documentation/dataFlow1.png)
![Sample Data Structure](/Documentation/dataFlow2.png)
*Quick overview of the process steps*

## Features
- Deserialize JSON Data: Converts specific content from the queue item into a DataTable format.
- Dynamic File Naming: Generates a unique file name for the report using project manager ID and the current timestamp.
- Data Filtering: Filters the DataTable to retain only relevant columns for the report.
- Excel Report Creation: Saves the filtered data into an Excel file with custom formatting.
- PDF Export: Converts the Excel file into a PDF document.
- Email Integration: Sends the Excel and PDF reports via email to the recipient specified in the queue item.

## Workflow Steps
- Convert JSON to DataTable:
- Deserializes JSON data from the queue item's SpecificContent("ReportData").
- Outputs the data as a DataTable (dt_data).
- Generate Report File Name:
Constructs a dynamic file name using the project manager's ID, report type, and current timestamp.
- Retains specific columns required for the report while removing unnecessary columns.
- Create and Format Excel Report:
- Writes the DataTable to an Excel file in the TempReport directory.
- Invokes VBA scripts for:
	1. Custom formatting (FormatTable).
	2. Deleting the default sheet (DeleteSheet1).
	3. Applying additional styles (AddStyleToSheet).
- Save Report as PDF
- Exports the formatted Excel file into a PDF document in the TempReport directory.
- Send Email:
	Attaches the Excel and PDF files to an email.
	Sends the email using the Microsoft Office 365 activity, with the recipient retrieved from the queue item.
	