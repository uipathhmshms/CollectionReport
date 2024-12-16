# <span style="color:#4CAF50;">Collection Report Generator - UiPath Automation</span>

## <span style="color:#2196F3;">Overview</span>

This UiPath automation project is designed to generate **Collection Reports** for each **Project Manager** in the company. The process includes the following key steps:

1. **Data Extraction**: The automation retrieves data from the company database.
2. **Data Filtering**: The data is filtered based on the project managers.
3. **Excel Generation**: A temporary Excel file is created for each Project Manager within a folder (`TempReport`).
4. **Excel Formatting**: A VBA script is inserted to format the generated Excel file.
5. **Conversion to PDF**: The Excel file is converted to a PDF.
6. **Email Notification**: The Excel file and the corresponding PDF are sent to the finance team via email.
7. **File Cleanup**: After the report is sent, the generated files are deleted from the folder to maintain cleanliness.

This process is automated using UiPath, making it faster and more efficient for the finance team.

## <span style="color:#2196F3;">Project Structure</span>

- **Input**: Data is fetched from the company's database.
- **Output**: An Excel report and a corresponding PDF report are sent via email.
- **Temporary Files**: Generated files are stored temporarily in the `TempReport` folder.

### Folder Structure

```plaintext
└── TempReport/             # Temporary folder for generated reports
    ├── report_project1.xlsx
    ├── report_project1.pdf
    └── report_project2.xlsx
    └── report_project2.pdf
