# 📊 Collection Report Automation – UiPath Project

This project automates the generation and distribution of a **collection report** to managers, summarizing all outstanding invoices. The report is extracted from the financial system, formatted in Excel with visual cues, and sent via email with summary insights.

---

## 🚀 Purpose

- Generate and send a **summary email** containing all unpaid tax invoices and a detailed Excel report.
- Sent automatically to **all project managers** in the company.

---

## 🎯 Business Value

- ⏱️ **Time-saving**: Automates a process that previously took significant manual effort.
- ❌ **Error reduction**: Minimizes human errors in data collection and reporting.
- ⚡ **Improved efficiency**: Speeds up the financial reporting process, enhancing decision-making for finance and project management teams.

---

## ⏲️ Runtime Comparison

| Process Type          | Duration         |
|-----------------------|------------------|
| Manual                | ???              |
| Automated (UiPath)    | ~20 seconds      |

---

## 🛠️ Automation Workflow Overview

1. **Data Extraction**  
   - From salary accounting database using stored procedures.
2. **Excel Report Generation**  
   - Includes detailed data with traffic light color coding by status.
3. **Statistics Visualization**  
   - Pie chart and summary table derived from Excel data.
4. **Email Dispatch**  
   - Email includes the summary visuals and Excel file as an attachment.

---

## 📁 Output Files

- **Excel Report**  
  - `Collection Report`: Includes all statuses (on time, delayed, debt at risk).  
  - `Collection Report - Limited`: Includes only `delayed` and `debt at risk`.
- **Email Formats**  
  - Desktop, mobile, dark mode views supported.
- **PDF + HTML**  
  - Summary visuals also attached as PDF and embedded as HTML.

---

## ⏰ Trigger

- Weekly – Every **Thursday at 13:00** via Orchestrator trigger.

---

## 🧩 Technical Overview

### Dispatcher Process

- Dispatcher name: `Finance_Report_Dispatch`
- Creates queue items per manager for 3 finance reports.
- Business process name for this report: `collectionReport`

### Performer Process

- Retrieves queue item and generates manager-specific report.

---

## 📦 Key Components

| Component                                 | Description                                  |
|------------------------------------------|----------------------------------------------|
| `Extract_Report_Data_From_DB.xaml`       | Full report data via `sp_GetCollectionReportDataByManager_RPA` |
| `Extract_Limited_Report_Data_From_DB.xaml` | Limited report data via `sp_GetLimitedCollectionReportDataByManager_RPA` |
| `Generate_Excel_And_PDF.xaml`            | Excel formatting and visualizations         |
| VBScript Files (`VB Scripts/`)           | Styling, summary table, pie chart, export   |
| `Send_Report_Via_Mail.xaml`              | Email sending with dynamic CC via asset     |

---

## 📊 Visualizations

- Pie Chart (`CreatePieChart.vbs`)
- Summary Table (`CreateSummaryTable.vbs`)
- PDF Export (`ExportSheetToPDF.vbs`)
- HTML Export (`ExportToHTML.vbs`)

---

## ☁️ Deployment & Cloud Info

- Location: Orchestrator folder `finance/CollectionReport`
- Queues:
  - `TEST` – For development/testing
  - `PROD` – For live runs
- Source Control: [GitHub Repo](https://github.com/uipathhmshms/CollectionReport)

---

## 🔮 Future Enhancements

- Reports filtered by specific companies (e.g., TCS, H.B. Eisenberger)
- Restricted reports for approved managers only
- Automated status update email to management post-execution

---

## 📞 Contact

For support or enhancements, contact: **LidorM**
