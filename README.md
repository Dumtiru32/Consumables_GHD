# Consumables_GHD
Automated generating of excels and sending emails using VBA
# 📧 Automated Email Sending via Power Automate and Access

This project automates the process of sending emails to Key Account Managers (KAMs) using Microsoft 365 and Power Automate. 
It replaces traditional Outlook VBA automation to avoid security prompts and improve reliability.

---

## 🚀 Overview

- Export KAM data from Microsoft Access to Excel.
- Use Power Automate to read the Excel file and send emails via Microsoft 365.
- Optionally attach files from dynamically constructed folder paths.

---

## 📁 Excel File Structure

The Excel file (`Consumables_GENT.xlsx`) should contain the following columns:

| Column         | Description                                      |
|----------------|--------------------------------------------------|
| `KAM_Name`     | Full name of the Key Account Manager             |
| `KAM_toemail`  | Primary recipient email address                  |
| `KAM_ccemail`  | CC email addresses (optional)                    |
| `KAM_bccemail` | BCC email addresses (optional)                   |
| `PostingMonth` | Month and year of the consumables report        |
| `FolderPath`   | Full path to the folder containing attachments   |

A sample file is included in this repository.

---

## 🧩 Access VBA Export Script

Use the following VBA code in Access to export your query to Excel:

```vba
# Access VBA Export Automation

This project automates the export of Excel reports per customer from a Microsoft Access database. Each report includes:

- A **Data** sheet with raw data from the `[Consumables]` table.
- An **Overview** sheet with a PivotTable summarizing the data.

## 🔧 Features

- Filters customers using `[KeyAccountManager]![KAM_segment6]`.
- Filters `[Customer]` records based on `[Customer_Seg4]` and `[Ledger]` suffix.
- Exports Excel files to a folder defined in `[KeyAccountManager]![KAM_location]`.
- Folder naming format: `Consumables MM YY`.
- File naming format: `Consumables <Customer Name> MM YY.xlsx`.
- Modular VBA code with robust error handling and logging.

## 📤 Email Automation

Includes a procedure to send emails via Outlook and attach all files from a specified folder.

## 🚀 How to Use

1. Import the `ExportConsumablesGHD.bas` module into your Access project.
2. Add a button to your form and name it `buttonConsumablesGHD`.
3. Set the button's `On Click` event to:

```vba
=buttonConsumablesGHD_Click()
```

4. Run the form and click the button to generate and export reports.

## 📁 Folder Structure

```
<Export Root Folder>/
├── Consumables 08 25/
│   ├── Consumables CustomerA 08 25.xlsx
│   ├── Consumables CustomerB 08 25.xlsx
│   └── ...
```

## 📄 Main Procedure

- `buttonConsumablesGHD_Click`: Main entry point for the export logic.

## 📬 Email Procedure (Optional)

- `SendEmailWithAttachments`: Sends an email with all files from a folder attached.

## 🛠 Requirements

- Microsoft Access with VBA support
- Microsoft Excel
- Microsoft Outlook (for email automation)

## 📜 License

MIT License

DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, _
    "YourQueryName", "C:\Path\To\Consumables_GENT.xlsx", True
