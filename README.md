#  Google Account Analysis Automation

##  Overview

This project was developed to automate the analysis of Google accounts, replacing a manual process that was time-consuming and prone to human error.

The application was built with **Python** using **Selenium** for browser automation and **openpyxl** for Excel file manipulation. It automatically accesses an internal system, analyzes account information, and updates a control spreadsheet in real time.

The goal of this automation is to significantly improve productivity, reduce analysis time, and ensure consistent validation based on predefined business rules.

---

## Features

-  Reads account information from an Excel spreadsheet (`gmails.xlsx`)
-  Automates account analysis through browser interaction
-  Inspects HTML elements to validate predefined conditions
-  Automatically updates the account status after validation
-  Generates a complete audit trail directly in the spreadsheet
-  Includes an executable (`.exe`) for non-technical users
-  User guide available for analysts

---

## Productivity Impact

### Manual Process

- Approximately **10–15 minutes** per account.

### Automated Process

- Approximately **1 minute** per account.
- Up to **480 accounts** processed by one analyst during an 8-hour workday.
- Around **3,000 accounts** processed daily by a team of 10 analysts.

### Benefits

- Significant reduction in processing time.
- Improved accuracy through automated validation.
- Reduced human error.
- Standardized account verification.
- Increased operational efficiency.

---

## Technologies

### Language

- Python

### Libraries

- Selenium
- openpyxl

---

## Project Structure

```text
.
├── gmails.xlsx          # Spreadsheet containing the accounts to analyze
├── config.json          # User configuration
├── main.py              # Main application
├── requirements.txt
└── README.md
```

---

## How It Works

1. Reads all email addresses from **gmails.xlsx**.
2. Logs into the internal system.
3. Opens each account.
4. Extracts and analyzes the required information.
5. Validates predefined business rules.
6. Updates the spreadsheet with the corresponding status.

If an account satisfies all validation criteria, it is automatically marked as:

```
Data Reviewed (DR)
```

indicating that the account is ready for deletion.

---

## Usage

> **Note:** This application was designed to run only within the company's internal environment.

1. Place **gmails.xlsx** inside:

```text
C:\Analyze
```

2. Add your email address to:

```text
config.json
```

3. Run the executable (`.exe`) included in this repository.

4. The application will automatically analyze all accounts and update their status in the spreadsheet.

---

## Results

The automation transformed a process that previously required months of manual work into a task that can be completed in just a few days.

By automating repetitive operations, analysts can focus on more valuable activities instead of manually reviewing thousands of accounts.

---

## Disclaimer

This project was originally designed for an internal corporate environment.

Sensitive business logic and proprietary information have been omitted from this public version.

The architecture can easily be adapted to support different workflows and validation rules.

---

## Author

**Ivan Bragante**

- Python
- Selenium
- openpyxl
