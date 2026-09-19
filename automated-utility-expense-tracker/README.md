# 📊 Automated Multi-Vendor Expense Tracking Pipeline

An automated ETL (Extract, Transform, Load) pipeline built in **Google Apps Script** that automates the collection, extraction, and visualization of monthly utility expenses across multiple Malaysian service providers into a centralized Google Sheets dashboard.

---

## 🛠️ Tech Stack & Concepts
* **Language**: JavaScript / Google Apps Script
* **APIs & Services**: Gmail API, Google Drive API, Google Sheets API
* **Data Processing**: Regular Expressions (Regex), Dynamic Date Normalization, Duplicate Handling
* **Automation**: Time-Driven Scheduled Triggers

---

## 💡 Problem & Solution

### The Problem
Tracking recurring monthly bills across different platforms (TIME Fibre Internet, SK Magic, TNB Electric via Boost & myTNB) required manual effort to locate e-receipts, parse varying date/amount formats, and input records manually into a budget spreadsheet.

### The Solution
A fully automated, zero-maintenance pipeline that:
1. **Scans Gmail & Drive**: Searches for monthly bill attachments (PDFs) and e-receipt confirmation emails.
2. **Parses & Transforms Data**: Uses custom regex to extract exact billing dates, reference numbers, and transaction amounts while handling format variations.
3. **Consolidates to Dashboard**: Aggregates entries into individual vendor sheets and updates a master summary dashboard.
4. **Automates Monthly Runs**: Scheduled time-driven trigger runs on the 1st of every month.

---

## 📁 Repository Structure

```text
├── Automation.js       # Master wrapper script for monthly triggers
├── TIME_Parser.js      # Gmail & Drive PDF invoice parsing
├── SKMagic_Parser.js   # HTML email body parsing for recurring charges
├── TNB_Parser.js       # Unified parser for Boost & myTNB e-receipts
└── README.md           # Documentation
