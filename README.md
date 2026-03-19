# 📊 Excel to PDF Automation System

## Overview

This project automates the process of converting Excel data into professional PDF reports.
Users can simply upload Excel files into a folder, and the system automatically generates structured PDF reports with company branding, tables, and visual charts.

## ✨ Features

* 📁 Automatic Excel file detection from a folder
* 📄 Converts Excel data into well-formatted PDF reports
* 🏢 Adds company logo and details
* 📊 Generates charts for easy data visualization
* 📋 Clean table formatting for better readability
* ⚡ Fully automated workflow (no manual intervention)

## 🛠️ Tech Stack

* Python
* Pandas
* Matplotlib / Seaborn (for charts)
* ReportLab / FPDF (for PDF generation)
* OpenPyXL (for Excel handling)

## 📂 Project Workflow

1. User uploads Excel file into the `input/` folder
2. Script reads and processes the data
3. Tables and charts are generated
4. PDF report is created with:

   * Company logo
   * Structured tables
   * Graphs
5. Final PDF is saved in the `output/` folder

## ▶️ How to Run

pip install -r requirements.txt
python main.py

## 📌 Use Cases

* Business reporting
* Sales analysis
* Invoice/report generation
* Data visualization automation

## 🔮 Future Improvements

* Web interface using Flask/Django
* Drag-and-drop upload UI
* Email integration to send reports automatically
* Support for multiple file formats

## 👩‍💻 Author
R.Kaviyapriya
