# 🛡️ Nessus Report Automation with Python

This Python script automates the creation of structured, client-ready vulnerability assessment reports from Nessus CSV exports. It formats the data into a well-organized DOCX file with color-coded risk sections, charts, and tables — saving you hours of manual work.

## 🚀 Features

- 📌 Groups vulnerabilities by host and severity
- 🎨 Color-coded sections for CRITICAL, MEDIUM, and LOW risks
- 📊 Automatically inserts bar charts for:
  - Overall vulnerability statistics
  - Host-specific vulnerability breakdowns
- 📋 Clean tables including:
  - Vulnerability Title
  - Description
  - Remediation
  - **Impact** (left blank for manual customization)
- 📄 Generates reports based on a DOCX template you provide

## 🧾 Requirements

To use this script, you’ll need to provide a custom DOCX file as `REPORT_TEMPLATE`, which should include:

- ✅ Company-branded cover page (Page 1)
- ✅ Page 2 left blank (for auto-generated Table of Contents)
- ✅ Header and footer elements applied consistently

## ✏️ Note

- The **Impact** field is not filled automatically — Nessus doesn't provide this, and it's typically tailored to each client.
- Screenshots and sample outputs are available in the `/examples` folder.

## 🛠️ Technologies Used

- Python 3
- `python-docx` for DOCX manipulation
- `matplotlib` for chart generation
- `pandas` for data processing

```plaintext
.
├── main.py
├── REPORT_TEMPLATE.docx
├── input.csv
├── output_report.docx

## 📧 Contact
For suggestions or questions, feel free to reach out via LinkedIn or open an issue.
