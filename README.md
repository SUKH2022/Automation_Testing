# 📊 Excel Report Validator

[![Python](https://img.shields.io/badge/python-3.8%2B-blue)](https://www.python.org/)
[![Pandas](https://img.shields.io/badge/pandas-1.3%2B-orange)](https://pandas.pydata.org/)

An automated testing tool for validating Excel reports against design specifications. 🔍✨

---

## 🌟 Features

### ✅ Cover Page Validation
- Title spelling check 📝  
- ETL date validation ⏳  
- Version matching 🔖  

### 📋 Column Comparison
- Exact match verification ✔️  
- Space difference detection ␣  
- Case sensitivity check 🔠  
- Word order validation 🔄  

### 🧮 Summary Calculations
- Brought forward count validation ➡️  
- Approved provider verification ✅  
- End-of-period calculation check 🧐  

---

## 🛠️ Installation

Clone the repository:

```bash
git clone https://github.com/SUKH2022/Automation_Testing/testing.git
cd Automation_Testing
```

Install dependencies:

```bash
pip install pandas openpyxl
```

## 🚀 Usage

```bash
from report_validator import run_all_tests

run_all_tests(
    report_path="your_report.xlsx",
    design_spec_path="design_spec.csv",
    expected_version="1.0"
)
```

## 📝 Sample Output
```bash
=== Cover Page Tests ===
TITLE_SPELLING: PASSED ✅ - All titles spelled correctly
ETL_DATES: PASSED ✅ - ETL dates valid
VERSION: FAILED ❌ - Version mismatch

=== Column Tests ===
COLUMN_MATCH: FAILED ❌ - 3 differences found

=== Summary Tests ===
BROUGHT_FORWARD: PASSED ✅ - Count matches
APPROVED: PASSED ✅ - Count matches
END_OF_PERIOD: PASSED ✅ - Calculation correct
```
