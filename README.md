# 📊 Excel Report Validator

[![Python](https://img.shields.io/badge/python-3.8%2B-blue)](https://www.python.org/)
[![Pandas](https://img.shields.io/badge/pandas-1.3%2B-orange)](https://pandas.pydata.org/)

Automated testing tool for validating Excel reports against design specifications with dynamic summary validation. 🔍✨

---

## 🌟 Enhanced Features

### ✅ Smart Cover Page Validation
- Title spelling verification 📝  
- ETL date sequence validation ⏳  
- Version number matching 🔖  

### 📋 Intelligent Column Comparison
- Exact match detection ✔️  
- Whitespace difference spotting ␣  
- Case sensitivity analysis 🔠  
- Word order validation 🔄  

### 🧮 Dynamic Summary Validation (NEW!)
- **Auto-reads expected values** from Summary_Page (Sheet 3) 📖  
  - Brought Forward from **B3**  
  - Approved from **B4**  
  - End of Period from **B6**  
- Real-time calculation verification 🧮  
- Clear source tracking in results 📌  

---

## 🛠️ Installation

```bash
git clone https://github.com/SUKH2022/Automation_Testing.git
cd Automation_Testing
pip install pandas openpyxl
```

---

## 🚀 Usage

from report_validator import run_all_tests

```bash
run_all_tests(
    report_path="your_report.xlsx",
    design_spec_path="design_spec.csv", 
    expected_version="1.5"
)
```

## 📊 Sample Output

```bash
=== Cover Page Tests ===
TITLE_SPELLING: PASSED ✅ - All titles correct
ETL_DATES: PASSED ✅ - Dates valid (21-Jul-2025 → 22-Jul-2025)
VERSION: FAILED ❌ - Expected 1.5, found 1.4

=== Column Tests ===
COLUMN_MATCH: FAILED ❌ - 4 differences found
   ▶ Column 7: Space difference
   ▶ Column 8: 'date'≠'end' 
   ▶ Column 9: 'date'≠'end'
   ▶ Column 15: 'codes'≠'code'

=== Summary Tests ===
BROUGHT_FORWARD: PASSED ✅ - 206 (matches B3)
APPROVED: PASSED ✅ - 144 (matches B4) 
END_OF_PERIOD: PASSED ✅ - 149 (matches B6)
   Calculation: 206 + 144 - 201 = 149 ✔️

=== FINAL RESULT ===
SOME TESTS FAILED ‼️
```

## 🆕 What's New in v1.1

- 🎯 Dynamic Summary Validation: No more hardcoded values!

- 📌 Clear Value Sourcing: Shows exactly which cells were used

- 🧮 Calculation Breakdown: Detailed math in End of Period results

- 🛠️ More Robust: Better error handling for summary page reads
