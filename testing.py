import pandas as pd
import re
from datetime import datetime

def test_cover_page(report_path, expected_version):
    """Test cover page elements"""
    # Read the first sheet of the Excel file for cover page content
    try:
        cover_df = pd.read_excel(report_path, sheet_name=0, header=None)
        content = cover_df.apply(lambda row: ' '.join(row.dropna().astype(str)), axis=1).tolist()
    except Exception as e:
        return {
            'title_spelling': {'passed': False, 'message': f"Error reading cover page: {str(e)}"},
            'etl_dates': {'passed': False, 'message': f"Error reading cover page: {str(e)}"},
            'version': {'passed': False, 'message': f"Error reading cover page: {str(e)}"}
        }
    
    # Initialize test results
    test_results = {
        'title_spelling': {'passed': False, 'message': ''},
        'etl_dates': {'passed': False, 'message': ''},
        'version': {'passed': False, 'message': ''}
    }
    
    # Test 1: Check title spelling
    expected_titles = ["Resource Providers Available", "Data accurate as of last successful ETL run"]
    title_errors = []
    
    for line in content:
        for title in expected_titles:
            if title.lower() in line.lower():
                if title not in line:
                    title_errors.append(f"Expected: '{title}' | Found: '{line.strip()}'")
    
    if not title_errors:
        test_results['title_spelling']['passed'] = True
        test_results['title_spelling']['message'] = "All titles spelled correctly"
    else:
        test_results['title_spelling']['message'] = "Title spelling errors:\n" + "\n".join(title_errors)
    
    # Test 2: Check ETL dates (started before completed)
    # Updated pattern to match "ETL - Started" and "CM - Completed"
    etl_pattern = r"ETL - Started: (\d{2}-[A-Za-z]{3}-\d{4} \d{2}:\d{2}:\d{2} [AP]M); CM - Completed: (\d{2}-[A-Za-z]{3}-\d{4} \d{2}:\d{2}:\d{2} [AP]M)"
    date_format = "%d-%b-%Y %I:%M:%S %p"
    
    for line in content:
        match = re.search(etl_pattern, line)
        if match:
            start_str, complete_str = match.groups()
            try:
                start_date = datetime.strptime(start_str, date_format)
                complete_date = datetime.strptime(complete_str, date_format)
                
                if start_date < complete_date:
                    test_results['etl_dates']['passed'] = True
                    test_results['etl_dates']['message'] = f"ETL dates valid: Started {start_str} before Completed {complete_str}"
                else:
                    test_results['etl_dates']['message'] = f"ETL dates invalid: Started {start_str} NOT before Completed {complete_str}"
            except ValueError:
                test_results['etl_dates']['message'] = "Could not parse ETL dates"
            break
        else:
            test_results['etl_dates']['message'] = "ETL date pattern not found in cover page"
    
    # Test 3: Check report version
    version_pattern = r"Version: (\d+\.\d+)"
    for line in content:
        match = re.search(version_pattern, line)
        if match:
            found_version = match.group(1)
            if found_version == expected_version:
                test_results['version']['passed'] = True
                test_results['version']['message'] = f"Version matches: {found_version}"
            else:
                test_results['version']['message'] = f"Version mismatch. Expected: {expected_version}, Found: {found_version}"
            break
    
    return test_results

def test_standard_report_columns(report_path, design_spec_path, design_header_row=7, report_sheet_name=1, report_header_row=2):
    """Test if standard report columns match design spec"""
    # Read design spec columns (CSV file)
    try:
        # Read entire CSV file
        with open(design_spec_path, 'r') as f:
            csv_lines = f.readlines()
        
        # Get header row (row 7, but 0-indexed as 6)
        if len(csv_lines) >= design_header_row:
            design_header_line = csv_lines[design_header_row-1].strip()
            design_columns = [col.strip() for col in design_header_line.split(',') if col.strip()]
        else:
            return {
                'passed': False,
                'message': f"Design spec CSV has fewer than {design_header_row} rows"
            }
    except Exception as e:
        return {
            'passed': False,
            'message': f"Error reading design spec CSV: {str(e)}"
        }
    
    # Read report columns (Excel file)
    try:
        # Read report Excel file
        report_df = pd.read_excel(report_path, sheet_name=report_sheet_name, header=None)
        
        # Get header row (row 2, but 0-indexed as 1)
        if len(report_df) >= report_header_row:
            report_columns = report_df.iloc[report_header_row-1].tolist()
            report_columns = [str(col).strip() for col in report_columns if pd.notna(col)]
        else:
            return {
                'passed': False,
                'message': f"Report sheet has fewer than {report_header_row} rows"
            }
    except Exception as e:
        return {
            'passed': False,
            'message': f"Error reading report Excel: {str(e)}"
        }
    
    # Compare column counts
    if len(design_columns) != len(report_columns):
        return {
            'passed': False,
            'message': f"Column count mismatch. Design: {len(design_columns)}, Report: {len(report_columns)}"
        }
    
    # Enhanced comparison with detailed analysis
    mismatches = []
    for i, (design_col, report_col) in enumerate(zip(design_columns, report_columns)):
        # Normalize spaces by replacing multiple spaces with single space
        design_norm = ' '.join(design_col.split())
        report_norm = ' '.join(report_col.split())
        
        # Check for exact match first
        if design_col == report_col:
            continue
            
        # Check for normalized match (space differences only)
        if design_norm == report_norm:
            # Find the actual space differences
            if len(design_col) != len(report_col):
                mismatches.append(f"Column {i+1}: Space difference - Design='{design_col}' vs Report='{report_col}'")
            else:
                # Character-by-character comparison for exact difference location
                diff_positions = [j for j, (d, r) in enumerate(zip(design_col, report_col)) if d != r]
                if all(design_col[p].isspace() or report_col[p].isspace() for p in diff_positions):
                    mismatches.append(f"Column {i+1}: Space difference - Design='{design_col}' vs Report='{report_col}'")
                else:
                    mismatches.append(f"Column {i+1}: Formatting difference - Design='{design_col}' vs Report='{report_col}'")
        else:
            # Check for case-insensitive match
            if design_norm.lower() == report_norm.lower():
                mismatches.append(f"Column {i+1}: Case difference - Design='{design_col}' vs Report='{report_col}'")
            else:
                # Check for word differences
                design_words = design_norm.lower().split()
                report_words = report_norm.lower().split()
                
                if design_words == report_words:
                    mismatches.append(f"Column {i+1}: Word order difference - Design='{design_col}' vs Report='{report_col}'")
                else:
                    # Find specific word differences
                    diff_words = [(dw, rw) for dw, rw in zip(design_words, report_words) if dw != rw]
                    if diff_words:
                        word_diff_msg = ", ".join(f"'{dw}'≠'{rw}'" for dw, rw in diff_words)
                        mismatches.append(f"Column {i+1}: Word difference - {word_diff_msg} (Design='{design_col}' vs Report='{report_col}')")
                    else:
                        mismatches.append(f"Column {i+1}: Content difference - Design='{design_col}' vs Report='{report_col}'")
    
    if not mismatches:
        return {
            'passed': True,
            'message': f"All {len(design_columns)} columns match perfectly between design spec and report"
        }
    else:
        return {
            'passed': False,
            'message': "Column header differences found:\n" + "\n".join(mismatches)
        }

def test_summary_calculations(report_path):
    """Test summary page calculations against standard report data"""
    try:
        # Read the second sheet (Standard Report)
        report_df = pd.read_excel(report_path, sheet_name=1, header=1)
        
        # Clean up column names by stripping whitespace
        report_df.columns = [str(col).strip() for col in report_df.columns]
        
        # Convert relevant columns to string for comparison
        for col in ['BF', 'Approved', 'Closed']:
            if col in report_df.columns:
                report_df[col] = report_df[col].astype(str).str.strip().str.lower()
        
        # Remove empty rows
        report_df = report_df.dropna(how='all')
    except Exception as e:
        return {
            'brought_forward': {'passed': False, 'message': f"Error reading report data: {str(e)}"},
            'approved': {'passed': False, 'message': f"Error reading report data: {str(e)}"},
            'end_of_period': {'passed': False, 'message': f"Error reading report data: {str(e)}"}
        }
    
    # Initialize test results
    test_results = {
        'brought_forward': {'passed': False, 'expected': 206, 'actual': 0, 'message': ''},
        'approved': {'passed': False, 'expected': 144, 'actual': 0, 'message': ''},
        'end_of_period': {'passed': False, 'expected': 149, 'actual': 0, 'message': ''}
    }
    
    # Test 1: Number of Distinct Approved Providers Brought Forward
    if 'BF' in report_df.columns:
        bf_count = report_df['BF'].str.lower().str.strip().eq('yes').sum()
        test_results['brought_forward']['actual'] = bf_count
        test_results['brought_forward']['passed'] = bf_count == test_results['brought_forward']['expected']
        test_results['brought_forward']['message'] = f"BF count: Expected {test_results['brought_forward']['expected']}, Found {bf_count}"
    else:
        test_results['brought_forward']['message'] = "BF column not found in report"
    
    # Test 2: Number of Providers Approved
    if 'Approved' in report_df.columns:
        approved_count = report_df['Approved'].str.lower().str.strip().eq('yes').sum()
        test_results['approved']['actual'] = approved_count
        test_results['approved']['passed'] = approved_count == test_results['approved']['expected']
        test_results['approved']['message'] = f"Approved count: Expected {test_results['approved']['expected']}, Found {approved_count}"
    else:
        test_results['approved']['message'] = "Approved column not found in report"
    
    # Test 3: Number of Distinct Approved Providers End of Period
    if 'Closed' in report_df.columns:
        closed_count = report_df['Closed'].str.lower().str.strip().eq('yes').sum()
        calculated_eop = (test_results['brought_forward']['actual'] + test_results['approved']['actual']) - closed_count
        test_results['end_of_period']['actual'] = calculated_eop
        test_results['end_of_period']['passed'] = calculated_eop == test_results['end_of_period']['expected']
        test_results['end_of_period']['message'] = f"End of Period: Expected {test_results['end_of_period']['expected']}, Calculated {calculated_eop} (BF: {test_results['brought_forward']['actual']} + Approved: {test_results['approved']['actual']} - Closed: {closed_count})"
    else:
        test_results['end_of_period']['message'] = "Closed column not found in report"
    
    return test_results

def run_all_tests(report_path, design_spec_path, expected_version):
    """Run all tests and return consolidated results"""
    print(f"Running tests for report: {report_path}")
    print(f"Expected version: {expected_version}")
    
    # Run cover page tests
    print("\n=== Cover Page Tests ===")
    cover_results = test_cover_page(report_path, expected_version)
    for test_name, result in cover_results.items():
        status = "PASSED" if result['passed'] else "FAILED"
        print(f"{test_name.upper()}: {status} - {result['message']}")
    
    # Run standard report column tests
    print("\n=== Standard Report Column Tests ===")
    column_result = test_standard_report_columns(report_path, design_spec_path)
    status = "PASSED" if column_result['passed'] else "FAILED"
    print(f"COLUMN_MATCH: {status} - {column_result['message']}")
    
    # Run summary calculations tests
    print("\n=== Summary Calculations Tests ===")
    summary_results = test_summary_calculations(report_path)
    for test_name, result in summary_results.items():
        status = "PASSED" if result['passed'] else "FAILED"
        print(f"{test_name.upper()}: {status} - {result['message']}")
    
    # Calculate overall status
    cover_passed = all(r['passed'] for r in cover_results.values())
    column_passed = column_result['passed']
    summary_passed = all(r['passed'] for r in summary_results.values())
    all_passed = cover_passed and column_passed and summary_passed
    
    print("\n=== FINAL RESULT ===")
    print("ALL TESTS PASSED" if all_passed else "SOME TESTS FAILED")

# Example usage
if __name__ == "__main__":
    report_excel = "CB080 - Resource Providers Available (Ottawa 2024-2025).xlsx"
    design_spec_csv = "CB080 - Design Spec - Resources Available(Standard 1 Report).csv"
    expected_version = "1.5"
    
    run_all_tests(report_excel, design_spec_csv, expected_version)

# PS D:\work\college_work\Coop_1\automation\Automation_Testing> & "C:/Program Files/Python312/python.exe" d:/work/college_work/Coop_1/automation/Automation_Testing/testing.py
# d:\work\college_work\Coop_1\automation\Automation_Testing\testing.py:286: SyntaxWarning: invalid escape sequence '\w'
#   '''
# Running tests for report: CB080 - Resource Providers Available (Ottawa 2024-2025).xlsx
# Expected version: 1.5

# === Cover Page Tests ===
# TITLE_SPELLING: PASSED - All titles spelled correctly
# ETL_DATES: PASSED - ETL dates valid: Started 21-Jul-2025 11:31:39 PM before Completed 22-Jul-2025 05:46:16 AM
# VERSION: FAILED - Version mismatch. Expected: 1.5, Found: 1.4

# === Standard Report Column Tests ===
# COLUMN_MATCH: FAILED - Column header differences found:
# Column 7: Space difference - Design='Provider Status Owner  First Name' vs Report='Provider Status Owner First Name'
# Column 8: Word difference - 'date'≠'end' (Design='Provider Owner  Last Name as of Report Date' vs Report='Provider Owner Last Name as of Report End Date')
# Column 9: Word difference - 'date'≠'end' (Design='Provider Owner  First Name as of Report Date' vs Report='Provider Owner First Name as of Report End Date')
# Column 15: Word difference - 'codes'≠'code' (Design='Secondary Eligibility Spectrum Codes' vs Report='Secondary Eligibility Spectrum Code')

# === Summary Calculations Tests ===
# BROUGHT_FORWARD: PASSED - BF count: Expected 206, Found 206
# APPROVED: PASSED - Approved count: Expected 144, Found 144
# END_OF_PERIOD: PASSED - End of Period: Expected 149, Calculated 149 (BF: 206 + Approved: 144 - Closed: 201)

# === FINAL RESULT ===
# SOME TESTS FAILED