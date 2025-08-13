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

def test_standard_report_columns(report_path, design_spec_path, header_row=8, data_start_row=2):
    """Test if standard report columns match design spec"""
    # Read design spec (CSV file)
    try:
        design_spec = pd.read_csv(design_spec_path, header=header_row-1, nrows=1)
        design_columns = [col.strip() for col in design_spec.columns]
    except Exception as e:
        return {
            'passed': False,
            'message': f"Error reading design spec: {str(e)}"
        }
    
    # Read report (Excel file)
    try:
        report_df = pd.read_excel(report_path, sheet_name=1, header=data_start_row-1, nrows=1)
        report_columns = [col.strip() for col in report_df.columns]
    except Exception as e:
        return {
            'passed': False,
            'message': f"Error reading report: {str(e)}"
        }
    
    # Compare columns
    if len(design_columns) != len(report_columns):
        return {
            'passed': False,
            'message': f"Column count mismatch. Design: {len(design_columns)}, Report: {len(report_columns)}"
        }
    
    mismatches = []
    for i, (design_col, report_col) in enumerate(zip(design_columns, report_columns)):
        if str(design_col).strip().lower() != str(report_col).strip().lower():
            mismatches.append(f"Column {i+1}: Design='{design_col}' | Report='{report_col}'")
    
    if not mismatches:
        return {
            'passed': True,
            'message': f"All {len(design_columns)} columns match design spec"
        }
    else:
        return {
            'passed': False,
            'message': "Column mismatches:\n" + "\n".join(mismatches)
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