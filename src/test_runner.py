import os
import sys
from execution_functions import process_files

def check_file_exists(file_path):
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")

def run_test():
    try:
        # Specify the paths to your test files here
        job_list_files = ['../test_input/CSM by job as at 29 May 2024.xlsx']
        budget_files = [
            '../test_input/FE19052024_BudgetRoster_extract_summedvalues.csv',
            '../test_input/FE05012024_BudgetRoster_extract_summedvalues.csv',
            '../test_input/FE27122024_BudgetRoster_extract_summedvalues.csv'
        ]
        payroll_files = [
            '../test_input/Pay By Type_FE05012024.xlsx',
            '../test_input/Pay By Type_FE19052024.xlsx',
            '../test_input/Pay By Type_FE27122024.xlsx'
        ]
        workbill_files = [
            '../test_input/Workbills Hours_FE05012024.xlsx',
            '../test_input/Workbills Hours_FE19052024.xlsx',
            '../test_input/Workbills Hours_FE27122024.xlsx',
        ]
        
        # Specify the output path
        output_path = '../test_output/test.xlsx'
        
        # Check if all input files exist
        for file in job_list_files + budget_files + payroll_files:
            check_file_exists(file)
        
        # Call the process_files function directly
        process_files(job_list_files, budget_files, payroll_files, workbill_files, output_path)
        
        # Check if output file was created
        if not os.path.exists(output_path):
            raise FileNotFoundError(f"Output file was not created: {output_path}")
        
        # Check if output file is not empty
        if os.path.getsize(output_path) == 0:
            raise ValueError(f"Output file is empty: {output_path}")
        
        print(f"Test completed successfully. Output file: {output_path}")
        print(f"Output file size: {os.path.getsize(output_path)} bytes")
        
        # You can add more specific checks here, e.g., checking the content of the Excel file
        
    except Exception as e:
        print(f"Test failed: {str(e)}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    run_test()