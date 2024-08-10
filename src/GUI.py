import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import sys
import pandas as pd

from execution_functions import process_files
from export_individual_sheets import export_all_sheets,export_csm_sheet

def validate_file_extension(file_path, expected_extension):
    _, extension = os.path.splitext(file_path)
    return extension.lower() == expected_extension.lower()

def upload_files(upload_type, expected_extension):
    root = tk.Tk()
    root.withdraw()
    
    if upload_type == "job list":
        file_paths = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        file_paths = [file_paths] if file_paths else []
    elif upload_type == "budget":
        file_paths = filedialog.askopenfilenames(filetypes=[("CSV files", "*.csv")])
        file_paths = list(file_paths)
    elif upload_type == "payroll" or upload_type == "workbills":
        file_paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx")])
        file_paths = list(file_paths)
    
    valid_files = [file for file in file_paths if validate_file_extension(file, expected_extension)]
    
    invalid_files = set(file_paths) - set(valid_files)
    if invalid_files:
        messagebox.showerror("Invalid Files", f"Error: Invalid file extension for {', '.join(invalid_files)}. Expected {expected_extension} file(s).")
    
    return valid_files

def create_gui():
    job_list_files = []
    budget_files = []
    payroll_files = []
    workbills_files = []
    generated_workbook_path = None

    upload_types = ["job list", "budget", "payroll", "workbills"]
    current_upload_index = 0

    def reset_application():
        nonlocal job_list_files, budget_files, payroll_files, workbills_files, generated_workbook_path, current_upload_index
        job_list_files = []
        budget_files = []
        payroll_files = []
        workbills_files = []
        generated_workbook_path = None
        current_upload_index = 0
        
        # Reset GUI elements
        upload_type_var.set(upload_types[0])
        job_list_success_var.set(0)
        budget_success_var.set(0)
        payroll_success_var.set(0)
        workbills_success_var.set(0)
        confirm_uploads_var.set(0)
        
        # Hide export section
        export_frame.pack_forget()
        
        # Clear CSM dropdown
        csm_dropdown['values'] = []
        csm_dropdown.set('')
        
        # Disable buttons
        export_csm_button['state'] = 'disabled'
        export_all_button['state'] = 'disabled'
        open_file_button['state'] = 'disabled'
        open_folder_button['state'] = 'disabled'
        
        messagebox.showinfo("Reset", "Application has been reset.")
        
    def open_generated_file():
        if generated_workbook_path and os.path.exists(generated_workbook_path):
            try:
                if sys.platform == "win32":
                    os.startfile(generated_workbook_path)
                elif sys.platform == "darwin":
                    subprocess.call(["open", generated_workbook_path])
                else:
                    subprocess.call(["xdg-open", generated_workbook_path])
            except Exception as e:
                messagebox.showerror("Error", f"Failed to open file: {str(e)}")
        else:
            messagebox.showerror("Error", "Generated file not found.")

    def open_generated_folder():
        if generated_workbook_path and os.path.exists(generated_workbook_path):
            folder_path = os.path.dirname(generated_workbook_path)
            try:
                if sys.platform == "win32":
                    os.startfile(folder_path)
                elif sys.platform == "darwin":
                    subprocess.call(["open", folder_path])
                else:
                    subprocess.call(["xdg-open", folder_path])
            except Exception as e:
                messagebox.showerror("Error", f"Failed to open folder: {str(e)}")
        else:
            messagebox.showerror("Error", "Generated file folder not found.")

    def on_upload():
        nonlocal current_upload_index
        upload_type = upload_types[current_upload_index]
        if upload_type in ['job list', 'payroll', 'workbills']:
            expected_extension = '.xlsx'
        elif upload_type == 'budget':
            expected_extension = '.csv'
        else:
            expected_extension = '*.*'
        valid_files = upload_files(upload_type, expected_extension)
        if valid_files:
            if upload_type == 'job list':
                job_list_files.extend(valid_files)
                job_list_success_var.set(1)
            elif upload_type == 'budget':
                budget_files.extend(valid_files)
                budget_success_var.set(1)
            elif upload_type == 'payroll':
                payroll_files.extend(valid_files)
                payroll_success_var.set(1)
            elif upload_type == 'workbills':
                workbills_files.extend(valid_files)
                workbills_success_var.set(1)
            current_upload_index = (current_upload_index + 1) % len(upload_types)
            upload_type_var.set(upload_types[current_upload_index])
        else:
            if upload_type == 'job list':
                job_list_success_var.set(0)
            elif upload_type == 'budget':
                budget_success_var.set(0)
            elif upload_type == 'payroll':
                payroll_success_var.set(0)
            elif upload_type == 'workbills':
                workbills_success_var.set(0)

        if all([job_list_files, budget_files, payroll_files, workbills_files]):
            confirm_uploads_var.set(1)

    def on_generate():
        nonlocal generated_workbook_path
        if not all([job_list_files, budget_files, payroll_files, workbills_files]):
            messagebox.showerror("Error", "Please upload all required files before generating the report.")
            return
        output_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if output_path:
            try:
                generated_workbook_path = process_files(job_list_files, budget_files, payroll_files, workbills_files, output_path)
                if generated_workbook_path:
                    messagebox.showinfo("Success", f"CSM KPI SHEET generated at: {generated_workbook_path}")
                    populate_csm_dropdown(generated_workbook_path)
                    enable_export_section()
                    open_file_button['state'] = 'normal'
                    open_folder_button['state'] = 'normal'
                else:
                    messagebox.showerror("Error", "Failed to generate Excel file. Check error log for details.")
            except Exception as e:
                messagebox.showerror("Error", f"An error occurred: {str(e)}")
                with open("error_log.txt", "a") as log_file:
                    log_file.write(f"Error processing files: {e}\n")

    def populate_csm_dropdown(workbook_path):
        try:
            xl = pd.ExcelFile(workbook_path)
            csm_names = [sheet for sheet in xl.sheet_names if sheet != 'Index']
            csm_dropdown['values'] = csm_names
            if csm_names:
                csm_dropdown.set(csm_names[0])
        except Exception as e:
            messagebox.showerror("Error", f"Failed to populate CSM dropdown: {str(e)}")    
        
    def enable_export_section():
        export_frame.pack(pady=10)
        csm_dropdown['state'] = 'readonly'
        export_csm_button['state'] = 'normal'
        export_all_button['state'] = 'normal'

    def on_export_csm():
        csm_name = csm_var.get()
        if csm_name and generated_workbook_path:
            export_csm_sheet(generated_workbook_path, csm_name)

    def on_export_all():
        if generated_workbook_path:
            export_all_sheets(generated_workbook_path)

    def on_closing():
        root.quit()
        root.destroy()
        sys.exit(0)

    root = tk.Tk()
    root.title("CSM KPI Sheet Generator")
    root.geometry("500x600")  # Increased height to accommodate new section
    root.protocol("WM_DELETE_WINDOW", on_closing)

    main_frame = ttk.Frame(root, padding="10")
    main_frame.pack(fill=tk.BOTH, expand=True)

    # File Upload Section
    upload_frame = ttk.LabelFrame(main_frame, text="File Upload", padding="10")
    upload_frame.pack(fill=tk.X, pady=10)

    tk.Label(upload_frame, text="Upload Type").grid(row=0, column=0, pady=5)
    upload_type_var = tk.StringVar(value=upload_types[current_upload_index])
    tk.OptionMenu(upload_frame, upload_type_var, *upload_types).grid(row=0, column=1)
    tk.Label(upload_frame, text="File Types: .csv, .xlsx").grid(row=1, columnspan=2, pady=5)
    tk.Button(upload_frame, text="Upload", command=on_upload).grid(row=2, column=1, pady=5)

    job_list_success_var = tk.IntVar()
    tk.Checkbutton(upload_frame, text="Job List Uploaded", variable=job_list_success_var, state='disabled').grid(row=3, column=0, pady=5)

    budget_success_var = tk.IntVar()
    tk.Checkbutton(upload_frame, text="Budget Uploaded", variable=budget_success_var, state='disabled').grid(row=3, column=1, pady=5)

    payroll_success_var = tk.IntVar()
    tk.Checkbutton(upload_frame, text="Payroll Uploaded", variable=payroll_success_var, state='disabled').grid(row=4, column=0, pady=5)
    
    workbills_success_var = tk.IntVar()
    tk.Checkbutton(upload_frame, text="Workbills Uploaded", variable=workbills_success_var, state='disabled').grid(row=4, column=1, pady=5)

    confirm_uploads_var = tk.IntVar()
    tk.Checkbutton(upload_frame, text="Confirm all files uploaded", variable=confirm_uploads_var, state='disabled').grid(row=5, columnspan=2, pady=5)

    # Generate and Reset Buttons
    button_frame = ttk.Frame(upload_frame)
    button_frame.grid(row=6, column=0, columnspan=2, pady=10)

    generate_button = tk.Button(button_frame, text="Generate", command=on_generate)
    generate_button.pack(side=tk.LEFT, padx=5)

    reset_button = tk.Button(button_frame, text="Reset", command=reset_application)
    reset_button.pack(side=tk.LEFT, padx=5)
    
    # Open File/Folder Buttons
    open_file_button = tk.Button(upload_frame, text="Open Generated File", command=open_generated_file, state='disabled')
    open_file_button.grid(row=7, column=0, pady=5, padx=5, sticky='ew')

    open_folder_button = tk.Button(upload_frame, text="Open Generated Folder", command=open_generated_folder, state='disabled')
    open_folder_button.grid(row=7, column=1, pady=5, padx=5, sticky='ew')

    # Export Section (initially hidden)
    export_frame = ttk.LabelFrame(main_frame, text="Export Options", padding="10")

    csm_var = tk.StringVar()
    csm_dropdown = ttk.Combobox(export_frame, textvariable=csm_var, state='disabled')
    csm_dropdown.pack(pady=5)

    export_csm_button = tk.Button(export_frame, text="Export Selected CSM", command=on_export_csm, state='disabled')
    export_csm_button.pack(pady=5)

    export_all_button = tk.Button(export_frame, text="Export All CSMs", command=on_export_all, state='disabled')
    export_all_button.pack(pady=5)

    root.mainloop()

if __name__ == "__main__":
    create_gui()
