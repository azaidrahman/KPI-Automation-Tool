import pandas as pd
import re
from datetime import datetime
#################### DATES #####################
# Payroll date
# def get_date(file_path):
#     # payroll_date_str = pd.read_excel(file_path, usecols=[0], nrows=2, header=None).iloc[1,0]
    
#     file_name = file_path.split('/')[-1]
#     parts = file_name.split('_')
#     # The date part will always be the part that matches '%Y%m%d'
#     for part in parts:
#         if len(part) == 8 and part.isdigit():
#             payroll_date_str = part
#             break
#     else:
#         raise ValueError("No valid date found in file name")
#     # print(payroll_date_str)
#     date = pd.to_datetime(payroll_date_str, format='%Y%m%d')

#     return date

def get_date(file_path,error_dict):
    file_name = file_path.split('/')[-1]
    
    # Regular expression to match date patterns
    date_pattern = r'FE(\d{8})\b'
    
    match = re.search(date_pattern, file_name)
    if match:
        date_str = match.group(1)
        
        # Try parsing as YYYYMMDD
        try:
            return pd.Timestamp(datetime.strptime(date_str, '%Y%m%d'))
        except ValueError:
            pass
        
        # Try parsing as DDMMYYYY
        try:
            return pd.Timestamp(datetime.strptime(date_str, '%d%m%Y'))
        except ValueError:
            pass
        
        # If both fail, assume it's DDMMYYYY but with ambiguous day/month
        try:
            day = int(date_str[:2])
            month = int(date_str[2:4])
            year = int(date_str[4:])
            
            # If day is greater than 12, it must be DDMMYYYY
            if day > 12:
                return pd.Timestamp(year, month, day)
            # If month is greater than 12, it must be MMDDYYYY
            elif month > 12:
                return pd.Timestamp(year, day, month)
            # If both are 12 or less, default to DDMMYYYY
            else:
                return pd.Timestamp(year, month, day)
        except ValueError:
            error_dict[file_path] = f"Unable to parse date from string: {date_str}"
            return None
    else:
        error_dict[file_path] = f"No valid date found in file name: {file_name}"
    
    error_dict[file_path] = f"No valid date found in file name: {file_name}"
    return None

#################### INPUTS #####################
#INPUT 1
def read_job_list(file_path):
    df = pd.read_excel(file_path,usecols=[0,1,2,3,4])
    df = df.dropna()
    # site code calibration
    df['site_code'] = df['Site code'].astype(str).str.strip()
    df['csm_name'] = df['CSM Name'].astype(str).str.strip()
    return df

#INPUT 2
def read_budget(file_path, error_dict):
    budget_date = get_date(file_path, error_dict)
    if budget_date is None:
        return None
    df = pd.read_csv(file_path) #Get only the first 5 columns relevent
    df['date'] = budget_date
    df['site_code'] = df['SiteCode'].astype(str).str.strip()
    return df

#INPUT 3
def read_payroll(file_path, error_dict):
    payroll_date = get_date(file_path, error_dict)
    if payroll_date is None:
        return None
    df = pd.read_excel(file_path,skiprows=5)
    df['date'] = payroll_date
    df['site_code'] = df['Site'].str.split(' - ').str[0]
    df['pay_type_code'] = df['Pay type'].str.split(' - ').str[0].str.strip()
    return df

#INPUT 4 
def read_workbills(file_path, error_dict):
    workbills_date = get_date(file_path, error_dict)
    if workbills_date is None:
        return None
    df = pd.read_excel(file_path)
    df = df[(df['Workbill hrs'] != 0) | (df['Workbill pay'] != 0)]
    df['date'] = workbills_date
    df['site_code'] = df['Site'].str.split(' - ').str[0]
    return df

