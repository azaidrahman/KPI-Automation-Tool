import pandas as pd
import warnings

warnings.simplefilter(action='ignore', category=pd.errors.PerformanceWarning)

start_date = pd.Timestamp(f'{pd.Timestamp.now().year}-1-1')
end_date = pd.Timestamp(f'{pd.Timestamp.now().year}-12-31')

def initialize_main_dataframe(job_list_df, start_date, end_date):
    

    csm_site_index = pd.MultiIndex.from_frame(job_list_df.loc[:,['csm_name','Site','site_code']], names=['csm_name','site_name', 'site_code'])
    

    try:
        # Converting the csm_site_index to lists of tuples
        csm_site_list = list(csm_site_index)

        # Creating a combined MultiIndex from product
        csm_site_index = pd.MultiIndex.from_tuples(
            [(csm, site, code, kpi) for (csm, site, code) in csm_site_list for kpi in ['Planned', 'Actual']],
            names=['csm_name', 'site_name', 'site_code', 'type_of_KPI']
        )
    except Exception as e:
        print(e)
    else:
        print('Index creation successful')


    # Create a date range for the bi-weekly periods
    date_range = pd.date_range(start=start_date, end=end_date, freq='2W-MON')
    date_range = date_range + pd.DateOffset(days=13)

    # Create the column MultiIndex for planned and actual hours/cost
    columns = pd.MultiIndex.from_product([date_range.normalize(), ['Hours','Cost']], names=['biweekly_period', 'metric'])

    # Create the main DataFrame with CSM Name, site code, and bi-weekly periods
    main_df = pd.DataFrame(index=csm_site_index, columns=columns)

    try:
        return main_df
    except:
        print('Error creating main DataFrame. Please check the input files and try again.')
        

def process_budget(df):
    # Check if the DataFrame has daily columns or weekly columns
    days_in_week = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']
    daily_columns = [f'CbHours{day}' for day in days_in_week] + [f'CbValue{day}' for day in days_in_week]
    has_daily_columns = all(col in df.columns for col in daily_columns[:7])

    planned_data = []

    for (site_code,date), site_data in df.groupby(['site_code','date']):
        if has_daily_columns:
            # Calculate weekly budget from daily columns
            planned_hours = sum([site_data[f'CbHours{day}'].sum() for day in days_in_week])
            planned_cost = sum([site_data[f'CbValue{day}'].sum() for day in days_in_week])
        else:
            # Calculate weekly budget from weekly columns
            planned_hours = site_data['HoursWk'].sum()
            planned_cost = site_data['ValueWk'].sum()

        planned_hours *= 2
        planned_cost *= 2

        planned_data.append({'site_code': site_code, 'date': date,'Hours': planned_hours, 'Cost': planned_cost,})

    planned_data = pd.DataFrame(planned_data)
    planned_data.set_index('site_code', inplace=True)

    return planned_data

def process_payroll(df):
    # Filter out unwanted Pay Types
    exclude_pay_types = ['144','145','245','345','IA9', '336','338','134']
    df = df[~df['pay_type_code'].isin(exclude_pay_types)]

    # Group by Site_code and calculate actual hours and actual cost
    actual_data = df.groupby(['site_code','date']).agg({
        'Hours': 'sum',
        'Value': 'sum'
    }).reset_index()

    # Rename columns
    actual_data.columns = ['site_code', 'date' ,'Hours', 'Cost']

    actual_data.set_index('site_code',inplace=True)

    return actual_data

def process_workbills(df):
    workbills_data = df.groupby(['site_code', 'date']).agg({
        'Workbill hrs': 'sum',
        'Workbill pay': 'sum'
    }).reset_index()
    workbills_data.columns = ['site_code', 'date', 'Hours', 'Cost']
    workbills_data.set_index('site_code', inplace=True)
    return workbills_data

def process_report(type,job_list_df,data_df, main_df):
    # Create a bi-weekly period column based on the 'date' column

    # print(data_df['date'][0])
    biweekly_period_index = ((data_df['date'] - start_date).dt.days // 14) + 1
    number_of_sub_columns = 2
    biweekly_period_date = main_df.columns[biweekly_period_index*number_of_sub_columns-1][0][0]
    # Iterate over each row in the data DataFrame
    for index, row in data_df.iterrows():
        site_code = index
        try:
            # Extract csm_name from job_list_df where site_code matches
            csm_name = job_list_df[job_list_df['site_code'] == site_code]['csm_name'].values[0]
            site_name = job_list_df[job_list_df['site_code'] == site_code]['Site'].values[0]
        except IndexError:
            # If no matching site_code is found, skip this iteration
            continue
        
        # curr_hours = row['Hours']
        # curr_cost = row['Cost']
        # print(site_code)
        # print('before')
        # print(main_df.loc[(csm_name, site_name, site_code, 'Planned'), ( biweekly_period_date, 'Hours')])
        # Update the main DataFrame with the corresponding values
        if type == 'budget':
            main_df.loc[(csm_name, site_name, site_code, 'Planned'), ( biweekly_period_date, 'Hours')] = row['Hours']
            main_df.loc[(csm_name, site_name, site_code, 'Planned'), ( biweekly_period_date, 'Cost')] = row['Cost']
        elif type == 'payroll':
            main_df.loc[(csm_name, site_name, site_code, 'Actual'), ( biweekly_period_date, 'Hours')] = row['Hours']
            main_df.loc[(csm_name, site_name, site_code, 'Actual'), ( biweekly_period_date, 'Cost')] = row['Cost']
        elif type == 'workbills':
            # Subtract Workbills data from Actual
            current_hours = main_df.loc[(csm_name, site_name, site_code, 'Actual'), (biweekly_period_date, 'Hours')]
            current_cost = main_df.loc[(csm_name, site_name, site_code, 'Actual'), (biweekly_period_date, 'Cost')]
            main_df.loc[(csm_name, site_name, site_code, 'Actual'), (biweekly_period_date, 'Hours')] = current_hours - row['Hours']
            main_df.loc[(csm_name, site_name, site_code, 'Actual'), (biweekly_period_date, 'Cost')] = current_cost - row['Cost']

    return main_df

    
def remove_empty_rows_and_columns(df):
    # Remove rows where all values are NaN for both 'Planned' and 'Actual'
    mask = df.groupby(level=['site_name', 'site_code']).transform(lambda x: x.notna().any())
    df = df[mask.any(axis=1)]

    # Remove columns (date pairs) where all values are NaN
    columns_to_keep = []
    for i in range(0, len(df.columns), 2):
        if not df.iloc[:, i:i+2].isna().all().all():
            columns_to_keep.extend([i, i+1])
    
    df = df.iloc[:, columns_to_keep]

    return df