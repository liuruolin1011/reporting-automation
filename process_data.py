import os
import numpy as np
import pandas as pd
import datetime
from tools import get_file, extract_branch, extract_restriction_code, extract_sector_code, extract_sector_description, extract_segment, segment_type, offboarding_reason, pep_type

pd.options.mode.chained_assignment = None  # Disable warning about chained assignments

def process_data(previous_month, previous_year, current_month, current_year, input_folder_path, output_folder_path):
    try:
        file_name_1 = f"Compliance Reporting Metrics_{previous_month} {previous_year}.xlsx"
        file_name_2 = f"Restricted Customers_{previous_month} {previous_year}.xlsx"
        
        file_path_1 = get_file(file_name_1, input_folder_path)
        file_path_2 = get_file(file_name_2, input_folder_path)

        output_file_name = f"Heatmap_{current_month}_{current_year}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        output_file_path = os.path.join(output_folder_path, output_file_name)
        
        # Load data from the input Excel files
        prev_active_cust = pd.read_excel(file_path_1, sheet_name='4.1 Customer List', header=1)
        prev_PEP_list = pd.read_excel(file_path_1, sheet_name='4.2 PEP', header=1)
        prev_restricted_cust = pd.read_excel(file_path_2)
        
        # Loading additional raw data
        raw_data_RTO_filename = "BaseQueuesReporting_V3.xlsx"
        raw_data_RTO = pd.read_excel(get_file(raw_data_RTO_filename, input_folder_path))
        
        raw_data_CDO_filename = "CDO_data.xlsx"
        raw_data_CDO = pd.read_excel(get_file(raw_data_CDO_filename, input_folder_path))
        
        raw_data_T24_filename = "T24_data.xlsx"
        raw_data_T24 = pd.read_excel(get_file(raw_data_T24_filename, input_folder_path))
        
        status_reporting_filename = "StatusHistoryReporting.xlsx"
        status_reporting_dfs = pd.read_excel(get_file(status_reporting_filename, input_folder_path), sheet_name=None, header=None)
        status_reporting = pd.concat(status_reporting_dfs, ignore_index=True)
        
        print(f"All necessary files loaded successfully!")
    except FileNotFoundError as e:
        raise FileNotFoundError(f"Required file not found: {str(e)}")
    
    # Data Transformation and Extraction
    raw_data_RTO['Segment'] = raw_data_RTO['LE Type'].apply(extract_segment)
    raw_data_RTO['Branch'] = raw_data_RTO['Branch'].apply(extract_branch)
    raw_data_RTO['RTO Restriction Status'] = raw_data_RTO['Restriction Status'].apply(extract_restriction_code).fillna('NO RESTRICTION')
    raw_data_RTO['Sector Description'] = raw_data_RTO['Sector Code'].apply(extract_sector_description)
    raw_data_RTO['Sector Code'] = raw_data_RTO['Sector Code'].apply(extract_sector_code)
    raw_data_RTO['RTO Risk Rating'] = raw_data_RTO['Overall Risk']
    raw_data_RTO['Negative News Indicator'] = raw_data_RTO['Negative News (Material, Relevant)']
    raw_data_RTO['T24 ID / Customer No.'] = raw_data_RTO['BoC Client ID']
    raw_data_RTO['PEP TYPE'] = raw_data_RTO.apply(pep_type, axis=1)
    raw_data_RTO['Segment'] = raw_data_RTO.apply(segment_type, axis=1)
    
    # Continue your data processing logic here...
    # Filtering and processing the data based on the requirements

    # Prepare output Excel file with multiple sheets
    with pd.ExcelWriter(output_file_path) as writer:
        raw_data_RTO.to_excel(writer, sheet_name='All Customers', index=False)
        # Write additional processed dataframes into separate sheets
        # current_active_cust.to_excel(writer, sheet_name='Active Customer List', index=False)
        # Add more dataframes as needed...

    print(f"Data processed and written to {output_file_name}")
    return output_file_path, removal_list, current_restricted_cust, current_active_cust, additional_cust, negative_news, total_offboard_cust

