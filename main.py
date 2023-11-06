import os
import pandas as pd
from configparser import ConfigParser

from preprocess import *
# Get the absolute path to the config.ini file
config_file_path = os.path.join(os.path.dirname(__file__), 'config.ini')

# Load the configuration from config.ini
config = ConfigParser()
config.read(config_file_path)

# Get configuration values
unprocessed_file_path = config.get('Excel', 'unprocessed_file_path')
processed_file_path = config.get('Excel', 'processed_file_path')
output_vendor_file_path = config.get('Excel', 'output_vendor_file_path')
output_agent_file_path = config.get('Excel', 'output_agent_file_path')
increment_percentage = float(config.get('Excel', 'increment_percentage'))
flat_dollars = float(config.get('Excel', 'flat_dollars'))
keyword_map = eval(config.get('Excel', 'keyword_map'))
#print(unprocessed_file_path)
#print(os.listdir(unprocessed_file_path))
# Step 1: Get all the files from unprocessed folder
for filename in os.listdir(unprocessed_file_path):
    file_path = os.path.join(unprocessed_file_path, filename)

    # Step 2: Read the data from the file
    unprocessed_file = preprocessfile(file_path)

    # Step 3: Replace values in 'Polish' and 'Symmetry' columns
    unprocessed_file['Polish'].replace(keyword_map, inplace=True)
    unprocessed_file['Symmetry'].replace(keyword_map, inplace=True)
    unprocessed_file['Shape'].replace(keyword_map, inplace=True)
    # Step 4: Calculate 'Final Price' column
    # Step 4: Calculate 'Final Price' column
    unprocessed_file['Amount'] = unprocessed_file['Amount'].str.replace('[\$,]', '', regex=True).astype(float)
    unprocessed_file['Final Price'] = unprocessed_file['Amount'] * (1 + increment_percentage / 100) + flat_dollars

    # Round 'Amount' and 'Final Price' to the nearest whole number
    unprocessed_file['Amount'] = unprocessed_file['Amount'].round()
    unprocessed_file['Final Price'] = unprocessed_file['Final Price'].round()
    unprocessed_file= unprocessed_file[['Sr.No','Stone Id','Location','Lab','Report #','Shape','Weight','Col','Clarity','Cut','Polish','Symmetry','Fluor','Amount','Final Price','Status','Certificate_url','Image_url','Video_url']]
    unprocessed_file = unprocessed_file.dropna(subset=['Image_url', 'Video_url'], how='any')
    unprocessed_file = unprocessed_file[~((unprocessed_file['Image_url'] == '-') | (unprocessed_file['Video_url'] == '-'))]

    # Step 5: Rename columns
    column_rename_dict = {
        'Col': 'Color',
        'Weight': 'Carat',
        'Fluor': 'Fluorescence'
    }
    unprocessed_file = unprocessed_file.rename(columns=column_rename_dict)
    unprocessed_file['Sr.No'] = range(1, len(unprocessed_file) + 1)
    # Step 6: Save to the output_agent_file_path
    output_file_path = os.path.join(output_agent_file_path, "agent_"+filename)
    #print(unprocessed_file.head())
    # Save the DataFrame to an Excel file
    with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
        unprocessed_file.to_excel(writer, sheet_name='Sheet1', index=False)

    # Step 7: Pick specific columns and save in output_agent_file_path
    columns_to_pick = ['Sr.No','Lab', 'Shape','Carat', 'Color', 'Clarity', 'Cut', 'Polish','Symmetry', 'Fluorescence', 'Final Price','Certificate_url', 'Image_url', 'Video_url']
    selected_columns_file = unprocessed_file[columns_to_pick]
    output_file_path = os.path.join(output_vendor_file_path,"vendor_"+filename)
    

    # Save the DataFrame to an Excel file
    with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
        selected_columns_file.to_excel(writer, sheet_name='Sheet1', index=False)

    # Step 7: Move the file from unprocessed to processed
    os.rename(file_path, os.path.join(processed_file_path, filename))

