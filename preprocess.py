import xlwings as xw
import pandas as pd
import re
import pandas as pd
def extract_url(cell_value):
    if cell_value is None:
        return None
    match = re.search(r'"(https?://[^"]+)"', cell_value)
    return match.group(1)
def preprocessfile(filepath):
    raw=pd.read_excel(filepath)
    # Open the Excel file
    wb = xw.Book(filepath)

    # Select the specific sheet you want to work with
    sheet = wb.sheets[0]  # Replace 'Sheet1' with the actual sheet name

    # Create lists to store the data
    cell_addresses = []
    formulas = []

    # Iterate through the cells and extract the formulas
    for cell in sheet.api.UsedRange:
        if cell.HasFormula:
            cell_address = cell.Address
            formula = cell.Formula
        else:
            cell_address = cell.Address
            formula = None

        cell_addresses.append(cell_address)
        formulas.append(formula)

    data = {'Row': [], 'Column': [], 'Formula': []}

    # Split the cell address and populate the data dictionary
    for cell_address, formula in zip(cell_addresses, formulas):
        match = re.match(r'(\$?)([A-Z]+)(\$?)([0-9]+)', cell_address)
        if match:
            row = int(match.group(4))
            column = match.group(2)
            data['Row'].append(row)
            data['Column'].append(column)
            data['Formula'].append(formula)

    # Create the DataFrame
    df = pd.DataFrame(data).pivot(index='Row', columns='Column', values='Formula')

    # Close the Excel file
    wb.close()
    df=df[['C', 'D', 'E', 'F','J']]
    df.columns=['Detail_url', 'Certificate_url', 'Image_url', 'Video_url','Report #_url']


    columns_to_extract = ['Detail_url', 'Certificate_url', 'Image_url', 'Video_url', 'Report #_url']

    for col in columns_to_extract:
        raw[col] = df[col].apply(extract_url)
    return(raw)