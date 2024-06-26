import pandas as pd

desired_columns = ['Organisation', 'Organisational ID', 'Number', 'Model', 'Serial Number', 'Description', 'Material Type', 'Asset Class ID', 'Sector ID', 'Size', 'Size Measurement (unit)', 'Capacity', 'Quantity', 'Utilisation', 'Replacement Equivalent', 'Year Constructed / Purchase Date', 'Supplier Name', 'Criticality Grade ID', 'Condition ID', 'Location X', 'Location Y', 'Custodian', 'Replacement Cost Per Item', 'Purchase Price', 'Actual Deemed Cost', 'Depreciated Replacement Cost', 'Current Replacement Cost', 'Additional Amount', 'EUL in Years', 'Age', 'RUL in Years'] 

data = pd.read_excel("C:/Users/NEW ERA/Documents/DATA4MIA/DATA4MIA-DASHBOARD/data/raw/Copy_of_Emalahleni_Infrastructure_2023.xlsx")

filtered_columns = []

for col in desired_columns:
    if col in data.columns:
        filtered_columns.append(col)
        
        
filtered_data = data[filtered_columns]

