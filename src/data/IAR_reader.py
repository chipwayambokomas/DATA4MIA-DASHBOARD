import pandas as pd
##from db_connection import get_db_connection

#### Read database

data = pd.read_excel("../../data/raw/Copy_of_Emalahleni_Infrastructure_2023.xlsx",engine='openpyxl')

#### Filtering database to retreive desired columns as specified by standard accounting rules

desired_columns = ['Organisation', 'Organisational ID', 'Number', 'Model', 'Serial Number', 'Description', 'Material Type', 'Asset Class ID', 'Sector ID', 'Size', 'Size Measurement (unit)', 'Capacity', 'Quantity', 'Utilisation', 'Replacement Equivalent', 'Year Constructed / Purchase Date', 'Supplier Name', 'Criticality Grade ID', 'Condition ID', 'Location X', 'Location Y', 'Custodian', 'Replacement Cost Per Item', 'Purchase Price', 'Actual Deemed Cost', 'Depreciated Replacement Cost', 'Current Replacement Cost', 'Additional Amount', 'EUL in Years', 'Age', 'RUL in Years']
 
filtered_columns = []

for col in desired_columns:
    if col in data.columns:
        filtered_columns.append(col)
        
        
filtered_data = data[filtered_columns]

##cnx = get_db_connection()

#### Checking completeness of the Munincipality IAR ~ out of the 31 required fields, how many do we have?

num_columns = filtered_data.shape[1]

completeness = (num_columns/31)*100

string_completeness = "{:.2f}".format(completeness)

print(string_completeness)

#### Checking completeness of individual categories within a munincipality

for col in filtered_data:
    
    nan_count = filtered_data[col].isna().sum()
    
    str_nan_count = str(nan_count)
    print(nan_count)

    cat_completeness = ((filtered_data.shape[0]-nan_count)/filtered_data.shape[0])*100

    string_cat_completeness = "{:.2f}".format(cat_completeness)

    print(string_cat_completeness)

    




