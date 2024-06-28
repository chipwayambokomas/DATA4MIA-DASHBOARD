# DATA4MIA's IAR reader
# Saul Chipwayambokoma - chipwayambokomas@outlook.com
# 1.0.0 June 28, 2024

import pandas as pd
import openpyxl
##from db_connection import get_db_connection

#### Read database

## use if using interactive window
#data = pd.read_excel("../../data/raw/ASSETS_REGISTER_2021_2022_FAR_DLM_Draft_ver_4_2.xlsb", sheet_name= 3)

data = pd.read_excel("data/raw/Copy_of_Emalahleni_Infrastructure_2023.xlsx")

#### Filtering database to retreive desired columns as specified by standard accounting rules

desired_columns = ['Organisation', 'Organisational ID', 'Number', 'Model', 'Serial Number', 'Description', 'Material Type', 'Asset Class ID', 'Sector ID', 'Size', 'Size Measurement (unit)', 'Capacity', 'Quantity', 'Utilisation', 'Replacement Equivalent', 'Year Constructed / Purchase Date', 'Supplier Name', 'Criticality Grade ID', 'Condition ID', 'Location X', 'Location Y', 'Custodian', 'Replacement Cost Per Item', 'Purchase Price', 'Actual Deemed Cost', 'Depreciated Replacement Cost', 'Current Replacement Cost', 'Additional Amount', 'EUL in Years', 'Age', 'RUL in Years']
 
filtered_columns = []

for col in desired_columns:
    if col in data.columns:
        filtered_columns.append(col)
        
        
filtered_data = data[filtered_columns]

##cnx = get_db_connection()

#### Checking completeness of the Munincipality IAR ~ out of the 31 required fields, how many do we have?

Organisation = Organisational_ID = Number = Model = Serial_Number = Description = Material_Type = Asset_Class_ID = Sector_ID = Size = Size_Measurement = Capacity = Quantity = Utilisation = Replacement_Equivalent = Year_Constructed_Purchase_Date = Supplier_Name = Criticality_Grade_ID = Condition_ID = Location_X = Location_Y = Custodian = Replacement_Cost_Per_Item = Purchase_Price = Actual_Deemed_Cost = Depreciated_Replacement_Cost = Current_Replacement_Cost = Additional_Amount = EUL_in_Years = Age = RUL_in_Years = 0

num_columns = filtered_data.shape[1]

overall_completeness = (num_columns/31)*100


#### function to calculate completeness of categoriy within a munincipality IAR

def populate_cat(col):
    nan_count = filtered_data[col].isna().sum()

    cat_completeness = ((filtered_data.shape[0]-nan_count)/filtered_data.shape[0])*100
    
    return cat_completeness

#### Populate xlxs completeness sheet with values

my_list = [0] * 32 

my_list[1] = overall_completeness  

for col in filtered_data:
    if col == "Organisation":
        Organisation = populate_cat(col)
        my_list[0] = Organisation
    elif col == "Organisational ID":
        Organisational_ID = populate_cat(col)
        my_list[2] = Organisational_ID
    elif col == "Number":
        Number = populate_cat(col)
        my_list[3] = Number
    elif col == "Model":
        Model = populate_cat(col)
        my_list[4] = Model
    elif col == "Serial Number":
        Serial_Number = populate_cat(col)
        my_list[5] = Serial_Number
    elif col == "Description":
        Description = populate_cat(col)
        my_list[6] = Description
    elif col == 'Material Type':
        Material_Type = populate_cat(col)
        my_list[7] = Material_Type
    elif col == "Asset Class ID":
        Asset_Class_ID = populate_cat(col)
        my_list[8] = Asset_Class_ID
    elif col == "Sector ID":
        Sector_ID = populate_cat(col)
        my_list[9] = Sector_ID
    elif col == "Size":
        Size = populate_cat(col)
        my_list[10] = Size
    elif col == "Size Measurement (unit)":
        Size_Measurement = populate_cat(col)
        my_list[11] = Size_Measurement
    elif col == "Capacity":
        Capacity = populate_cat(col)
        my_list[12] = Capacity
    elif col == "Quantity":
        Quantity = populate_cat(col)
        my_list[13] = Quantity
    elif col == "Utilisation":
        Utilisation = populate_cat(col)
        my_list[14] = Utilisation
    elif col == "Replacement Equivalent":
        Replacement_Equivalent = populate_cat(col)
        my_list[15] = Replacement_Equivalent
    elif col == "Year Constructed / Purchase Date":
        Year_Constructed_Purchase_Date = populate_cat(col)
        my_list[16] = Year_Constructed_Purchase_Date
    elif col == "Supplier Name":
        Supplier_Name = populate_cat(col)
        my_list[17] = Supplier_Name
    elif col == "Criticality Grade ID":
        Criticality_Grade_ID = populate_cat(col)
        my_list[18] = Criticality_Grade_ID
    elif col == "Condition ID":
        Condition_ID = populate_cat(col)
        my_list[19] = Condition_ID
    elif col == "Location X":
        Location_X = populate_cat(col)
        my_list[20] = Location_X
    elif col == "Location Y":
        Location_Y = populate_cat(col)
        my_list[21] = Location_Y
    elif col == "Custodian":
        Custodian = populate_cat(col)
        my_list[22] = Custodian
    elif col == "Replacement Cost Per Item":
        Replacement_Cost_Per_Item = populate_cat(col)
        my_list[23] = Replacement_Cost_Per_Item
    elif col == "Purchase Price":
        Purchase_Price = populate_cat(col)
        my_list[24] = Purchase_Price
    elif col == "Actual Deemed Cost":
        Actual_Deemed_Cost = populate_cat(col)
        my_list[25] = Actual_Deemed_Cost
    elif col == "Depreciated Replacement Cost":
        Depreciated_Replacement_Cost = populate_cat(col)
        my_list[26] = Depreciated_Replacement_Cost
    elif col == "Current Replacement Cost":
        Current_Replacement_Cost = populate_cat(col)
        my_list[27] = Current_Replacement_Cost
    elif col == "Additional Amount":
        Additional_Amount = populate_cat(col)
        my_list[28] = Additional_Amount
    elif col == "EUL in Years":
        EUL_in_Years = populate_cat(col)
        my_list[29] = EUL_in_Years
    elif col == "Age":
        Age = populate_cat(col)
        my_list[30] = Age
    elif col == "RUL in Years":
        RUL_in_Years = populate_cat(col)
        my_list[31] = RUL_in_Years

#### Populate completeness dashboard

#Use for interactive window
#wb = openpyxl.load_workbook("../../dashboards/completeness_dashboard.xlsx")

wb = openpyxl.load_workbook("dashboards/completeness_dashboard.xlsx")

sheet_name = 'Sheet1'
ws = wb[sheet_name]

ws.append(my_list)

#Use for interactive window
#wb.save("../../dashboards/completeness_dashboard.xlsx")

wb.save("dashboards/completeness_dashboard.xlsx")


        
        
    
        
        
        



