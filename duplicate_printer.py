# # Started by Gregory Power at 11/06/19 @ 4:27 PM

# secnd_list_test = [["Not","Me"], ["Dash", "Final"], ["Not_It", "Left"]]
# first_list_test = [["Dash", "Not"], ["Extra", "Data"]]

# first_set_test = set(map(tuple, first_list_test))
# secnd_set_test = set(map(tuple, secnd_list_test))

# # Creates list of duplicates.

# duplicates_test = first_set_test.intersection(secnd_set_test)


# print(first_list_test)
# print(secnd_list_test)
# if len(duplicates_test) > 1:
#     print(duplicates_test)
# else:
#     print("No duplicates have been found.")

# # Above is the basic principle behind this program.


# https://blogs.harvard.edu/rprasad/2014/06/16/reading-excel-with-python-xlrd/

import xlrd
import pandas as pd

# Find Master Sheet

workbook = xlrd.open_workbook('./2019 Inspections Billing.xlsx')
pd.read_excel('./2019 Inspections Billing.xlsx')

# Find All of the Sheets in the Workbook

# print(workbook.sheet_names())
master_sheet = pd.read_excel('./2019 Inspections Billing.xlsx', sheet_name=None, usecols=[0, 1, 2, 9])


# Combine all sheets into a single list of lists.

df_master = pd.concat(pd.read_excel('./2019 Inspections Billing.xlsx', sheet_name=None, usecols=[0, 1, 2, 9], skiprows=0),sort=True, ignore_index=True)

df_master_Street_Address_And_Service = pd.concat(pd.read_excel('./2019 Inspections Billing.xlsx', sheet_name=None, usecols=[2, 9], skiprows=0), sort=True, ignore_index=True)

# Below this line is the current sheet.

df_current_sheet_Street_Address_And_Service = pd.concat(pd.read_excel('./11-7-19.xlsx', sheet_name=None, usecols=[2, 9], skiprows=0), sort=True, ignore_index=True)

# First Have to make them into a list of lists (https://stackoverflow.com/questions/22341271/get-list-from-pandas-dataframe-column)

ser_aggRows = pd.Series(df_master_Street_Address_And_Service.values.tolist())


# print('Printing: ser_aggRows (This collapses each row in Excel into a row this script can read.)',ser_aggRows, sep='\n', end='\n\nWe have finished organizing the rows of the Master Workbook\'s sheets into lists.\n\n\n')

ser_aggRows_current_sheet = pd.Series(df_current_sheet_Street_Address_And_Service.values.tolist())

print('Printing: ser_aggRows_current_sheet (This collapses each row in Excel into a row this script can read.)',ser_aggRows_current_sheet, sep='\n', end='\n\nWe have finished organizing the rows of the Current Workbook\'s sheets into lists.\n\n\n')


first_set = set(map(tuple, ser_aggRows))
secnd_set = set(map(tuple, ser_aggRows_current_sheet))

duplicates = first_set.intersection(secnd_set)

if len(duplicates) < 0:
    print("Duplicates are: \n")
    print("Duplicates are: \n", duplicates, sep="\n", end="Please use these records to find the duplicates.")
    print("\n\n\n\n\nWould you like to export the duplicates list as an Excel File?")
else: 
    print("There are no duplicates!")