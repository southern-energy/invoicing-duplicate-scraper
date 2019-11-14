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
import xlwt
import xlutils # Module Requires both XLRD & XLWT to be imported.
import pandas as pd

# Find Master Sheet

workbook = xlrd.open_workbook('./2019 Inspections Billing.xlsx')
pd.read_excel('./2019 Inspections Billing.xlsx')

# =======================================================

# The current_sheet variable needs to be named the sheet you want to check for duplicates.

current_sheet = './11-13-19.xlsx'

# =======================================================

# Find All of the Sheets in the Workbook

master_sheet = pd.read_excel('./2019 Inspections Billing.xlsx', sheet_name=None, usecols=[0, 1, 2, 9])

# Combine all sheets of Master Sheet into a single list of lists.

df_master_Street_Address_And_Service = pd.concat(pd.read_excel('./2019 Inspections Billing.xlsx', sheet_name=None, usecols=[2, 9], skiprows=0), sort=True, ignore_index=True)

# Read all of the sheets, using just the columns that have Street Address and the Service.

df_current_sheet_Street_Address_And_Service = pd.concat(pd.read_excel(current_sheet, sheet_name=None, usecols=[2, 9], skiprows=0), sort=True, ignore_index=True)

# First Have to make them into a list of lists (https://stackoverflow.com/questions/22341271/get-list-from-pandas-dataframe-column)

ser_aggRows = pd.Series(df_master_Street_Address_And_Service.values.tolist())


# print('Printing: ser_aggRows (This collapses each row in Excel into a row this script can read.)',ser_aggRows, sep='\n', end='\n\nWe have finished organizing the rows of the Master Workbook\'s sheets into lists.\n\n\n')

ser_aggRows_current_sheet = pd.Series(df_current_sheet_Street_Address_And_Service.values.tolist())

print('Printing: ser_aggRows_current_sheet (This collapses each row in Excel into a row this script can read.)',ser_aggRows_current_sheet, sep='\n', end='\n\nWe have finished organizing the rows of the Current Workbook\'s sheets into lists.\n\n')


first_set = set(map(tuple, ser_aggRows))
secnd_set = set(map(tuple, ser_aggRows_current_sheet))

duplicates = first_set.intersection(secnd_set)

if len(duplicates) > 0:
    print("Duplicates are: ", duplicates, sep="\n", end="\nPlease use these records above to find the duplicates.\n")
else: 
    print("There are no duplicates!")

# Duplicates_list converts the duplicates (an object type: set, with tuples inside it, to a list of lists again)

duplicates_list = list(map(list, duplicates))

print(duplicates_list)
