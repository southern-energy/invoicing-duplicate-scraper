# # Started by Gregory Power at 11/06/110 @ 4:27 PM
# # Basic Functionality Achieved on 11/12/110
# # Able to Print Indexes on 11/110/110
# For Environment, use Python 3.6.0
# Began Creation of New Branch for New Excel format that includes Service IDs

import xlrd
import xlwt
import xlutils # Module Requires both XLRD & XLWT to be imported.
import pandas as pd
import numpy as np

# =======================================================
# =======================================================

# The current_sheet variable needs to be named the sheet you want to check for duplicates.

current_sheet = './dummySheetnewformat.xlsx'

large_sheet = './2019 Inspections Billing_New_Format.xlsx'

# =======================================================
# =======================================================
# =======================================================
# =======================================================
# =======================================================

# Find Master Sheet

workbook = xlrd.open_workbook(large_sheet)
pd.read_excel(large_sheet)



# Find All of the Sheets in the Workbook

# master_sheet = pd.read_excel(large_sheet, sheet_name=None, usecols=[0, 1, 2, 10])

# Combine all sheets of Master Sheet into a single list of lists.

df_master_Street_Address_And_Service = pd.concat(pd.read_excel(large_sheet, sheet_name=None, usecols=[2, 10], skiprows=0), sort=False, ignore_index=False)

print(df_master_Street_Address_And_Service)

# Read all of the sheets, using just the columns that have Street Address and the Service.

df_current_sheet_Street_Address_And_Service = pd.concat(pd.read_excel(current_sheet, sheet_name=None, usecols=[2, 10], skiprows=0), sort=False, ignore_index=False)

print(df_current_sheet_Street_Address_And_Service)

# First Have to make them into a list of lists (https://stackoverflow.com/questions/22341271/get-list-from-pandas-dataframe-column)

ser_aggRows = pd.Series(df_master_Street_Address_And_Service.values.tolist())


# print('Printing: ser_aggRows (This collapses each row in Excel into a row this script can read.)',ser_aggRows, sep='\n', end='\n\nWe have finished organizing the rows of the Master Workbook\'s sheets into lists.\n\n\n')

ser_aggRows_current_sheet = pd.Series(df_current_sheet_Street_Address_And_Service.values.tolist())

# print('Printing: ser_aggRows_current_sheet (This collapses each row in Excel into a row this script can read.)',ser_aggRows_current_sheet, sep='\n', end='\n\nWe have finished organizing the rows of the Current Workbook\'s sheets into lists.\n\n')


first_set = set(map(tuple, ser_aggRows))
secnd_set = set(map(tuple, ser_aggRows_current_sheet))
second_set_storage = (map(tuple, ser_aggRows_current_sheet))


duplicates = first_set.intersection(secnd_set)

if len(duplicates) > 0:
    print("Duplicates are: ", duplicates, sep="\n", end="\nPlease use these records above to find the duplicates.\n")
else: 
    print("There are no duplicates!")

# Duplicates_list converts the duplicates (an object type: set, with tuples inside it, to a list of lists again)

duplicates_list = list(map(list, duplicates))

secnd_set_list = list(map(list, second_set_storage))

# print(duplicates_list)
# print("The Length of the Second Set List: " + str(len(secnd_set_list)))
# print("The Length of the Aggrows Sheet: " + str(len(ser_aggRows_current_sheet)))


k = 0
Excel_Indexes = []
while k < len(duplicates_list):
    # print(secnd_set_list.index(duplicates_list[k]))
    Excel_Indexes.append(int(2) + int(secnd_set_list.index(duplicates_list[k])))
    k += 1
else:
    Excel_Indexes.sort()
    print("\nWe are done. Look to the following rows in Excel: \n", Excel_Indexes, sep="\n")



