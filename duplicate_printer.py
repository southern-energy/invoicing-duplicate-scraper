# # Started by Gregory Power at 11/06/19 @ 4:27 PM
# # Basic Functionality Achieved on 11/12/19
# # Able to Print Indexes on 11/19/19
# For Environment, use Python 3.6.0


# We need to find duplicates that exist, there is one in the master sheet at the bottom of the most recent sheet.


import xlrd
import xlwt
import xlutils # Module Requires both XLRD & XLWT to be imported.
import pandas as pd
import numpy as np

# =======================================================
# =======================================================

# The current_sheet variable needs to be named the sheet you want to check for Duplicate (Service Types &Addresses) OR duplicate (ServiceIDs).

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


#Start of Service ID Section

df_master_ServiceID = pd.concat(pd.read_excel(large_sheet, sheet_name=None, usecols=[9], skiprows=0), sort=False, ignore_index=False)

df_current_sheet_ServiceID = pd.concat(pd.read_excel(current_sheet, sheet_name=None, usecols=[9], skiprows=0), sort=False, ignore_index=False)

ser_aggRows_master_ServiceID = pd.Series(df_master_ServiceID.values.tolist())

ser_aggRows_current_sheet_ServiceID = pd.Series(df_current_sheet_ServiceID.values.tolist())

first_set_ServiceID = set(map(tuple, ser_aggRows_master_ServiceID))

secnd_set_ServiceID = set(map(tuple, ser_aggRows_current_sheet_ServiceID))

second_set_storage_ServiceID = (map(tuple, ser_aggRows_current_sheet_ServiceID))

duplicates_ServiceID = first_set_ServiceID.intersection(secnd_set_ServiceID)

duplicates_list_ServiceID_Duplicates = list(map(list, duplicates_ServiceID))

secnd_set_list_ServiceID_Duplicates = list(map(list, second_set_storage_ServiceID))

if len(duplicates_ServiceID) > 0:
    print("Duplicates for ServiceID Condition are: ", duplicates_ServiceID, sep="\n", end="\nPlease use these records above to find the duplicate ServiceID.\n")
else: 
    print("There are no duplicates for ServiceID condition!")


q = 0
Excel_Indexes_for_ServiceID_Duplicates = []
while q < len(duplicates_list_ServiceID_Duplicates):
    # print(secnd_set_list_Address_Service_Duplicates.index(duplicates_list_ServiceID_Duplicates[q]))
    Excel_Indexes_for_ServiceID_Duplicates.append(int(2) + int(secnd_set_list_ServiceID_Duplicates.index(duplicates_list_ServiceID_Duplicates[q])))
    q += 1
else:
    Excel_Indexes_for_ServiceID_Duplicates.sort()
    print("\nWe are done. Look to the following rows in Excel for Service ID Duplicates: \n", Excel_Indexes_for_ServiceID_Duplicates)

# End of Service ID Section;