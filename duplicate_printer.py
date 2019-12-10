# # Started by Gregory Power at 11/06/19 @ 4:27 PM
# # Basic Functionality Achieved on 11/12/19
# # Able to Print Indexes on 11/19/19
# For Environment, use Python 3.6.0

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

# Find All of the Sheets in the Workbook (12/09/19; This version looks for Index, DASHID, Address, ServiceType) 

master_sheet = pd.read_excel(large_sheet, sheet_name=None, usecols=[0, 1, 2, 9, 10])


# Combine all sheets of Master Sheet into a single list of lists.

df_master_Street_Address_And_ServiceType = pd.concat(pd.read_excel(large_sheet, sheet_name=None, usecols=[2, 10], skiprows=0), sort=False, ignore_index=False)



# Read all of the sheets, using just the columns that have Street Address and the Service.

df_current_sheet_Street_Address_And_ServiceType = pd.concat(pd.read_excel(current_sheet, sheet_name=None, usecols=[2, 10], skiprows=0), sort=False, ignore_index=False)




# First Have to make them into a list of lists (https://stackoverflow.com/questions/22341271/get-list-from-pandas-dataframe-column)

ser_aggRows_Street_Address_And_ServiceType = pd.Series(df_master_Street_Address_And_ServiceType.values.tolist())

print('Printing: ser_aggRows_Street_Address_And_ServiceType (This collapses each row in Excel into a row this script can read.)',ser_aggRows_Street_Address_And_ServiceType, sep='\n', end='\n\nWe have finished organizing the rows of the Master Workbook\'s sheets into lists.\n\n\n')



ser_aggRows_current_sheet_Street_Address_And_ServiceType = pd.Series(df_current_sheet_Street_Address_And_ServiceType.values.tolist())

# print('Printing: ser_aggRows_current_sheet_Street_Address_And_ServiceType (This collapses each row in Excel into a row this script can read.)',ser_aggRows_current_sheet_Street_Address_And_ServiceType, sep='\n', end='\n\nWe have finished organizing the rows of the Current Workbook\'s sheets into lists.\n\n')



first_set_Street_Address_And_ServiceType = set(map(tuple, ser_aggRows_Street_Address_And_ServiceType))



secnd_set_Street_Address_And_ServiceType = set(map(tuple, ser_aggRows_current_sheet_Street_Address_And_ServiceType))



second_set_storage_Street_Address_And_ServiceType = (map(tuple, ser_aggRows_current_sheet_Street_Address_And_ServiceType))



duplicates_Street_Address_And_ServiceType = first_set_Street_Address_And_ServiceType.intersection(secnd_set_Street_Address_And_ServiceType)



if len(duplicates_Street_Address_And_ServiceType) > 0:
    print("Duplicates for Address & Cervice Type Condition are: ", duplicates_Street_Address_And_ServiceType, sep="\n", end="\nPlease use these records above to find the duplicate Street Address and ServiceType.\n")
else: 
    print("There are no duplicates for Address & Service Type Condition!")



# Duplicates_list converts the duplicates_Street_Address_And_ServiceType (an object type: set, with tuples inside it, to a list of lists again)

duplicates_list_Address_Service_Duplicates = list(map(list, duplicates_Street_Address_And_ServiceType))

secnd_set_list_Address_Service_Duplicates = list(map(list, second_set_storage_Street_Address_And_ServiceType))

# print(duplicates_list_Address_Service_Duplicates)
# print("The Length of the Second Set List: " + str(len(secnd_set_list_Address_Service_Duplicates)))
# print("The Length of the Aggrows Sheet: " + str(len(ser_aggRows_current_sheet_Street_Address_And_ServiceType)))


k = 0
Excel_Indexes_for_Address_Service_Type_Duplicates = []
while k < len(duplicates_list_Address_Service_Duplicates):
    # print(secnd_set_list_Address_Service_Duplicates.index(duplicates_list_Address_Service_Duplicates[k]))
    Excel_Indexes_for_Address_Service_Type_Duplicates.append(int(2) + int(secnd_set_list_Address_Service_Duplicates.index(duplicates_list_Address_Service_Duplicates[k])))
    k += 1
else:
    Excel_Indexes_for_Address_Service_Type_Duplicates.sort()
    print("\nWe are done. Look to the following rows in Excel for Address & Service Type Duplicates: \n", Excel_Indexes_for_Address_Service_Type_Duplicates, sep="\n")




# Start of ServiceID Section:

df_master_ServiceID = pd.concat(pd.read_excel(large_sheet, sheet_name=None, usecols=[9], skiprows=0), sort=False, ignore_index=False)

# End of ServiceID Section;
# Start of ServiceID Section:

df_current_sheet_ServiceID = pd.concat(pd.read_excel(current_sheet, sheet_name=None, usecols=[9], skiprows=0), sort=False, ignore_index=False)

# End of ServiceID Section;
# Start of ServiceID Section: We are aggregating rows of ServiceIDs into a list for the master sheet.

ser_aggRows_master_ServiceID = pd.Series(df_master_ServiceID.values.tolist())

# End of ServiceID Section;
# Start of ServiceID Section: We are aggregating rows of ServiceIDs into a list for the current sheet.

ser_aggRows_current_sheet_ServiceID = pd.Series(df_current_sheet_ServiceID.values.tolist())

# End of ServiceID Section;
# Start of ServiceID:

first_set_ServiceID = set(map(tuple, ser_aggRows_current_sheet_ServiceID))

# End of ServiceID Section;
# Start of Service ID Section:

secnd_set_ServiceID = set(map(tuple, ser_aggRows_current_sheet_ServiceID))

# End of ServiceID Section;
# Start of Service ID Section:

second_set_storage_ServiceID = (map(tuple, ser_aggRows_current_sheet_ServiceID))

# End of ServiceID Section;
# Start of Service ID Section:

duplicates_ServiceID = first_set_ServiceID.intersection(secnd_set_ServiceID)

# End Of ServiceID Section;
#Start of Service ID Section:

duplicates_list_ServiceID_Duplicates = list(map(list, duplicates_ServiceID))

secnd_set_list_ServiceID_Duplicates = list(map(list, second_set_storage_ServiceID))

# End of Service ID Section;

# Start of Service ID Section:

if len(duplicates_ServiceID) > 0:
    print("Duplicates for ServiceID Condition are: ", duplicates_ServiceID, sep="\n", end="\nPlease use these records above to find the duplicate ServiceID.\n")
else: 
    print("There are no duplicates for ServiceID condition!")

# End of Service ID Section;

# Start of Service ID Section:

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