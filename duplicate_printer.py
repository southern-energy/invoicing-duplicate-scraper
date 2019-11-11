# Started by Gregory Power at 11/06/19 @ 4:27 PM

# secnd_list = [["Not","Me"], ["Dash", "Final"], ["Not_It", "Left"]]
# first_list = [["Dash", "Final"], ["Extra", "Data"]]

# first_set = set(map(tuple, first_list))
# secnd_set = set(map(tuple, secnd_list))

# # Creates list of duplicates.

# duplicates = first_set.intersection(secnd_set)


# #print(first_list)
# #print(secnd_list)
# print(duplicates)

# Above is the basic principle behind this program.


# https://blogs.harvard.edu/rprasad/2014/06/16/reading-excel-with-python-xlrd/

import xlrd
import pandas as pd

# Pipelines

master_list = []
comparison_list = []

# Find Master Sheet

workbook = xlrd.open_workbook('./2019 Inspections Billing.xlsx')
pd.read_excel('./2019 Inspections Billing.xlsx')

# Find All of the Sheets in the Workbook

# print(workbook.sheet_names())
master_sheet = pd.read_excel('./2019 Inspections Billing.xlsx', sheet_name=None, usecols=[0, 1, 2, 9])


# With Sheet Names, it returns a list of all of the sheets, in the file. I can use the length of the list to iterate through the Master file.

print(len(workbook.sheet_names()))

# Combine them all into a Single Array

df_master = pd.concat(pd.read_excel('./2019 Inspections Billing.xlsx', sheet_name=None, usecols=[0, 1, 2, 9], skiprows=0),sort=True, ignore_index=True)
df_master_Street_Address_And_Service = pd.concat(pd.read_excel('./2019 Inspections Billing.xlsx', sheet_name=None, usecols=[2, 9], skiprows=0), sort=True, ignore_index=True)

# print(df_master_Street_Address_And_Service)

# Compare to Current Sheet

# First Have to make them into a list of lists (https://stackoverflow.com/questions/22341271/get-list-from-pandas-dataframe-column)

ser_aggRows = pd.Series(df_master_Street_Address_And_Service.values.tolist())

print('ser_aggRows (This collapses each row in Excel into a row this script can read.)',ser_aggRows, sep='\n', end='\n\nWe have finished organizing the rows of the Master Workbook\'s sheets into lists.\n\n\n')

# first_set = set(map(tuple, ser_aggRows))
# secnd_set = set(map(tuple, ser_aggRows))

# duplicates = first_set.intersection(secnd_set)

# print("Duplicates are: ")

# print(duplicates)