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

# Started by Gregory Power at 11/06/19 @ 4:27 PM
# https://blogs.harvard.edu/rprasad/2014/06/16/reading-excel-with-python-xlrd/

import xlrd

# Pipelines

master_list = []
comparison_list = []

# Find Master Sheet

workbook = xlrd.open_workbook('./2019 Inspections Billing.xlsx')

# Find All of the Sheets in the Workbook

print(workbook.sheet_names())

# With Sheet Names, it returns a list of all of the sheets, in the file. I can use the length of the list to iterate through the Master file.

print(len(workbook.sheet_names()))

# Combine them all into a Single Array

print(workbook.sheets())

# Compare to Current Sheet

