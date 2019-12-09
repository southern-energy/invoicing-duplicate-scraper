# Duplicate Sniffer Project: Bloodhound

## Steps

- ✔ Organize Pipelines
- ✔ Master Sheet
- ✔ Current DataSheet to Query From
- ✔ Make Each Dataframe into a List of Lists for Each Sheet
- ✔ Compare the two lists of lists for duplicates
- ✔ Create list of lists for duplicates
- ✔ Print that duplicates list to user
- (Optional) Edit the "Current DataSheet" by highlighting the indexes that correspond to the duplicates list
- ✔ Use the "index()" method to compare the second_set list and duplicates list
- (New Addition: 12/09/19) Original requirements to be flagged as a duplicate were:
  - If "Address" and "Service Type" column matched.
    - Address was col 2
    - ServiceType was col 9
    - ServiceID was non-existant
  - New Criteria "ServiceID" added to Reports.
    - Address now col 2
    - ServiceID now col 9
    - ServiceType Now col 10

## Outputs

- Prints the duplicates out to the user.
- Print the rows where you can find duplicates in the Excel file.

## Resources

- For Editing Python Styles:
  - <https://xlutils.readthedocs.io/en/latest/styles.html>
  - <http://www.python-excel.org/>
  