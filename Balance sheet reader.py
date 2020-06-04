import openpyxl, os
import re
import Workbook


wb = openpyxl.load_workbook(r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Marlin Westwood Tenant Balance Sheets/1736 Westwood Tenant Balance Sheets.xlsx')



os.chdir(r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Marlin Westwood Tenant Balance Sheets')
print("cwd = " + os.getcwd())

print(wb.sheetnames)

while True:
    month = input("What month is being checked for the balance? Enter the first three letters of any month.")
    if month[0:3].lower() == 'jan':
        month_index = 1
    elif month[0:3].lower() == 'feb':
        month_index = 2
    elif month[0:3].lower() == 'mar':
        month_index = 3
    elif month[0:3].lower() == 'apr':
        month_index = 4
    elif month[0:3].lower() == 'may':
        month_index = 5
    elif month[0:3].lower() == 'jun':
        month_index = 6
    elif month[0:3].lower() == 'jul':
        month_index = 7
    elif month[0:3].lower() == 'aug':
        month_index = 8
    elif month[0:3].lower() == 'sep':
        month_index = 9
    elif month[0:3].lower() == 'oct':
        month_index = 10
    elif month[0:3].lower() == 'nov':
        month_index = 11
    elif month[0:3].lower() == 'dec':
        month_index = 12
    else:
        print("Invalid entry.")
        continue
    break


# month_index is the month number, 1 through 12
month_regex = re.compile(f'^{month_index}')
print("DEBUG month index = " + str(month_index))

for sheet in wb.sheetnames:
    current_sheet = sheet
