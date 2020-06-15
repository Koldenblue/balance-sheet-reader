import openpyxl
import os
import re
from Workbook_class import Workbook
from Month_class import Month
from colorama import init, Fore, Back, Style
from termcolor import colored
init()

wb = openpyxl.load_workbook(r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Marlin Westwood Tenant Balance Sheets/1736 Westwood Tenant Balance Sheets.xlsx', data_only=True)
os.chdir(r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Marlin Westwood Tenant Balance Sheets')
# print(wb.sheetnames)

# 2d Array where wb_list[i][0] == filename, and wb_list[i][1] == working directory
# wb_list[i][2] is the name of the corresponding property.
# wb_list[i][3] is the name of the corresponding company.
# The file locations may have to be changed per year.
wb_list = [
    [r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Marlin Westwood Tenant Balance Sheets/1736 Westwood Tenant Balance Sheets.xlsx',
    r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Marlin Westwood Tenant Balance Sheets', "1736 Westwood", "Marlin Westwood"],

    [r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Marlin Westwood Tenant Balance Sheets/1740 Westwood Tenant Balance Sheet.xlsx',
    r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Marlin Westwood Tenant Balance Sheets', "1740 Westwood", "Marlin Westwood"],

    [r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Marlin Westwood Tenant Balance Sheets/1750 Westwood Tenant Balance Sheets.xlsx',
    r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Marlin Westwood Tenant Balance Sheets', "1750 Westwood", "Marlin Westwood"],

    [r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Marlin Westwood Tenant Balance Sheets/1760 Westwood Tenant Balance Sheets.xlsx',
    r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Marlin Westwood Tenant Balance Sheets', "1760 Westwood", "Marlin Westwood"],

    [r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Marlin Westwood Tenant Balance Sheets/MW Hilts 2020 Tenant Balance Sheets.xlsx',
    r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Marlin Westwood Tenant Balance Sheets', "1624 Hilts", "Marlin Westwood"],

    [r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Twinwood Tenant Balance Sheets/TW Cochran 2020 Tenant Balance Sheets.xlsx',
    r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Twinwood Tenant Balance Sheets', "366 S. Cochran", "Twinwood"],

    [r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Twinwood Tenant Balance Sheets/TW Irene 2020 Tenant Balance Sheets.xlsx',
    r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Twinwood Tenant Balance Sheets', "10416 Irene", "Twinwood"],

    [r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Twinwood Tenant Balance Sheets/TW Mayfield 2020 Tenant Balance Sheets.xlsx',
    r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Twinwood Tenant Balance Sheets', "11628 Mayfield", "Twinwood"],

    [r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Twinwood Tenant Balance Sheets/TW Pelham 2020 Tenant Balance Sheets.xlsx',
    r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Twinwood Tenant Balance Sheets', "1817 Pelham", "Twinwood"],

    [r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Twinwood Tenant Balance Sheets/TW Reeves 2020 Tenant Balance Sheets.xlsx',
    r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Twinwood Tenant Balance Sheets', "220-222 S. Reeves", "Twinwood"],

    [r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Twinwood Tenant Balance Sheets/TW So Palm 2020 Tenant Balance Sheets.xlsx',
    r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Twinwood Tenant Balance Sheets', "137 So. Palm", "Twinwood"],

    [r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Brighton Trading Tenants Individualized Balance Sheet Dr and Cr..xlsx',
    r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent', "Sherbourne / Cavendish", "Brighton Trading"],

    [r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Palmaher Tenants Individualized Balance Sheets Dr. and Cr..xlsx',
    r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent', "3263 Motor", "Palmaher"]
]


def most_recent_search(current_sheet):
    ''' A function that searches for the most recent entry in column 'E', (column 5).'''

    # Search starting at the max row. Increment by -1 rows. There is no row 0, so stop at row 1.
    for i in range(current_sheet.max_row, 1, -1):
        cell = current_sheet.cell(row=i, column=5)
        #  Return the first non-empty cell found.
        if cell.value != None:
            return cell.coordinate


while True:
    '''Initial user input loop. Asks whether the user would like to check the most recent balance, or to check by month.'''
    print ("~" * 30)
    recent = input("Check most recent balance y/n?")
    print ("~" * 30)
    if recent.lower() == 'y' or recent.lower() == 'yes':
        recent_balance_check = True
        break
    elif recent.lower() == 'n' or recent.lower() == 'no':
        recent_balance_check = False
        break
    else:
        print("Invalid entry")


if not recent_balance_check:
    print("Search by month not yet implemented.")


''' Load each workbook. For each workbook, print out desired output.'''
for wbIndex in range(len(wb_list)):
    #Load each workbook one by one, and change the working directory as well.
    wb = openpyxl.load_workbook(wb_list[wbIndex][0], data_only=True)
    os.chdir(wb_list[wbIndex][1])

    while not recent_balance_check:
        month = input("What month is being checked for the balance? Enter the first three letters of any month.")
        month = month[0:3].lower()
        if month not in Month.month_list:
            print("Invalid month")
            continue
        else:
            month_checked = Month(month)
        break

    if recent_balance_check:
        for sheet in wb.sheetnames:
            # ignore the security deposit sheets
            ignore_list = ["Brighton Trading Tenants", "Chart1", "Palmaher Tenants"]
            if sheet in ignore_list:
                continue
            # wb[sheet] is the active sheet. most_recent_search(wb[sheet]) returns a cell coordinate.
            # Print the name of the tenant, which correspondes to the current sheetname.
            print("\n")
            print("Tenant name = ", sheet)
            print("Most recent payment = $", wb[sheet][most_recent_search(wb[sheet])].value)
            print("Balance after most recent payment = $", wb[sheet].cell(row = wb[sheet][most_recent_search(wb[sheet])].row, column = wb[sheet][most_recent_search(wb[sheet])].column + 1).value)
            if (wb[sheet].cell(row = wb[sheet][most_recent_search(wb[sheet])].row, column = wb[sheet][most_recent_search(wb[sheet])].column + 1)).value != None and (wb[sheet].cell(row = wb[sheet][most_recent_search(wb[sheet])].row, column = wb[sheet][most_recent_search(wb[sheet])].column + 1).value < 0):
                if wb_list[wbIndex][0] != r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Brighton Trading Tenants Individualized Balance Sheet Dr and Cr..xlsx' and wb_list[wbIndex][0] != r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Palmaher Tenants Individualized Balance Sheets Dr. and Cr..xlsx':
                    print(colored("BALANCE OWED", 'red'))
            if (wb[sheet].cell(row = wb[sheet][most_recent_search(wb[sheet])].row, column = wb[sheet][most_recent_search(wb[sheet])].column + 1)).value != None and (wb[sheet].cell(row = wb[sheet][most_recent_search(wb[sheet])].row, column = wb[sheet][most_recent_search(wb[sheet])].column + 1).value > 0):
                if wb_list[wbIndex][0] == r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Brighton Trading Tenants Individualized Balance Sheet Dr and Cr..xlsx' or wb_list[wbIndex][0] == r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Palmaher Tenants Individualized Balance Sheets Dr. and Cr..xlsx':
                    print(colored("BALANCE OWED", 'red'))