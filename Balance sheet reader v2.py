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
# The file locations may have to be changed per year.
wb_list = [
    [r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Marlin Westwood Tenant Balance Sheets/1736 Westwood Tenant Balance Sheets.xlsx',
    r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Marlin Westwood Tenant Balance Sheets'],

    [r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Marlin Westwood Tenant Balance Sheets/1740 Westwood Tenant Balance Sheet.xlsx',
    r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Marlin Westwood Tenant Balance Sheets'],

    [r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Marlin Westwood Tenant Balance Sheets/1750 Westwood Tenant Balance Sheets.xlsx',
    r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Marlin Westwood Tenant Balance Sheets'],

    [r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Marlin Westwood Tenant Balance Sheets/1760 Westwood Tenant Balance Sheets.xlsx',
    r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Marlin Westwood Tenant Balance Sheets'],

    [r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Marlin Westwood Tenant Balance Sheets/MW Hilts 2020 Tenant Balance Sheets.xlsx',
    r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Marlin Westwood Tenant Balance Sheets'],

    [r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Twinwood Tenant Balance Sheets/TW Cochran 2020 Tenant Balance Sheets.xlsx',
    r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Twinwood Tenant Balance Sheets'],

    [r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Twinwood Tenant Balance Sheets/TW Irene 2020 Tenant Balance Sheets.xlsx',
    r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Twinwood Tenant Balance Sheets'],

    [r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Twinwood Tenant Balance Sheets/TW Mayfield 2020 Tenant Balance Sheets.xlsx',
    r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Twinwood Tenant Balance Sheets'],

    [r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Twinwood Tenant Balance Sheets/TW Pelham 2020 Tenant Balance Sheets.xlsx',
    r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Twinwood Tenant Balance Sheets'],

    [r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Twinwood Tenant Balance Sheets/TW Reeves 2020 Tenant Balance Sheets.xlsx',
    r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Twinwood Tenant Balance Sheets'],

    [r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Twinwood Tenant Balance Sheets/TW So Palm 2020 Tenant Balance Sheets.xlsx',
    r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Twinwood Tenant Balance Sheets'],

    [r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Brighton Trading Tenants Individualized Balance Sheet Dr and Cr..xlsx',
    r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent'],

    [r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Palmaher Tenants Individualized Balance Sheets Dr. and Cr..xlsx',
    r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent'],

    [r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Palmaher Tenants Individualized Balance Sheets Dr. and Cr..xlsx',
    r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent']
]

def most_recent_search(current_sheet):
    # column 5 is the 'credit' column, column E
    for i in range(current_sheet.max_row, 1, -1):
        cell = current_sheet.cell(row=i, column=5)
        if cell.value != None:
            return cell.coordinate


while True:
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
            # wb[sheet] is the active sheet. most_recent_search(wb[sheet]) returns a cell coordinate.
            print("\n")
            print("active sheet = ", wb[sheet])
            print("Most recent payment = $", wb[sheet][most_recent_search(wb[sheet])].value)
            print("Balance after most recent payment = $", wb[sheet].cell(row = wb[sheet][most_recent_search(wb[sheet])].row, column = wb[sheet][most_recent_search(wb[sheet])].column + 1).value)
            if (wb[sheet].cell(row = wb[sheet][most_recent_search(wb[sheet])].row, column = wb[sheet][most_recent_search(wb[sheet])].column + 1).value < 0):
                print(colored("BALANCE OWED", 'red'))

