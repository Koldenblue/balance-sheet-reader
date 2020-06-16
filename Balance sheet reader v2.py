import openpyxl
import os
import re
from search_by_recent_credit import search_by_recent_credit
from Workbook_class import wb_list
from Month_class import Month
from colorama import init, Fore, Back, Style
from termcolor import colored
init()

'''INSTRUCTIONS'''
'''See readme. In word, '<br>' can be replaced with the find and replace command, with "^m". This will cause page breaks to appear.
Word can also be used to color the "BALANCE OWED" entries with find and replace.

Make sure the files being loaded are the most recent files!
This program will break if the original formatting of the balance sheets is changed (ie column 5 no longer is the credit column, 
or column 6 no longer is the balance column). This program also assumes that Brighton and Palmaher list positive balances as balances owed.
The other companies are assumed to list negative numbers as balances owed.'''

'''The month regex under Month_class also assumes the year is 2020.'''

'''TODO'''
'''Update readme. Update issues. Implement file writing and saving. Possibly implement an easy way to select new balance sheets and add them to the array.
Possibly implement balance search by month. Possibly implement writing to multiple file types (excel, text format, html).'''


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


if recent_balance_check:
    search_by_recent_credit()


while not recent_balance_check:
    month = input("What month is being checked for the balance? Enter the first three letters of any month.")
    if month not in Month.month_list:
        print("Invalid month")
        continue
    # Create a month object with the given month string.
    else:
        month_checked = Month(month)
        break

def last_search(current_sheet, current_row, month):
    '''Given current sheet and the month name, recursively search for the last occurence of that month under the date column.
        Starting at current_row = 1, and then recursively adding to current_row.'''
    # Search each cell in column A for the month match, starting at current_row.
    for i in range(current_row, current_sheet.max_row):
        cell = current_sheet.cell(row=i, column = 1)
        mo = month.month_regex.search(str(cell.value))
        # If the cell is not empty and the month matches, get the values and coordinates of the cell.
        if cell.value != None and mo != None:
            coord = cell.row
            val = cell.value
            # If not yet at the max row, call the search function again, starting at new row.
            if cell.row < current_sheet.max_row:
                coord2, val2 = last_search(current_sheet, cell.row + 1, month_checked)
            # if, after the search function is recursively called, new values are found, return them (base case, no new values, so these == None).
            if coord2 != None and val2 != None:
                return coord2, val2
            # otherwise return the originally found values.
            else:
                return coord, val
    # Return none if nothing found (will be base case, where coord2 = None and val2 = None).
    return None, None

if not recent_balance_check:
    for wbIndex in range(len(wb_list)):
        #Load each workbook one by one, and change the working directory as well.
        wb = openpyxl.load_workbook(wb_list[wbIndex][0], data_only=True)
        os.chdir(wb_list[wbIndex][1])
        print ("\n")
        for sheet in wb.sheetnames:
            current_sheet = wb[sheet]
            coord, val = last_search(current_sheet, 1, month_checked)
            print("active sheet = ", current_sheet)
            if coord != None:
                print("cell = A" + str(coord))
                print("Date = ", val)
                balance = current_sheet.cell(row=coord, column=6).value
                print(balance)
            else:
                print(f"No entry for {month}")

            print("")