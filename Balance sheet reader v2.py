#! python 3

import openpyxl
import os
import re
from search_by_recent_credit import search_by_recent_credit, prev_payment_search
from Updateable_values import wb_list, ignore_list, DATE_COLUMN, BALANCE_COLUMN, CREDIT_COLUMN
from Month_class import Month
from colorama import init, Fore, Back, Style
from termcolor import colored
init()
import datetime

'''INSTRUCTIONS'''
'''See readme. In word, '<br>' can be replaced with the find and replace command, with "^m". This will cause page breaks to appear.
Word can also be used to color the "BALANCE OWED" entries with find and replace.

Make sure the files being loaded are the most recent files!

The program will search for months in 2020, according to the YEAR constant. Other contsants 
designate which columns correspond to credit, date, and balance columns.
'''

'''Some sheets are manually ignored, since the balance sheet excel files have extra non-balance sheets.'''

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


def last_search(current_sheet, current_row, month):
    '''Given current sheet and the month name, recursively search for the last occurence of that month under the date column.
        Starting at current_row = 1, and then recursively adding to current_row.'''
    # Search each cell in column A with a loop, starting at current_row.
    for i in range(current_row, current_sheet.max_row):
        # First get the target month and subsequent month in datetime format.
        target_time_value = month.datetime_format
        next_month_value = month.one_month_later()
        match = False
        # Get the current cell in the date column.
        cell = current_sheet.cell(row=i, column = DATE_COLUMN)

        # Initialize cell_time_value to None if empty, or to the value in the cell.
        cell_time_value = None
        if cell.value != None and type(cell.value) == datetime.datetime:
            cell_time_value = cell.value
            # Find a match if subsequent cells are the same month or a previous month.
            # Done this way because sometimes months are entered out of order.
            # Will not match future months.
            if cell_time_value < target_time_value:
                match = True
            if cell_time_value >= target_time_value and cell_time_value < next_month_value:
                match = True

        # If the cell is not empty and the month matches, get the values and coordinates of the cell.
        if match == True: #found proper month:
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

while not recent_balance_check:
    month = input("What month is being checked for the balance? Enter the first three letters of any month.")
    try:
        month_checked = Month(month)
        break
    # Handle error if the Month constructor is given an invalid month as a value.
    except ValueError:
        print('Invalid month entry')
        continue


if not recent_balance_check:
    for wbIndex in range(len(wb_list)):
        #Load each workbook one by one, and change the working directory as well.
        wb = openpyxl.load_workbook(wb_list[wbIndex][0], data_only=True)
        os.chdir(wb_list[wbIndex][1])
        print ("\n")

        # for each sheet in the workbook, use last_search to find the last entry in the date column that 
        # corresponds to the input month, or to any previous months. Start searching at row 1.
        for sheet in wb.sheetnames:
            if sheet in ignore_list:
                continue
            current_sheet = wb[sheet]
            coord, val = last_search(current_sheet, 1, month_checked)
            print("active sheet = ", current_sheet)

            # If there is an entry, print out the cell location (A1, A2, etc.) then print out the date.
            # Print out the balance in the same row as the date.
            if coord != None:
                readable_date = datetime.datetime.strftime(val, '%B %d, %Y')
                # print("cell = A" + str(coord))
                payment = current_sheet.cell(row=coord, column=CREDIT_COLUMN).value
                balance = current_sheet.cell(row=coord, column=BALANCE_COLUMN).value

                print(f"Final balance entry, dated {readable_date}:", balance)
                
                # The balance column should never be empty, so it should never == None.
                # The payment column is only filled if a payment has been made.
                if payment != None:
                    print("Payment of ${1} received on {0}".format(readable_date, payment))
                else:
                    print(f"No payment listed for final balance entry on {readable_date}.")
                    prev_payment_coord, prev_payment_row = prev_payment_search(current_sheet, coord)
                    if prev_payment_coord != None:
                        prev_payment_date = current_sheet.cell(row=prev_payment_row, column=DATE_COLUMN).value
                        if prev_payment_date != None:
                            readable_prev_payment_date = datetime.datetime.strftime(prev_payment_date, '%B %d, %Y')
                            print(f"Previous payment entry listed as ${current_sheet[str(prev_payment_coord)].value} received on {readable_prev_payment_date}.")
                        else:
                            print("No previous payment found.")
                    else:
                        print("No previous payment found.")
            else:
                print(f"No entry for {month} or for previous months.")

            print("")

'''TODO'''
'''Update readme. Update issues list. Implement file writing and saving. Possibly implement an easy way to select new balance sheets and add them to the array.
Update search by recent credit to be clear about the date corresponding to the credit.
Possibly implement writing to multiple file types (excel, text format, html).
Print number of tenants checked.'''


