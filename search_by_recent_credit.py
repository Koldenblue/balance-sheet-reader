import openpyxl
import colorama
from termcolor import colored
from Updateable_values import wb_list, CREDIT_COLUMN, ignore_list, DATE_COLUMN, BALANCE_COLUMN
import os
import datetime

def most_recent_search(current_sheet):
    ''' A function that searches for the most recent entry in column 'E', (column 5).'''

    # Search starting at the max row. Increment by -1 rows. There is no row 0, so stop at row 1.
    for i in range(current_sheet.max_row, 1, -1):
        cell = current_sheet.cell(row=i, column=CREDIT_COLUMN)
        #  Return the first non-empty cell found.
        if cell.value != None:
            return cell.coordinate
    return None

def prev_payment_search(current_sheet, row):
    '''Searches the credit column, similar to most_recent_search(),
         but starts from the input row, increments by -1, and stops at row 1. '''
    for i in range(row, 1, -1):
        cell = current_sheet.cell(row=i, column=CREDIT_COLUMN)
        #  Return the first non-empty cell found.
        if cell.value != None:
            return cell.coordinate, cell.row
    return None, None


def search_by_recent_credit():
    ''' Load each workbook. For each workbook, print out desired output.'''
    new_workbook = False
    for wbIndex in range(len(wb_list)):
        #Load each workbook one by one, and change the working directory as well.
        wb = openpyxl.load_workbook(wb_list[wbIndex][0], data_only=True)
        os.chdir(wb_list[wbIndex][1])

        # Print out company name at the top of each new workbook.
        if new_workbook == True:
            print("<br>")
            print("Company = " + wb_list[wbIndex][3] + ", ", end="")
            print("Property = " + wb_list[wbIndex][2])
            print("~" * 80)

        if new_workbook == False:
            print("Company = " + wb_list[wbIndex][3] + ", ", end="")
            print("Property = " + wb_list[wbIndex][2])
            print("~" * 80)
            new_workbook = True

        for sheet in wb.sheetnames:
            # ignore the security deposit sheets
            if sheet in ignore_list:
                continue
            if most_recent_search(wb[sheet]) == None:
                print("No credit entries listed on", sheet, ".")
                continue
            most_recent_search_cell = wb[sheet][most_recent_search(wb[sheet])]
            # wb[sheet] is the active sheet. most_recent_search(wb[sheet]) returns a cell coordinate.
            # Print the name of the tenant, which correspondes to the current sheetname.
            print("")
            print("Tenant name = ", sheet)
            try:
                most_recent_credit_date = datetime.datetime.strftime(wb[sheet].cell(row = most_recent_search_cell.row, column = DATE_COLUMN).value, '%B %d, %Y')
                print("Most recent payment = $", most_recent_search_cell.value, "listed on", most_recent_credit_date, ".")
            except TypeError:
                print("Most recent payment = $", most_recent_search_cell).value, "No date Listed."
            # If value of cell is None, balance owed is $0. Else print out balance = cell value.
            if wb[sheet].cell(row = most_recent_search_cell.row, column = most_recent_search_cell.column + 1).value == None:
                print("Balance after most recent payment = $0")
            else:
               print("Balance after most recent payment = $", wb[sheet].cell(row = most_recent_search_cell.row, column = BALANCE_COLUMN).value)

            # Print "balance owed" if the balance is negative, and the cell is not empty. 
            if (wb[sheet].cell(row = most_recent_search_cell.row, column = most_recent_search_cell.column + 1)).value != None and (wb[sheet].cell(row = most_recent_search_cell.row, column = most_recent_search_cell.column + 1).value < 0):
                    print(colored("BALANCE OWED", 'red'))