import openpyxl
import colorama
from termcolor import colored
from Workbook_class import wb_list
import os

def most_recent_search(current_sheet):
    ''' A function that searches for the most recent entry in column 'E', (column 5).'''

    # Search starting at the max row. Increment by -1 rows. There is no row 0, so stop at row 1.
    for i in range(current_sheet.max_row, 1, -1):
        cell = current_sheet.cell(row=i, column=5)
        #  Return the first non-empty cell found.
        if cell.value != None:
            return cell.coordinate


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
            ignore_list = ["Brighton Trading Tenants", "Chart1", "Palmaher Tenants"]
            if sheet in ignore_list:
                continue
            # wb[sheet] is the active sheet. most_recent_search(wb[sheet]) returns a cell coordinate.
            # Print the name of the tenant, which correspondes to the current sheetname.
            print("")
            print("Tenant name = ", sheet)
            print("Most recent payment = $", wb[sheet][most_recent_search(wb[sheet])].value)

            # If value of cell is None, balance owed is $0. Else print out balance = cell value.
            if wb[sheet].cell(row = wb[sheet][most_recent_search(wb[sheet])].row, column = wb[sheet][most_recent_search(wb[sheet])].column + 1).value == None:
                print("Balance after most recent payment = $0")
            else:
               print("Balance after most recent payment = $", wb[sheet].cell(row = wb[sheet][most_recent_search(wb[sheet])].row, column = wb[sheet][most_recent_search(wb[sheet])].column + 1).value)

            # Print "balance owed" if the balance is negative, and the cell is not empty. 
            if (wb[sheet].cell(row = wb[sheet][most_recent_search(wb[sheet])].row, column = wb[sheet][most_recent_search(wb[sheet])].column + 1)).value != None and (wb[sheet].cell(row = wb[sheet][most_recent_search(wb[sheet])].row, column = wb[sheet][most_recent_search(wb[sheet])].column + 1).value < 0):
                    print(colored("BALANCE OWED", 'red'))