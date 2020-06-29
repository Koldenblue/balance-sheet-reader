class Workbook:
    '''Excel balance sheet to be loaded. Not currently in use.'''
    def __init__(self, file, working_directory, building, company, tenant):
        '''Each excel file has a file location, a working directory, a building, a company which the building belongs to, and a sheet for each tenant.'''
        self.file = file
        self.working_directory = working_directory
        self.building = building
        self.company = company
        self.tenant = tenant

    def get_workbook():
        while True:
            wb = input("What is the full filepath of the workbook?")
            wb_list = []
            wb_list.append(wb)


# Search according to month_regex. Is better to search using datetime format. 
def last_search(current_sheet, current_row, month):
    '''Given current sheet and the month name, recursively search for the last occurence of that month under the date column.
        Starting at current_row = 1, and then recursively adding to current_row.'''
    # Search each cell in column A for the month match, starting at current_row.
    for i in range(current_row, current_sheet.max_row):
        cell = current_sheet.cell(row=i, column = DATE_COLUMN)
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

    class Month:
    '''Defines a month such that it can be found with a regex,
        using excel's datetime.datetime format.'''
    month_list = ['narf', 'jan', 'feb', 'mar', 'apr', 'may', 
    'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec']

    def __init__ (self, month):
        '''Assigns a number, 1-12, to a string consisting of the first three letters of a month.'''
        self.month = month
        # month_index is 01 thru 12
        self.month_index = "{0:02d}".format(self.month_list.index(month[0:3].lower()))
        self.month_regex = re.compile(f'{YEAR}-{self.month_index}-')

    def month_lesser(self, second_entry):
        '''Returns True if the second month in a sequence comes before the first month.'''
        # Could also overload < and > operators using __gt__ and __lt__ (see docs)
        if self.month_index > second_entry.month_index:
            return True
        return False


        # Test Example:
        # month1 = Month('April')
        # month2 = Month('mar')
        # print(month1.month_lesser(month2))


''' 
Under search_by_recent_credit():
            # If value of cell is None, balance owed is $0. Else print out balance = cell value.
            if wb[sheet].cell(row = credit_cell.row, column = credit_cell.column + 1).value == None:
                print("Balance after most recent payment = $0")
            else:
               print("Balance after most recent payment = $", wb[sheet].cell(row = credit_cell.row, column = BALANCE_COLUMN).value)

            # Print "balance owed" if the balance is negative, and the cell is not empty. 
            if (wb[sheet].cell(row = credit_cell.row, column = credit_cell.column + 1)).value != None:
                if (wb[sheet].cell(row = credit_cell.row, column = credit_cell.column + 1).value < 0):
                    print(colored("BALANCE OWED", 'red'))
'''