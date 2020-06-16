import re

'''Note: month regex is specific to 2020'''
'''Note: datetime.datetime format looks like 2020-01-30 00:00:00'''

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
        self.month_regex = re.compile(f'2020-{self.month_index}-')

    def month_lesser(self, second_entry):
        '''Returns True if the second month in a sequence comes before the first month.'''
        if self.month_index > second_entry.month_index:
            return True
        return False

        # Test Example:
        # month1 = Month('April')
        # month2 = Month('mar')
        # print(month1.month_lesser(month2))



