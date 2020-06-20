import datetime
from Updateable_values import YEAR

'''Note: datetime.datetime format looks like 2020-01-30 00:00:00'''

class Month:
    '''Defines a month such that it can be found with a regex,
        using excel's datetime.datetime format.'''
    month_list = ['narf', 'jan', 'feb', 'mar', 'apr', 'may', 
    'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec']
    full_month_names = ['narf', 'January', 'February', 'March', 'April', 'May', 
    'June', 'July', 'August', 'September', 'October', 'November', 'December']
    
    def __init__ (self, month):
        '''Assigns a number, 1-12, to a string consisting of the first three letters of a month.
        Gives the full month name corresponding to the 3 letter abbreviation.
        Converts the month to datetime.datetime format.'''
        self.month = month
        # month_index is 01 thru 12
        self.month_index = self.month_list.index(month[0:3].lower())
        self.full_name = self.full_month_names[self.month_index]
        self.datetime_format = datetime.datetime.strptime('{0} {1}'.format(self.full_name, YEAR), '%B %Y')
        
    def one_month_later(self):
        '''Gets datetime format for a month after the input month.'''
        self.next_month_index = self.month_index + 1
        new_year = False
        if self.next_month_index > 12:
            self.next_month_index = 1
            new_year = True
            new_year_num = YEAR + 1
        if new_year:
            self.next_month_datetime_format = datetime.datetime.strptime('{0} {1}'.format(self.full_month_names[self.next_month_index], new_year_num), '%B %Y')
        else:
            self.next_month_datetime_format = datetime.datetime.strptime('{0} {1}'.format(self.full_month_names[self.next_month_index], YEAR), '%B %Y')
        return self.next_month_datetime_format

'''
test_month = Month('aug')
print(test_month.month)
print(test_month.month_index)
print(test_month.full_name)

print(test_month.datetime_format)
print(test_month.one_month_later())

test2 = Month('dec')
print(test2.one_month_later())
'''
