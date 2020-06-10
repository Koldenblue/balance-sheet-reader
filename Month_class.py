class Month:
    month_list = ['narf', 'jan', 'feb', 'mar', 'apr', 'may', 
    'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec']

    def __init__ (self, month):
        self.month = month
        # month_index is 01 thru 12
        self.month_index = "{0:02d}".format(self.month_list.index(month[0:3].lower()))

