class Workbook:
    '''Excel balance sheet to be loaded'''
    def __init__(self, file, company, working_directory):
        '''Each excel file has a file location, a company which it belongs to, and a sheet for each tenant.'''
        self.file = file
        self.company = company
        self.working_directory = working_directory