'''THIS FILE CONTAINS ALL UPDATEABLE YEARLY VALUES, CONSTANTS, AND FILE LOCATIONS'''

# Year is the current year.
# Columns numbers are assigned according to the formatting of the balance sheets.
YEAR = 2020
DATE_COLUMN = 1
CREDIT_COLUMN = 5
BALANCE_COLUMN = 6

# ignore_list is a list of Excel chart names to be ignored by this program
# Ignored because the named charts are not balance sheets.
ignore_list = ["Brighton Trading Tenants", "Chart1", "Palmaher Tenants"]

# wb_list is a 2d Array where:
# wb_list[i][0] == filename
# wb_list[i][1] == working directory
# wb_list[i][2] is the name of the corresponding property.
# wb_list[i][3] is the name of the corresponding company.
# The file locations may have to be changed per year.
wb_list = [
    [r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Marlin Westwood Tenant Balance Sheets/1736 Westwood Tenant Balance Sheets.xlsx',
    r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Marlin Westwood Tenant Balance Sheets', "1736 Westwood", "Marlin Westwood"],

    [r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Marlin Westwood Tenant Balance Sheets/1740 Westwood Tenant Balance Sheet.xlsx',
    r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Marlin Westwood Tenant Balance Sheets', "1740 Westwood", "Marlin Westwood"],

    [r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Marlin Westwood Tenant Balance Sheets/1750 Westwood Tenant Balance Sheets.xlsx',
    r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Marlin Westwood Tenant Balance Sheets', "1750 Westwood", "Marlin Westwood"],

    [r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Marlin Westwood Tenant Balance Sheets/1760 Westwood Tenant Balance Sheets.xlsx',
    r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Marlin Westwood Tenant Balance Sheets', "1760 Westwood", "Marlin Westwood"],

    [r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Marlin Westwood Tenant Balance Sheets/MW Hilts 2020 Tenant Balance Sheets.xlsx',
    r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Marlin Westwood Tenant Balance Sheets', "1624 Hilts", "Marlin Westwood"],

    [r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Twinwood Tenant Balance Sheets/TW Cochran 2020 Tenant Balance Sheets.xlsx',
    r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Twinwood Tenant Balance Sheets', "366 S. Cochran", "Twinwood"],

    [r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Twinwood Tenant Balance Sheets/TW Irene 2020 Tenant Balance Sheets.xlsx',
    r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Twinwood Tenant Balance Sheets', "10416 Irene", "Twinwood"],

    [r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Twinwood Tenant Balance Sheets/TW Mayfield 2020 Tenant Balance Sheets.xlsx',
    r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Twinwood Tenant Balance Sheets', "11628 Mayfield", "Twinwood"],

    [r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Twinwood Tenant Balance Sheets/TW Pelham 2020 Tenant Balance Sheets.xlsx',
    r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Twinwood Tenant Balance Sheets', "1817 Pelham", "Twinwood"],

    [r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Twinwood Tenant Balance Sheets/TW Reeves 2020 Tenant Balance Sheets.xlsx',
    r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Twinwood Tenant Balance Sheets', "220-222 S. Reeves", "Twinwood"],

    [r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Twinwood Tenant Balance Sheets/TW So Palm 2020 Tenant Balance Sheets.xlsx',
    r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Twinwood Tenant Balance Sheets', "137 So. Palm", "Twinwood"],

    [r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Brighton Trading Tenants Individualized Balance Sheet Dr and Cr..xlsx',
    r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent', "Sherbourne / Cavendish", "Brighton Trading"],

    [r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent\Palmaher Tenants Individualized Balance Sheets Dr. and Cr..xlsx',
    r'\\Optiplex7440\c\Rents\Rent 2020\Tenant Rent', "3263 Motor", "Palmaher"]
]


# To use sample worksheet: switch the variable names of wb_list and sample. Update the file location of the sample sheet.
sample = [
    [r'<FILEPATH>\Sample Balance Sheet.xlsx', r'<WORKING DIRECTORY - SAMPLE FILE LOCATION>', 'SAMPLE ADDRESS',
    'SAMPLE COMPANY']
    ]