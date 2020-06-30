# balance-sheet-reader
Reads company balance sheets and gets the most recent balance for each tenant.

Balance sheets should follow the format in the given Sample Balance Sheet.xlsx. Column A is the date, B is Description, 
C is check number, D is debit, E is credit, and F is balance. If columns A, B, E, or F are changed, it will break the program!

In the "Updateable_values" file there are several values that can be edited. This includes file locations, the current year, ignored excel worksheets, and column numbers.

To update, add, or subtract excel files, wb_list will have to be updated. To update them, the format given in "sample" can be followed.

The year is currently hardcoded, and should be updated to reflect the current year,
as it will affect the monthly search range.


To use this program:
  1) Make sure that the file locations are correct. Make sure the year is correct. These are in the updateable_values file.
  2) All excel files should follow the format given in the Sample Balance sheet.
  3) Run the program in Python 3. The program should output an excel file.
  
 
 
 Possible features: Updateable values for workbook locations and year within the program. Better excel tables. Color coding. 
 Better search (but this one also depends on the quality of the balance sheets). 
