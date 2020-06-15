month_list = ['narf', 'jan', 'feb', 'mar', 'apr', 'may', 
'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec']

month = 'jan'
print((month_list.index(month[0:3].lower())))

print("{0:02d}".format(month_list.index(month[0:3].lower())))