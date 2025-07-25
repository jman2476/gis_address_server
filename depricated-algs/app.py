import shapefile
from openpyxl import Workbook

sf = shapefile.Reader('../Pennypack-st-proj/mailing_list_pennypack.shp')

print(sf)

fieldnames = [f[0] for f in sf.fields[1:]]

wb = Workbook() 
ws = wb.active

# ws.append([])
ws.append(fieldnames)

records = sf.records()

for rec in records:
    ws.append(rec)

# remove excess information
# ws.delete_cols(3,5) 
ws.delete_cols(75, 4)
ws.delete_cols(46, 28)
# ws.delete_cols(36, 8)
# ws.delete_cols(31, 2)
ws.delete_cols(30, 14)
ws.delete_cols(10, 19)
ws.delete_cols(1, 8)

# reorder the columns
ws.move_range('B1:B4232', cols=4)
ws.move_range('C1:D4232', cols=-1)
ws.move_range('F1:F4232', cols=-2)

wb.save('address_list_pennypack.xlsx')
# fields needed
#  - location
#  - zip_code
#  - mailing_ci
#  - mailing_st
#  - mailing_zi
#  - owner_1
#  - owner_2