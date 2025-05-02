import shapefile
from openpyxl import Workbook

sf = shapefile.Reader('../May-5-Phila-proj/properties_selection.shp')

print(sf)

fieldnames = [f[0] for f in sf.fields[1:]]

for field in fieldnames:
    print(field)

wb = Workbook()

ws = wb.active

ws.append(fieldnames)

# for row in fields:
#     ws.append(row)

# wb.save('fields.xlsx')

records = sf.records()

for rec in records:
    ws.append(rec)

wb.save('records.xlsx')