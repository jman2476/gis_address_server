import shapefile
from openpyxl import Workbook

# Read shapefile (.shp)
sf = shapefile.Reader('../Upper-Derby-proj/mailing_area.shp', encoding='latin1')
print(sf)
# Read fieldnames and records
fieldnames = [f[0] for f in sf.fields[1:]]

records = sf.records()
culled_records = []

# select columns
required_columns = [1, 12, 13, 14, 15, 16]
col_names = ['PIN', 'number', 'ADRADD', 'ADRDIR', 'street', 'ADRSUF']

# select required data
for record in records:
    new_row = []
    for i in required_columns:
        new_row.append(record[i])
    culled_records.append(new_row)

for rec in culled_records:
    print(rec)

# Create workbook
wb = Workbook()
ws = wb.active

# append data to worksheet
ws.append(col_names)
for rec in culled_records:
    ws.append(rec)

wb.save('upperdarby.xlsx')