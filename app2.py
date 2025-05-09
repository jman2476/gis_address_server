import shapefile
from openpyxl import Workbook

print('done')
# Read shapefile (.shp)
sf = shapefile.Reader('../Pennypack-st-proj/mailing_list_pennypack.shp')

# Read fieldnames and records
fieldnames = [f[0] for f in sf.fields[1:]]
fields = [sf.fields[1:]] # not necessary
records = sf.records()

# setup buckets
fields_culled = []
records_culled = []

# select columns
required_columns = [28, 29, 30, 32, 33, 34, 43, 44, 73]

# iterate through fields, selecting only what is required
for i in range(0,77):
    try:
        required_columns.index(i)
        print(i, fieldnames[i])
        fields_culled.append(fieldnames[i])
    except:
        continue

# iterate through records, selecting only required columns
for rec in records:
    row_culled = []
    for i in range(0,77):
        try:
            required_columns.index(i)
            row_culled.append(rec[i])
        except:
            continue
    records_culled.append(row_culled)

         

# Create workbook
wb = Workbook()
ws = wb.active

# append data to the worksheet
ws.append(fields_culled)
for rec in records_culled:
    ws.append(rec)

wb.save('culling_attmpt1.xlsx')