import shapefile
from openpyxl import Workbook

print('done')
# Read shapefile (.shp)
sf = shapefile.Reader('../Line2678_proj/mailing-area.shp', encoding='latin1')
print(sf)
# Read fieldnames and records
fieldnames = [f[0] for f in sf.fields[1:]]

records = sf.records()

# setup buckets
fields_culled = []
records_culled = []

# select columns
required_columns = [8, 28, 29, 30, 32, 33, 34, 43, 44, 73]

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

# TODO: Split location and mailing addresses 
# set references on relative index
loc_columns = [0, 1, 7, 8, 9]
mail_columns = [0, 1, 2, 3, 4, 5, 6]

# setup fields and records
loc_fields = []
mail_fields = []
loc_records = []
mail_records = []

# iterate fields
for i in range(0,10):
    try:
        loc_columns.index(i)
        loc_fields.append(fields_culled[i])
    except: 
        continue 
for i in range(0, 8):
    try:
        mail_columns.index(i)
        mail_fields.append(fields_culled[i])
    except:
        continue

# iterate records
for rec in records_culled:
    loc_row = []
    mail_row = []
    for i in range(0, 10):
        try:
            loc_columns.index(i)
            loc_row.append(rec[i])
        except:
            continue
    for i in range(0, 8):
        try:
            mail_columns.index(i)
            mail_row.append(rec[i])
        except:
            continue
    loc_records.append(loc_row)
    mail_records.append(mail_row)

# Filter mailing list
filtered_mail_records = []
for row in mail_records:
    if row[1] == row[5]:
        print(row[1], row[5])
    else:
        filtered_mail_records.append(row)    

for filt_rec in filtered_mail_records:
    print(filt_rec)
# Create workbook
wb1 = Workbook()
wb2 = Workbook()
ws1 = wb1.active
ws2 = wb2.active

# append data to the worksheet
ws1.append(mail_fields)
for rec in filtered_mail_records:
    ws1.append(rec)

ws2.append(loc_fields)
for rec in loc_records:
    ws2.append(rec)

# # Reorder columns mailing list
ws1.move_range('F1:F1186', cols=-2)
ws1.move_range('G1:G1186', cols=-1)
# Reorder columns locations sheet
ws2.move_range('B1:B1186', cols=4)
ws2.move_range('C1:D1186', cols=-1)
ws2.move_range('F1:F1186', cols=-2)


# Save files
# wb1.save('side_woodland_list.xlsx')
wb2.save('Line-2678.xlsx')