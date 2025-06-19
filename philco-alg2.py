#  This script is intended to be run in the
#  QGIS python console, and as such does not
#  import the necessary modules for it to run

# choose active layer
layer = iface.activeLayer()
size = layer.featureCount()
print(size)
# parse addresses into dict
addresses = []
for i in range(0, size):
    addresses.append(layer.getFeature(i))

# select columns that are important
fields = addresses[0].fields()
namesf = fields.names

req_col = [8, 28, 43, 44, 73]
adds_reap = []
field_reap = QgsFields()

# iterate through addresses
for lot in addresses:
    row_reap = []
    for col in range(0, len(fields)):
        try:
            # print(col, req_col)
            req_col.index(col)
            # print(req_col.index(col))
            row_reap.append(lot[col])
        except:
            continue
    adds_reap.append(row_reap)

# iterate through fields
for i in range(0, len(fields)):
    try:
        req_col.index(i)
        field_reap.append(fields[i])
    except:
        continue

# write address dict to new layer
# fields = field_reap
# output driver --> csv
crs = QgsProject.instance().crs()
transform_context = QgsProject.instance().transformContext()
save_opt = QgsVectorFileWriter.SaveVectorOptions()
save_opt.driverName = 'CSV'
save_opt.fileEncoding = 'UTF-8'

writer = QgsVectorFileWriter.create(
    ##################################################
    # Change the file name below to suit the project  #
    # This script will overwrite a file with the same #  
    #   name without warning, so make it unique       #
    "placeholder-name.csv",                           #
    ##################################################
    field_reap,
    QgsWkbTypes.Point,
    crs,
    transform_context,
    save_opt
)

# transform addresses into dict
add_dict = {}
for i in range(0, len(adds_reap)):
    add_dict.update({i: adds_reap[i]})


# add features to file
for entry in add_dict:
    feat = QgsFeature()
    feat.setAttributes(add_dict[entry])
    writer.addFeature(feat)

# save layer as csv
del writer

print("success? dont forget to check")