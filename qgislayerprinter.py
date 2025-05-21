# layer printer

layer = iface.activeLayer()
size = layer.featureCount()
addresses = []

for i in range(0, size):
    #print(layer.getFeature(i))
    addresses.append(layer.getFeature(i))
    
fields = addresses[0].fields()

for field in fields:
    print(field)
    
print(addresses[0].attributeMap())
