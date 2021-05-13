import xml.etree.ElementTree as ET
context = ET.iterparse('a.xml', events=('end', ))
context = ET.iterparse('a.xml')
index = 0
for event, elem in context:
    print(elem)
    if elem.tag == '{http://schemas.microsoft.com/office/2006/xmlPackage}part':
        index += 1
        filename = format(elem.attrib.Target, str(index) + ".xml")
        with open(filename, 'wb') as f:
            f.write("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n")
            f.write(ET.tostring(elem))