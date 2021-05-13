import xml.etree.ElementTree as ET
import os
import re
from xml.etree import ElementTree

def register_all_namespaces(filename):
    namespaces = dict([node for _, node in ET.iterparse(filename, events=['start-ns'])])
    for ns in namespaces:
        # if ( ns == "w14" or ns == "w15" or ns == "w16se" or ns == "w16cid" or
        #     ns == "w16" or ns == "w16cex" or ns == "wp14"):
        #     print("registering (skipped?):", ns)
        # else:
        #     print("registering:", ns)
        try:
            ET.register_namespace(ns, namespaces[ns])
        except ValueError as e:
            print(e)
xml_original_filename = "a.xml"
xml_original_filename = "D:/file_error.DOC.xml"
if __name__ == '__main__':
    with open(xml_original_filename, "r",encoding="utf-8") as f:
        xml = f.read()
    pattern = re.compile(r'<pkg:part.*?>.*?</pkg:part>', re.MULTILINE | re.DOTALL)
    result = pattern.finditer( xml)
    #print(result)
    # <callable_iterator object at 0x10b0efa90>

    # print(type(result))
    # <class 'callable_iterator'>
    list_content_types = []
    for m in result:
        print(m.group()[0:50])
        result_name = re.finditer(r'pkg:name=\".*?\"',m.group())
        filepath = "./outxx"
        for name in result_name:
            filename = name.group()[10:-1]
        filepath = filepath + filename
        dirname, basename = os.path.split(filepath)
        try:
            os.makedirs(dirname)
        except FileExistsError as e:
            a = e

        xml_contentType = ""
        result_contentType = re.findall(r'pkg:contentType=\"(.*?)\"',m.group())
        for ct in result_contentType:
            content_type_detail = ct
            print(content_type_detail,filename)
            xml_contentType = content_type_detail = ct


        xml_name = filename

        pattern = re.compile(r'<pkg:xmlData.*?>(.*?)</pkg:xmlData>', re.MULTILINE | re.DOTALL)
        result_xml = pattern.findall(m.group())
        for xml in result_xml:
            with open(filepath, 'wb') as f:
                #"<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
                root = ElementTree.fromstring( xml)
                register_all_namespaces(xml_original_filename)

                for child in root:

                    #print(child.tag, child.attrib)
                    # filepath = "./out"+format(child.attrib['{http://schemas.microsoft.com/office/2006/xmlPackage}name'] )
                    # xml_name = child.attrib['{http://schemas.microsoft.com/office/2006/xmlPackage}name']
                    if '{http://schemas.microsoft.com/office/2006/xmlPackage}contentType' in child.attrib:
                        xml_contentType = child.attrib['{http://schemas.microsoft.com/office/2006/xmlPackage}contentType']
                if xml_contentType != 'application/xml':
                    f.write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n".encode('utf-8'))
                else:
                    f.write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"no\"?>\n".encode('utf-8'))
                list_content_types.append([xml_name,xml_contentType])
                f.write(xml.encode('utf-8'))

    with open("./outxx/[Content_Types].xml", 'w') as f:
        f.write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n")
        f.write('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">')
        f.write('<Default Extension="png" ContentType="image/png"/>')
        f.write('<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>')
        f.write('<Default Extension="xml" ContentType="application/xml"/>')
        for content_type in list_content_types:
            if ( content_type[1] != 'application/xml' and
                content_type[1] != 'application/vnd.openxmlformats-package.relationships+xml'):
                f.write('<Override PartName="' + content_type[0] +
                        '" ContentType="' + content_type[1] + '"/>')
        f.write('</Types>')
