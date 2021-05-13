from xml.etree import ElementTree
import xml.etree.ElementTree as ET
import os
xml_original_filename = "a.xml"
def register_all_namespaces(filename):
    namespaces = dict([node for _, node in ET.iterparse(filename, events=['start-ns'])])
    for ns in namespaces:
        if ( ns == "w14" or ns == "w15" or ns == "w16se" or ns == "w16cid" or
            ns == "w16" or ns == "w16cex" or ns == "wp14"):
            print("registering (skipped?):", ns)
        else:
            print("registering:", ns)
        try:
            ET.register_namespace(ns, namespaces[ns])
        except ValueError as e:
            print(e)

if __name__ == '__main__':
    with open(xml_original_filename, "r",encoding="utf-8") as f:
        xml = f.read()
    root = ElementTree.fromstring(xml)
    register_all_namespaces(xml_original_filename)
    #最上位階層のタグと中身
    print(root.tag,root.attrib)
    # ET.register_namespace('Relationships',"http://schemas.openxmlformats.org/package/2006/relationships")

    #子階層のタグと中身
    list_content_types = []
    for child in root:

        print(child.tag, child.attrib)
        filepath = "./out"+format(child.attrib['{http://schemas.microsoft.com/office/2006/xmlPackage}name'] )
        xml_name = child.attrib['{http://schemas.microsoft.com/office/2006/xmlPackage}name']
        xml_contentType = child.attrib['{http://schemas.microsoft.com/office/2006/xmlPackage}contentType']
        list_content_types.append([xml_name,xml_contentType])
        dirname, basename = os.path.split(filepath)
        try:
            os.makedirs(dirname)
        except FileExistsError as e:
            a = e
            # print(e.strerror)  # エラーメッセージ ('Cannot create a file when that file already exists')
            # print(e.errno)  # エラー番号 (17)
            # print(e.filename)  # 作成できなかったディレクトリ名 ('foo')
        print(xml_name + " CT:" + xml_contentType )

        with open(filepath, 'wb') as f:
            if xml_contentType != 'application/xml':
                f.write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n".encode('utf-8'))
            else:
                f.write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"no\"?>\n".encode('utf-8'))
            for child1 in child:
                for child2 in child1:
                    a = ET.tostring(child2)

            f.write(ET.tostring(child2, short_empty_elements=False))

    with open("./out/[Content_Types].xml", 'w') as f:
        f.write("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n")
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
    # < Override PartName = "/word/document.xml" ContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml" / >
    #
    # for name in root.iter('w:t'):
    #     print(name.text)
# {http://schemas.microsoft.com/office/2006/xmlPackage}package {}
# {http://schemas.microsoft.com/office/2006/xmlPackage}part {'{http://schemas.microsoft.com/office/2006/xmlPackage}name': '/_rels/.rels', '{http://schemas.microsoft.com/office/2006/xmlPackage}contentType': 'application/vnd.openxmlformats-package.relationships+xml', '{http://schemas.microsoft.com/office/2006/xmlPackage}padding': '512'}
# {http://schemas.microsoft.com/office/2006/xmlPackage}part {'{http://schemas.microsoft.com/office/2006/xmlPackage}name': '/word/document.xml', '{http://schemas.microsoft.com/office/2006/xmlPackage}contentType': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'}
# {http://schemas.microsoft.com/office/2006/xmlPackage}part {'{http://schemas.microsoft.com/office/2006/xmlPackage}name': '/word/_rels/document.xml.rels', '{http://schemas.microsoft.com/office/2006/xmlPackage}contentType': 'application/vnd.openxmlformats-package.relationships+xml', '{http://schemas.microsoft.com/office/2006/xmlPackage}padding': '256'}
# {http://schemas.microsoft.com/office/2006/xmlPackage}part {'{http://schemas.microsoft.com/office/2006/xmlPackage}name': '/word/footnotes.xml', '{http://schemas.microsoft.com/office/2006/xmlPackage}contentType': 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml'}
# {http://schemas.microsoft.com/office/2006/xmlPackage}part {'{http://schemas.microsoft.com/office/2006/xmlPackage}name': '/word/endnotes.xml', '{http://schemas.microsoft.com/office/2006/xmlPackage}contentType': 'application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml'}
# {http://schemas.microsoft.com/office/2006/xmlPackage}part {'{http://schemas.microsoft.com/office/2006/xmlPackage}name': '/word/header2.xml', '{http://schemas.microsoft.com/office/2006/xmlPackage}contentType': 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'}
# {http://schemas.microsoft.com/office/2006/xmlPackage}part {'{http://schemas.microsoft.com/office/2006/xmlPackage}name': '/word/footer2.xml', '{http://schemas.microsoft.com/office/2006/xmlPackage}contentType': 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml'}
# {http://schemas.microsoft.com/office/2006/xmlPackage}part {'{http://schemas.microsoft.com/office/2006/xmlPackage}name': '/word/header1.xml', '{http://schemas.microsoft.com/office/2006/xmlPackage}contentType': 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'}
# {http://schemas.microsoft.com/office/2006/xmlPackage}part {'{http://schemas.microsoft.com/office/2006/xmlPackage}name': '/word/footer1.xml', '{http://schemas.microsoft.com/office/2006/xmlPackage}contentType': 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml'}
# {http://schemas.microsoft.com/office/2006/xmlPackage}part {'{http://schemas.microsoft.com/office/2006/xmlPackage}name': '/word/_rels/header1.xml.rels', '{http://schemas.microsoft.com/office/2006/xmlPackage}contentType': 'application/vnd.openxmlformats-package.relationships+xml'}
# {http://schemas.microsoft.com/office/2006/xmlPackage}part {'{http://schemas.microsoft.com/office/2006/xmlPackage}name': '/word/theme/theme1.xml', '{http://schemas.microsoft.com/office/2006/xmlPackage}contentType': 'application/vnd.openxmlformats-officedocument.theme+xml'}
# {http://schemas.microsoft.com/office/2006/xmlPackage}part {'{http://schemas.microsoft.com/office/2006/xmlPackage}name': '/word/media/image1.png', '{http://schemas.microsoft.com/office/2006/xmlPackage}contentType': 'image/png', '{http://schemas.microsoft.com/office/2006/xmlPackage}compression': 'store'}
# {http://schemas.microsoft.com/office/2006/xmlPackage}part {'{http://schemas.microsoft.com/office/2006/xmlPackage}name': '/word/settings.xml', '{http://schemas.microsoft.com/office/2006/xmlPackage}contentType': 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml'}
# {http://schemas.microsoft.com/office/2006/xmlPackage}part {'{http://schemas.microsoft.com/office/2006/xmlPackage}name': '/customXml/item1.xml', '{http://schemas.microsoft.com/office/2006/xmlPackage}contentType': 'application/xml', '{http://schemas.microsoft.com/office/2006/xmlPackage}padding': '32'}
# {http://schemas.microsoft.com/office/2006/xmlPackage}part {'{http://schemas.microsoft.com/office/2006/xmlPackage}name': '/customXml/itemProps1.xml', '{http://schemas.microsoft.com/office/2006/xmlPackage}contentType': 'application/vnd.openxmlformats-officedocument.customXmlProperties+xml', '{http://schemas.microsoft.com/office/2006/xmlPackage}padding': '32'}
# {http://schemas.microsoft.com/office/2006/xmlPackage}part {'{http://schemas.microsoft.com/office/2006/xmlPackage}name': '/customXml/item2.xml', '{http://schemas.microsoft.com/office/2006/xmlPackage}contentType': 'application/xml', '{http://schemas.microsoft.com/office/2006/xmlPackage}padding': '32'}
# {http://schemas.microsoft.com/office/2006/xmlPackage}part {'{http://schemas.microsoft.com/office/2006/xmlPackage}name': '/customXml/itemProps2.xml', '{http://schemas.microsoft.com/office/2006/xmlPackage}contentType': 'application/vnd.openxmlformats-officedocument.customXmlProperties+xml', '{http://schemas.microsoft.com/office/2006/xmlPackage}padding': '32'}
# {http://schemas.microsoft.com/office/2006/xmlPackage}part {'{http://schemas.microsoft.com/office/2006/xmlPackage}name': '/customXml/item3.xml', '{http://schemas.microsoft.com/office/2006/xmlPackage}contentType': 'application/xml', '{http://schemas.microsoft.com/office/2006/xmlPackage}padding': '32'}
# {http://schemas.microsoft.com/office/2006/xmlPackage}part {'{http://schemas.microsoft.com/office/2006/xmlPackage}name': '/customXml/itemProps3.xml', '{http://schemas.microsoft.com/office/2006/xmlPackage}contentType': 'application/vnd.openxmlformats-officedocument.customXmlProperties+xml', '{http://schemas.microsoft.com/office/2006/xmlPackage}padding': '32'}
# {http://schemas.microsoft.com/office/2006/xmlPackage}part {'{http://schemas.microsoft.com/office/2006/xmlPackage}name': '/customXml/item4.xml', '{http://schemas.microsoft.com/office/2006/xmlPackage}contentType': 'application/xml', '{http://schemas.microsoft.com/office/2006/xmlPackage}padding': '32'}
# {http://schemas.microsoft.com/office/2006/xmlPackage}part {'{http://schemas.microsoft.com/office/2006/xmlPackage}name': '/customXml/itemProps4.xml', '{http://schemas.microsoft.com/office/2006/xmlPackage}contentType': 'application/vnd.openxmlformats-officedocument.customXmlProperties+xml', '{http://schemas.microsoft.com/office/2006/xmlPackage}padding': '32'}
# {http://schemas.microsoft.com/office/2006/xmlPackage}part {'{http://schemas.microsoft.com/office/2006/xmlPackage}name': '/word/numbering.xml', '{http://schemas.microsoft.com/office/2006/xmlPackage}contentType': 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml'}
# {http://schemas.microsoft.com/office/2006/xmlPackage}part {'{http://schemas.microsoft.com/office/2006/xmlPackage}name': '/word/styles.xml', '{http://schemas.microsoft.com/office/2006/xmlPackage}contentType': 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml'}
# {http://schemas.microsoft.com/office/2006/xmlPackage}part {'{http://schemas.microsoft.com/office/2006/xmlPackage}name': '/word/webSettings.xml', '{http://schemas.microsoft.com/office/2006/xmlPackage}contentType': 'application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml'}
# {http://schemas.microsoft.com/office/2006/xmlPackage}part {'{http://schemas.microsoft.com/office/2006/xmlPackage}name': '/word/fontTable.xml', '{http://schemas.microsoft.com/office/2006/xmlPackage}contentType': 'application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml'}
# {http://schemas.microsoft.com/office/2006/xmlPackage}part {'{http://schemas.microsoft.com/office/2006/xmlPackage}name': '/docProps/core.xml', '{http://schemas.microsoft.com/office/2006/xmlPackage}contentType': 'application/vnd.openxmlformats-package.core-properties+xml', '{http://schemas.microsoft.com/office/2006/xmlPackage}padding': '256'}
# {http://schemas.microsoft.com/office/2006/xmlPackage}part {'{http://schemas.microsoft.com/office/2006/xmlPackage}name': '/docProps/app.xml', '{http://schemas.microsoft.com/office/2006/xmlPackage}contentType': 'application/vnd.openxmlformats-officedocument.extended-properties+xml', '{http://schemas.microsoft.com/office/2006/xmlPackage}padding': '256'}
# {http://schemas.microsoft.com/office/2006/xmlPackage}part {'{http://schemas.microsoft.com/office/2006/xmlPackage}name': '/docProps/custom.xml', '{http://schemas.microsoft.com/office/2006/xmlPackage}contentType': 'application/vnd.openxmlformats-officedocument.custom-properties+xml', '{http://schemas.microsoft.com/office/2006/xmlPackage}padding': '256'}
# {http://schemas.microsoft.com/office/2006/xmlPackage}part {'{http://schemas.microsoft.com/office/2006/xmlPackage}name': '/customXml/_rels/item1.xml.rels', '{http://schemas.microsoft.com/office/2006/xmlPackage}contentType': 'application/vnd.openxmlformats-package.relationships+xml', '{http://schemas.microsoft.com/office/2006/xmlPackage}padding': '256'}
# {http://schemas.microsoft.com/office/2006/xmlPackage}part {'{http://schemas.microsoft.com/office/2006/xmlPackage}name': '/customXml/_rels/item2.xml.rels', '{http://schemas.microsoft.com/office/2006/xmlPackage}contentType': 'application/vnd.openxmlformats-package.relationships+xml', '{http://schemas.microsoft.com/office/2006/xmlPackage}padding': '256'}
# {http://schemas.microsoft.com/office/2006/xmlPackage}part {'{http://schemas.microsoft.com/office/2006/xmlPackage}name': '/customXml/_rels/item3.xml.rels', '{http://schemas.microsoft.com/office/2006/xmlPackage}contentType': 'application/vnd.openxmlformats-package.relationships+xml', '{http://schemas.microsoft.com/office/2006/xmlPackage}padding': '256'}
# {http://schemas.microsoft.com/office/2006/xmlPackage}part {'{http://schemas.microsoft.com/office/2006/xmlPackage}name': '/customXml/_rels/item4.xml.rels', '{http://schemas.microsoft.com/office/2006/xmlPackage}contentType': 'application/vnd.openxmlformats-package.relationships+xml', '{http://schemas.microsoft.com/office/2006/xmlPackage}padding': '256'}
