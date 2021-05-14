"""
Convert doc to docx

Usage:
  decompose_regex.py <in_file> <out_folder>
  decompose_regex.py -h | --help
  decompose_regex.py --version

  <in_file>: filename of the doc file to be converted
  <out_folder>: output folder for the split files and
    zipped file named (new.docx)

Examples:
  decompose_regex.py examples/sample.doc out
    "sample.doc" in the current folder is converted to a docx file
    in the "out" folder under the name of "new.docx". The associated
    intermediate files such as "[Content_Types].xml" are also stored
    in the "out" folder.
  decompose_regex.py examples/sample_broken.doc out_broken
    "new.docx" will be created in "out_broken" folder. However, this
    file is corrupted with an inconsistent relationship id, which is
    changed from "rId4" to "rId40" in Relationships package.

    A file, "Error in the sample file sample_broken.doc.jpg", under
     examples folder shows how the inconsistency will be indicated
     in "Open XML SDK 2.5 Productivity Tool".

Options:
  -h --help     Show this screen.
  --version     Show version.
"""
import xml.etree.ElementTree as ET
import os
import re
from xml.etree import ElementTree
import base64
import zipfile
from docopt import docopt

def register_all_namespaces(filename):
    namespaces = dict([node for _, node in ET.iterparse(filename, events=['start-ns'])])
    for ns in namespaces:
        try:
            ET.register_namespace(ns, namespaces[ns])
        except ValueError as e:
            print(e)

if __name__ == '__main__':

    arguments = docopt(__doc__, version="0.1")

    in_file = arguments["<in_file>"]
    out_folder = arguments["<out_folder>"]
    xml_original_filename = in_file
    file_to_create = 'new.docx'

    with open(xml_original_filename, "r",encoding="utf-8") as f:
        xml = f.read()
    pattern = re.compile(r'<pkg:part.*?>.*?</pkg:part>', re.MULTILINE | re.DOTALL)
    result = pattern.finditer( xml)

    list_content_types = []

    print("Step: reading packages")

    for m in result:
        print("-->",m.group()[0:50])
        result_name = re.finditer(r'pkg:name=\".*?\"',m.group())
        filepath = out_folder
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
            print("---->",content_type_detail,filename)
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

        pattern = re.compile(r'<pkg:binaryData>(.*?)</pkg:binaryData>', re.MULTILINE | re.DOTALL)
        result_xml = pattern.findall(m.group())
        for xml in result_xml:
            with open(filepath, 'wb') as f:
                f.write(base64.urlsafe_b64decode(xml.encode('UTF-8')))

    #TODO: contentType is not exhaustive. Newly identified contentType has to be added
    # to avoid inconsistency in Word file

    with open(os.path.join(out_folder,"[Content_Types].xml"), 'w') as f:
        f.write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n")
        f.write('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">')
        f.write('<Default Extension="png" ContentType="image/png"/>')
        f.write('<Default Extension="jpeg" ContentType="image/jpeg"/>')
        f.write('<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>')
        f.write('<Default Extension="xml" ContentType="application/xml"/>')
        for content_type in list_content_types:
            if ( content_type[1] != 'application/xml' and
                content_type[1] != 'application/vnd.openxmlformats-package.relationships+xml'):
                f.write('<Override PartName="' + content_type[0] +
                        '" ContentType="' + content_type[1] + '"/>')
        f.write('</Types>')

    print("Step: zipping files")
    walk_result = os.walk(out_folder)
    with zipfile.ZipFile(os.path.join(out_folder,file_to_create), 'w', compression=zipfile.ZIP_STORED) as new_zip:
        for dirname, subdirs, files in walk_result:
            for filename in files:
                if filename == file_to_create:
                    continue
                dir_without_root = dirname.replace(out_folder,"")
                print("-->",dir_without_root, filename)
                new_zip.write(os.path.join(dirname, filename),os.path.join(dir_without_root,filename))

    print("\r\nCompleted.\n\r-->File \"{filename}\" is created under the folder \"{folder}\""
          " from the input file \"{in_file}\".".format(filename=file_to_create,folder=out_folder,in_file=in_file))