import xml.etree.ElementTree as ET
tree = ET.ElementTree(file='FR.xml')

indent = 0
ignoreElems = ['/word/document.xml']

def printRecur(root, f):
    """Recursively prints the tree."""
    # print('root tag:', root.tag, root.text)
    global indent
    if root.tag in ignoreElems:
        return
    f.write(' ' * indent + '%s: %s' % (root.tag.title(), root.attrib.get('name', root.text))+"\r")
    # print (' ' * indent + '%s: %s' % (root.tag.title(), root.attrib.get('name', root.text)))
    indent += 4
    for elem in root.getchildren():
        printRecur(elem, f)
    indent -= 4

path_w = 'FR.txt'
f = open(path_w, mode='w', encoding='utf-8')


root = tree.getroot()
printRecur(root, f)