"""
Convert doc to docx

Usage:
  decompose_regex.py <in_file> <out_folder> <out_zipfolder> <debug_level> <remove_sytles>
  decompose_regex.py -h | --help
  decompose_regex.py --version

  <in_file>: filename of the doc file to be converted
  <out_folder>: output folder for the split files of docx file
  <out_zipfolder>: output folder for zipped file named (new.docx) and
    segment docx and XML files.
  <debug_level>: INFO, DEBUG to configure the levels of logging messages
  <remove_sytles>: remove *Style tags in the file (1..remove/0..retain)

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
#20210614 current_tag_opening must be copied from actual paths
#TODO: use specifications within tag to expand. Need to create both specified / non-specified versions

import xml.etree.ElementTree as ET
from lxml import etree
import os
import re
from xml.etree import ElementTree
import base64
import zipfile
from docopt import docopt
import logging
from tqdm import tqdm
import pandas as pd
import win32com.client
import pywintypes

#Tree visualization
import networkx as nx
from networkx.drawing.nx_agraph import graphviz_layout
from matplotlib import pyplot as plt

def register_all_namespaces(filename):
    namespaces = dict([node for _, node in ET.iterparse(filename, events=['start-ns'])])
    for ns in namespaces:
        try:
            ET.register_namespace(ns, namespaces[ns])
        except ValueError as e:
            logger.debug(e)

def create_docx(loop_index = None):

    with open(xml_original_filename, "r",encoding="utf-8") as f:
        xml = f.read()
    pattern = re.compile(r'<pkg:part.*?>.*?</pkg:part>', re.MULTILINE | re.DOTALL)
    result = pattern.finditer( xml)

    list_content_types = []

    logging.debug("Step: reading packages")

    for m in result:
        logging.debug(str(("-->",m.group()[0:50])))
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
            logging.debug(str(("---->",content_type_detail,filename)))
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

                # TODO: switch for the following functionalities
                # 1) remove nbsp
                #   xml = xml.replace(u'\xa0', ' ')
                # 2) remove styles
                global remove_styles
                if remove_styles == "1":
                    if filename == "/word/document.xml":
                        pattern_style = re.compile(r'Style w:val="(.*?)"', re.MULTILINE | re.DOTALL)
                        xml = pattern_style.sub('Style w:val="Normal"',xml)

                f.write(xml.encode('utf-8'))

                if filename == "/word/document.xml":
                    document_string = xml

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

    return document_string

def zip_and_test(file_to_create = "", tagname = ""):
    logging.debug("Step: zipping files")
    walk_result = os.walk(out_folder)
    with zipfile.ZipFile(os.path.join(out_zip_folder,file_to_create), 'w', compression=zipfile.ZIP_STORED) as new_zip:
        for dirname, subdirs, files in walk_result:
            for filename in files:
                if filename == file_to_create:
                    continue
                dir_without_root = dirname.replace(out_folder,"")
                logging.debug(str(("-->",dir_without_root, filename)))
                new_zip.write(os.path.join(dirname, filename),os.path.join(dir_without_root,filename))

    logging.debug(str(("\r\nCompleted.\n\r-->File \"{filename}\" is created under the folder \"{folder}\""
          " from the input file \"{in_file}\".".format(filename=file_to_create,folder=out_zip_folder,in_file=in_file))))

    try:
        logging.debug(os.path.join(os.getcwd(),out_zip_folder,file_to_create)  )
        doc = wd_app.Documents.Open( os.path.join(os.getcwd(),out_zip_folder,file_to_create)   )
        doc.Close()
        return True
    except AttributeError as e:
        logging.debug("Attibute Error:",e)
        return False
    except pywintypes.com_error as e:
        logging.debug("com_error -  {}".format(str(e.excepinfo)))
        broken_file_name = os.path.join(out_zip_folder, file_to_create + "broken_"+tagname+".docx")
        if os.path.exists(broken_file_name) == True:
            os.remove(broken_file_name)
            os.rename(os.path.join(out_zip_folder, file_to_create), broken_file_name)
        else:
            os.rename(os.path.join(out_zip_folder,file_to_create),broken_file_name)
        return False

def complement_empty_element(document_master = ""):
    # self-closing / empty element to closing element
    # TODO: automate tag discovery that needs closing element

    for replacing_tag in ['w:p','w:ind','w:bookmarkEnd']:
        p_tag = re.compile('(<' + replacing_tag + ' .*?>)')
        for string in p_tag.findall(document_master ):
            if string[-2:] == "/>":
                new_string = string.replace("/>","></" + replacing_tag+">")
                logging.debug(str(string) +"->" + str(new_string))
                document_master = document_master.replace(string,new_string)

    for replacing_tag in ['w:cantSplit','w:sectPr']:
        p_tag = re.compile('(<' + replacing_tag + '.*?>)')
        for string in p_tag.findall(document_master ):
            if string[-2:] == "/>":
                new_string = string.replace("/>","></" + replacing_tag+">")
                logging.debug(str(string) +"->" + str(new_string))
                document_master = document_master.replace(string,new_string)

    return document_master

def analyze_create_splitter(document_master=""):

    parser = etree.XMLParser(load_dtd=True,
                         no_network=False)
    try:
        root = etree.fromstring(document_master, parser=parser)
    except etree.DocumentInvalid as e:
        for error in e.error_log:
            logging.debug("xml_doc parser -  Line {}: {}".format(error.line, error.message))

    namespaces = dict([node for _, node in root.iter(['start-ns'])])
    for ns in namespaces:
        try:
            ET.register_namespace(ns, namespaces[ns])
        except ValueError as e:
            logger.debug(e)

    global node_index

    top_node = root
    parsed_path:str = ""
    node_index = 0
    printRecur(top_node, root,parsed_path,node_index,0)

def printRecur(top_node, root,parsed_path, parent_node_index, leaf_number):

    print("\r level: {:0>4} leaf at level 2 node {:0>4} - node: {:0>4} leaf:{:0>4}".format(0,0,parent_node_index,leaf_number),end="")
    global indent
    global df_stack
    global node_index
    local_index = 0
    node_index += 1
    remove_namespace = re.compile(r'\{.*\}(.*)')
    parsed_path += f"<{remove_namespace.findall( root.tag.title())[0]}[{leaf_number}]"
    logging.debug(' ' * indent + '%s: %s' % (root.tag.title(), root.attrib.get('name', root.text))+"\r")
    indent += 1
    elem_paraId = ""
    for i, elem in enumerate( root.getchildren()):
        local_index = i
        elem_paraId = elem.attrib.get('{http://schemas.microsoft.com/office/word/2010/wordml}paraId')
        tmp_se = pd.Series([
            parent_node_index, indent, local_index, remove_namespace.findall( root.tag.title())[0],remove_namespace.findall( elem.tag.title())[0].lower(),
            parsed_path + "<" + remove_namespace.findall( elem.tag.title())[0] + f"[{local_index}]" , f"/{elem.getroottree().getpath(elem)}",
            elem_paraId
                            ], index=df_stack.columns)
        df_stack = df_stack.append(tmp_se, ignore_index=True)

        printRecur(top_node,elem,parsed_path, node_index,i)
    if len(root.getchildren()) == 0:
        tmp_se = pd.Series([
            parent_node_index, indent, -1, remove_namespace.findall( root.tag.title())[0],None, parsed_path,
            f"/{root.getroottree().getpath(root)}",
            None
                            ], index=df_stack.columns)
        df_stack = df_stack.append(tmp_se, ignore_index=True)

    parsed_path += ">"
    logging.debug(parsed_path)

    indent -= 1

def extract_document_part(xml=""):
    document_part_string = ""
    compile_string_start_part = r'(<w:document.*?>)'
    pattern = re.compile(compile_string_start_part,        re.MULTILINE | re.DOTALL)
    found_xml = pattern.findall(xml)
    for index, part in enumerate(found_xml):
        document_part_string = part
    return document_part_string

def  extract_section(xml="", loop_index=None, base_path=None, target_tag = None):

    if target_tag == "tbl":
        target_tag_string = "tbl>"
        target_tag_string_end = target_tag
        target_tag_fast   = "<w:tbl>"
        target_tag_fast_end = "</w:tbl>"
    elif target_tag == "p":
        target_tag_string = "p .+?>"
        target_tag_string_end = target_tag
    elif target_tag == "tr":
        target_tag_string = "tr .+?>"
        target_tag_string_end = target_tag
        target_tag_fast   = "<w:tr "
        target_tag_fast_end = "</w:tr>"
    elif target_tag == "bookmarkend":
        target_tag_string = "bookmarkEnd .+?>"
        target_tag_string_end ="bookmarkEnd"
    else:
        target_tag_fast   = f"<w:{target_tag}>"
        target_tag_fast_end = f"</w:{target_tag}>"
        target_tag_string = target_tag + ">"
        target_tag_string_end = target_tag

    #TODO: target tag has to be loaded from mixedChartag (with or w/out IGNORECASE) and find()

    # find all nodes under the same path
    all_nodes_under_the_path_with_the_same_tag = df_stack[(df_stack["path"].str.contains(base_path, regex=False))
                                                          & (df_stack["tagself"] == target_tag)]
    # remove other parts from the tag
    compile_string_start = r'(<w:' + target_tag_string
    compile_string_insert = ""
    index = 0
    for i in all_nodes_under_the_path_with_the_same_tag.iterrows():
        index += 1
        if index == 1:
            continue
        compile_string_insert += r'.*?<w:' + target_tag_string + '.*?</w:' + target_tag +  '>'
    compile_string_end = r'.*?</w:' + target_tag_string_end + '>)'
    compile_string = compile_string_start + compile_string_insert + compile_string_end
    outer_pattern = re.compile(compile_string, re.MULTILINE | re.DOTALL| re.IGNORECASE)

    compile_string_inside_start = r'<w:' + target_tag_string + '('
    compile_string_inside_end = r'.*?)</w:' + target_tag_string_end + '>'
    compile_inside_string = compile_string_inside_start + compile_string_insert + compile_string_inside_end
    inside_pattern = re.compile(compile_inside_string, re.MULTILINE | re.DOTALL | re.IGNORECASE )

    extracted_part = ""
    extracted_inner_part = ""

    if index > 10:
        start_positon=0
        for i in range(1,index+1,1):
            start_positon = xml.find(target_tag_fast,start_positon) + 1
        end_position = xml.find(target_tag_fast_end, start_positon)
        end_position = xml.find(target_tag_fast_end, end_position + 1)
        #TODO: for the nested ones, upper-leve closing must be found. It might be necessary to find multi-levels

        end_position += len(target_tag_fast_end)
        extracted_part = xml[0:end_position]
        extracted_inner_part = "test"
    else:
        logging.debug(f"extraction with the strings {compile_string} ")
        mached_xml = outer_pattern.findall(xml)
        logging.debug(f"matched_xml: {len(mached_xml)} ")
        for index, part in enumerate(mached_xml):
            extracted_part += part
            break
        if len(mached_xml) == 0:
            logging.info(f"\n\rNOT FOUND {target_tag} at {base_path} of object {loop_index}")

        logging.debug(f"extraction with the strings -inside {compile_inside_string}")
        mached_xml = inside_pattern.findall(xml)
        logging.debug(f"matched_xml: {len(mached_xml)} ")
        for index, part in enumerate(mached_xml):
            extracted_inner_part += part
            break
        if len(mached_xml) == 0:
            logging.info(f"\n\rNOT FOUND {target_tag} at {base_path} of object {loop_index}")

    return extracted_part,extracted_inner_part

def expand_further(base_path, depth, target_tag,extracted_inner_section, tag_set_opening, tag_set_close,parent_loop_index, parent_path):

    global graph_paths

    all_nodes_under_the_path = df_stack[
        (df_stack["path"].str.contains(base_path, regex=False))&(df_stack["Depth"] == depth) ]
    inner_remaining_part = extracted_inner_section
    previous_tag = target_tag
    if depth > 10:
        logging.debug("over level 10: stopped")
        return False
    if depth == 3:
        list_to_expand = range(1, len(df_stack[df_stack["Depth"] == depth])+1,1)
        pbar2 = tqdm(total=len(list_to_expand), desc='Creating section files level 3', position=0, leave=False)
    for i,row in all_nodes_under_the_path.iterrows():
        logging.debug(f"tagself: {row.tagself}")
        if depth == 3:
            pbar2.update(1)
        if pd.isna(row.tagself) is True:
            continue
        target_tag = row.tagself.lower()
        base_path = str(row.path)
        loop_index = i
        paraId = row.paraId

        current_tag_opening = None
        current_tag_closing = None
        compile_string = r'(<w:' + target_tag + '.*?>)'
        # target_tag might need to include one space character after the tag
        outer_pattern = re.compile(compile_string, re.MULTILINE | re.DOTALL | re.IGNORECASE)
        mached_tags = outer_pattern.findall(extracted_section)
        logging.debug(f"matched_tag: {len(mached_tags)} ")
        for index, part in enumerate(mached_tags):
            current_tag_opening = part
            break
        if current_tag_opening is None:
            current_tag_opening = f"<w:{target_tag}>"

        compile_string = r'(</w:' + target_tag + '>)'
        # target_tag might need to include one space character after the tag
        outer_pattern = re.compile(compile_string, re.MULTILINE | re.DOTALL | re.IGNORECASE)
        mached_tags = outer_pattern.findall(extracted_section)
        logging.debug(f"matched_tag: {len(mached_tags)} ")
        for index, part in enumerate(mached_tags):
            current_tag_closing = part
            break
        if current_tag_closing is None:
            current_tag_closing = f"</w:{target_tag}>"

        logging.debug(f"TEST:{loop_index} {inner_remaining_part[0:10]}")
        extracted_section2, extracted_inner_section2 = extract_section(xml=inner_remaining_part, loop_index=loop_index, base_path=base_path,
                                            target_tag=target_tag)
        logging.debug(f"tag:{base_path},target_tag:{target_tag},loopindex{loop_index}")
        logging.debug(extracted_section2)
        #add extracted section of upper level
        inner_remaining_part = inner_remaining_part.replace(extracted_section2, "") #TODO: only once
        updated_document = document_header + \
                           f"<!-- {base_path} -->" + \
                           f"<!-- {paraId} -->" + \
                           tag_set_opening + \
                           extracted_section2 + tag_set_close + document_footer
        with open(os.path.join(out_folder, "word/document.xml"), 'wb') as f:
            f.write(updated_document.encode('utf-8'))
        with open(os.path.join(out_zip_folder, "document_{:0>2}_{:0>4}_{:0>4}.xml".format(str(depth),str(loop_index),str(parent_loop_index))), 'wb') as f:
            f.write(updated_document.encode('utf-8'))
        file_to_create = 'new_{:0>2}_{:0>4}_{:0>4}.docx'.format(str(depth),str(loop_index),str(parent_loop_index))
        open_result = zip_and_test(file_to_create = file_to_create, tagname=target_tag)
        if depth >= 4:
            graph_paths.append([str(parent_loop_index), str(loop_index)])
            graph_text[str(loop_index)] = base_path.replace("//w:document/w:body/","").replace("/","\n")
        if open_result == False:
            if depth < 4:
                graph_paths.append([str(parent_loop_index), str(loop_index)])
                graph_text[str(loop_index)] = base_path.replace("//w:document/w:body/","").replace("/","\n")
            expand_further(base_path, depth + 1, target_tag, extracted_inner_section,tag_set_opening + current_tag_opening, current_tag_closing + tag_set_close,loop_index,base_path )

def nudge(pos, x_shift, y_shift):
    return {n:(x + x_shift, y + y_shift) for n,(x,y) in pos.items()}

if __name__ == '__main__':

    arguments = docopt(__doc__, version="0.1")

    in_file = arguments["<in_file>"]
    out_folder = arguments["<out_folder>"]
    out_zip_folder = arguments["<out_zipfolder>"]
    debug_level = arguments["<debug_level>"]
    remove_styles = arguments["<remove_sytles>"]
    xml_original_filename = in_file

    logger = logging.getLogger(__name__)

    if debug_level == "DEBUG":
        logging.basicConfig(level=logging.DEBUG)
    elif debug_level == "WARNING":
        logging.basicConfig(level=logging.WARNING)
    else:
        logging.basicConfig(level=logging.INFO)


    os.makedirs(out_zip_folder,exist_ok=True)


    wd_app = win32com.client.Dispatch('Word.Application')
    logger.debug("Running check")

    loop_index = None
    document_master:str = create_docx(loop_index)
    file_to_create = 'new_{:0>2}_{:0>4}.docx'.format(str(1),str(loop_index))
    zip_and_test(file_to_create = file_to_create)
    with open(os.path.join(out_zip_folder, "document_{:0>2}_{:0>4}.xml".format(str(1),str(loop_index))), 'wb') as f:
        f.write(document_master.encode('utf-8'))

    df_stack = pd.DataFrame(
        columns=['lv1','Depth', 'Leaf', 'tag', 'tagself','tagstring','path','paraId'])
    indent = 0
    node_index = 0

    document_master = complement_empty_element(document_master=document_master)
    if os.path.exists("df_stack.csv") == False:
        logging.info("A new analysis file will be created under the name of df_stack.csv")
        analyze_create_splitter(document_master=document_master)
        df_stack.to_csv("df_stack.csv")
    else:
        logging.info("Using analysis file already available under the name of df_stack.csv. Converter may fail if older CSV file is used")
        df_stack = pd.read_csv("df_stack.csv")

    list_to_expand = range(1, len(df_stack[df_stack["Depth"] == 2])+1,1)
    pbar = tqdm(total=len(list_to_expand), desc='Creating section files top level', position=0, leave=False)
    document_header = extract_document_part(document_master) + "<w:body>"
    document_footer = "</w:body></w:document>"
    document_remaining_part = document_master
    document_remaining_part = document_remaining_part.replace(document_header,"")
    document_remaining_part = document_remaining_part.replace(document_footer,"")
    graph_paths = []
    graph_text  = {}
    for i,row in df_stack[df_stack["Depth"] == 2].iterrows():
        target_tag = row.tagself.lower()
        base_path  = str(row.path)
        loop_index = i
        pbar.update(1)
        extracted_section, extracted_inner_section = extract_section(xml=document_remaining_part, loop_index=loop_index, base_path=base_path, target_tag = target_tag)
        document_remaining_part = document_remaining_part.replace(extracted_section,"")
        logging.debug(f"tag:{base_path},loopindex{loop_index}")
        logging.debug(extracted_section)
        paraId = row.paraId
        updated_document = document_header + \
                           f"<!-- {base_path} -->" + \
                           f"<!-- {paraId} -->" + \
                           extracted_section + document_footer
        with open(os.path.join(out_folder, "word/document.xml"), 'wb') as f:
            f.write(updated_document.encode('utf-8') )
        with open(os.path.join(out_zip_folder, "document_{:0>2}_{:0>4}.xml".format(str(2),str(loop_index))), 'wb') as f:
            f.write(updated_document.encode('utf-8') )
        file_to_create = 'new_{:0>2}_{:0>4}.docx'.format(str(2),str(loop_index))
        open_result = zip_and_test(file_to_create = file_to_create, tagname= target_tag)
        if open_result == False:
            graph_paths.append([target_tag,str(loop_index)])
            graph_text[target_tag]= "Top"
            graph_text[str(loop_index)] = base_path.replace("//w:document/w:body/","").replace("/","\n")

            current_tag_opening = None
            compile_string = r'(<w:' + target_tag + '.*?>)'
            # target_tag might need to include one space character after the tag
            outer_pattern = re.compile(compile_string, re.MULTILINE | re.DOTALL | re.IGNORECASE)
            mached_tags = outer_pattern.findall(extracted_section)
            logging.debug(f"matched_tag: {len(mached_tags)} ")
            for index, part in enumerate(mached_tags):
                current_tag_opening = part
                break
            if current_tag_opening is None:
                current_tag_opening = f"<w:{target_tag}>"

            compile_string = r'(</w:' + target_tag + '>)'
            # target_tag might need to include one space character after the tag
            outer_pattern = re.compile(compile_string, re.MULTILINE | re.DOTALL | re.IGNORECASE)
            mached_tags = outer_pattern.findall(extracted_section)
            logging.debug(f"matched_tag: {len(mached_tags)} ")
            for index, part in enumerate(mached_tags):
                current_tag_closing = part
                break
            if current_tag_closing is None:
                current_tag_closing = f"</w:{target_tag}>"

            expand_further(base_path, 3, target_tag, extracted_inner_section,
                           current_tag_opening, current_tag_closing, loop_index, base_path)

    wd_app.Quit()

    plt.figure(figsize=(32, 32))
    G = nx.DiGraph(engine='diagram')
    for path in graph_paths:
        nx.add_path(G, path)
    pos = graphviz_layout(G, prog='dot')
    pos_nodes = nudge(pos, 0, 10)
    nx.draw_networkx_labels(G, pos=pos_nodes, labels=graph_text, verticalalignment="bottom")
    nx.draw(G, pos=pos,
            node_color='lightgreen',
            node_size=10000,
            node_shape="s",
            with_labels=True,
            arrows=True,
            verticalalignment = "top")
    filename = "tree.png"
    plt.savefig(filename)
    logging.info(f"File {filename} is created. The file shows an object tree for broken sections.")