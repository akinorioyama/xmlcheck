`xmlcheck` is a tentative project to convert Word doc format file to docx file and to identify errors that stops opening a Word file.

The project helps users identify errors in doc file by converting doc format to docx format, then running a verification in the tool called "Open XML SDK productivity tool". The program also identify errors by splitting doc format file to docx and xml files that only contain error-causing sections.
![check result for incorrect document](examples/Error%20in%20the%20sample%20file%20sample_broken.doc.jpg)

The out_zipfolder for the <b>examples/sample_broken.doc</b> will have the following files from a doc file among other files. Docx files with "broken" in the middle of filename indicate the part causes error(s) while other docx files are error free. Opening "broken" files will issue Word error showing "Word found unreadable content in <b>filename</b>. Do you want to recover the contents of this document?"
```
document_01_None.xml               <-- top most XML that contains all parts of the file
document_02_0173.xml               <-- XML that has errornous part (level 02 and part 0173)
document_03_0178_0173.xml          <-- XML that has errornous part (level 03 and part 0178 under part 0173)
new_01_None.docxbroken_.docx       <-- damaged docx file that contains all parts of the file
new_02_0173.docxbroken_p.docx      <-- damaged docx file that has errornous part (level 02 and part 0173)
new_03_0178_0173.docxbroken_r.docx <-- damaged docx file that has errornous part (level 03 and part 0178 under part 0173)
```
The concerned nodes and levels of the above file has the following hierarchy. <b>df_stack.csv</b> file in the execution folder holds the hiearchy information along with the part ID.
```
//w:document                (part None)
//w:document/w:body         (part 0173)
//w:document/w:body/w:p[5]  (part 0178 under 0173)
```

Designed for Windows. *Not yet tested for other platforms.*

**Usage**
---

```
Convert doc to docx

Usage:
  decompose_regex.py <in_file> <out_folder> <out_zipfolder> <debug_level>
  decompose_regex.py -h | --help
  decompose_regex.py --version

  <in_file>: filename of the doc file to be converted
  <out_folder>: output folder for the split files of docx file
  <out_zipfolder>: output folder for zipped file named (new.docx) and
    segment docx and XML files. 
  <debug_level>: INFO, DEBUG to configure the levels of logging messages    

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

```

**Limitations**
---
The program is not exhaustive to cover all the content types and may miss creating a link for the corresponding file in the docx file. 

***Symptom:***
"Specified part does not exist in the package." appears in the Productivity Tool.

***Cause:***
"\[ContentTypes\].xml" in the root folder of the docx (zipped) file defines the types in the document. Lacking any content type results in "corrupted" eror message in Word application. If a Word doc file has an inserted jpeg image part and \[ContentTypes\] has no associated entry, Word issues the error message.

***Solution:***
Add code in *decompose_regex.py* to handle the missing content type in the "\[ContentTypes\].xml". For example, you may add jpeg as follows.
```
f.write('<Default Extension="jpeg" ContentType="image/jpeg"/>')
```

It will be helpful to compile *DocumentFormat.OpenXml* as Debug from the source(https://github.com/OfficeDev/Open-XML-SDK) so that the missing part can be identified. Most probably,  LoadReferencedPartsAndRelationships in OpenXmlPartContainer detects the missing link.


**Useful tools and references**
---
* Open XML SDK page (https://docs.microsoft.com/en-us/office/open-xml/open-xml-sdk).
* Open XML SDK 2.5 Productivity Tool (choose OpenXMLSDKToolV25.msi from https://www.microsoft.com/en-us/download/details.aspx?id=30425)

**Errors in other languages**
---
(Japanese language)

* "Specified part does not exist in the package." -> "指定されたパートはパッケージに存在しません。"
* "Word found unreadable content in <b>filename</b>. Do you want to recover the contents of this document?" -> "<b>filename</b>は破損しています。プログラムのエラーの結果破損したか、悪意のある作成元が意図的に破損させた可能性があります。"
