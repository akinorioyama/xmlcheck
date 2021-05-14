`xmlcheck` is a tentative project to convert Word doc format file to docx file.

The project helps users identify errors in doc file by converting doc format to docx format, then running a verification in the tool called "Open XML SDK productivity tool".
![check result for incorrect document](examples/Error%20in%20the%20sample%20file%20sample_broken.doc.jpg)

Designed for Windows. *Not yet tested for other platforms.*

**Usage**
---

```
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

```

**Limitations**
---
The program is not exhaustive to cover all the content types and may miss creating a link for the corresponding file in the docx file. 

***Symptom:***
"Specified part does not exist in the package." appears in the Productivity Tool. In other languages, the message reads differently, which could be "指定されたパートはパッケージに存在しません。" in Japanese language.

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
