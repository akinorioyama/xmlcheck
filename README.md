`xmlcheck` is a tentative project to convert Word doc format file to docx file.

The project helps users identify errors in doc file by converting doc format to docx format, then running a verification in the tool called "Open XML SDK productive tool".
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
