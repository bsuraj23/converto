converto
========

CGI-based file conversion server

Provides a simple HTTP API to convert documents from one format to another.
Makes use of various ActiveX components, so runs on Windows only (obviously).

######Example of usage:
```
curl -X POST \
     -H "Accept: application/pdf" \
     -H "Content-Type: application/msword" \
     --file-upload report.doc \
     -o report.pdf \
     http://localhost:8080/cgi-bin/convert.py
```

######Converters available :

* msoffice2pdf
    - Input: .doc, .docx, .xls, .xlsx, .ppt, .pptx
    - Output: PDF
    - Requires Microsoft Office to be installed (Word, Excel & PowerPoint).
  
* brava2pdf
    - Input: .dwg, .dwf, .dwfx
    - Output: PDF
    - Requires [IGC Brava Desktop](http://www.bravaviewer.com/brava-desktop) to be installed.
