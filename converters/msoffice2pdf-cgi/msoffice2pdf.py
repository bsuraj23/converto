#import cgi
import shutil
import os
import sys
import tempfile
import uuid
import win32com.client
# import pythoncom

# import cgitb
# cgitb.enable(display=0, logdir="/path/to/logdir")

win32constants = win32com.client.constants

# xlApp = win32com.client.Dispatch("BRAVADTX.BravaDTXView")
# xlApp.AllowFileOpen = True
# xlApp.Filename = r'C:\Test.dwg'

#workBook = xlApp.Workbooks.Open(r"C:\MyTest.xls")
#print str(workBook.ActiveSheet.Cells(i,1))
#workBook.ActiveSheet.Cells(1, 1).Value = "hello"                
#workBook.Close(SaveChanges=0) 
#xlApp.Quit()

# WARNING 1 : saving Excel (and PowerPoint ?) documents as PDF requires to have a printer installed.
# (e.g. "Local printer" >> "Generic" manafacturer >> "Text only" model)
# <http://social.technet.microsoft.com/Forums/en-US/officesetupdeploylegacy/thread/6756e463-78e3-4bef-80f0-dc8438eceb1e/>

# WARNING 2 : Office API constants are made available with "COM Makepy utility" tool run from PythonWin.
# From the list, choose:
#   Microsoft Excel 14.0 Object Library
#   Microsoft PowerPoint 14.0 Object Library
#   Microsoft Word 14.0 Object Library
#   Microsoft Office 14.O Object Library

#for application in ('Excel', 'PowerPoint', 'Word'):
#    win32com.client.gencache.EnsureDispatch(application + '.Application')



CRLF = '\r\n'
EXTENSIONS_FOR_CONTENT_TYPE = {
    # Word
    'application/msword': 'doc',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document': 'docx',
    # Excel
    'application/vnd.ms-excel': 'xls',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': 'xlsx',
    # PowerPoint
    'application/vnd.ms-powerpoint': 'ppt',
    'application/vnd.openxmlformats-officedocument.presentationml.presentation': 'pptx',
}

try:
    content_length = os.environ.get('CONTENT_LENGTH')
    if content_length is None:
        raise 'error: Content-Length is absent'
    content_length = int(content_length)

    content_type = os.environ.get('CONTENT_TYPE')
    if content_type is None:
        raise Exception('Content-Type header is absent.')
    elif content_type not in EXTENSIONS_FOR_CONTENT_TYPE:
        raise Exception('Invalid Content-Type header.')

    input_file_extension = \
        EXTENSIONS_FOR_CONTENT_TYPE[content_type]
    output_file_extension = 'pdf'

    tmp_dir = tempfile.mkdtemp()

    input_file_path = os.path.join(
        tmp_dir, uuid.uuid4().hex[:8] + '.' + input_file_extension
    )
    output_file_path = os.path.join(
        tmp_dir, uuid.uuid4().hex[:8] + '.' + output_file_extension
    )

    with open(input_file_path, 'wb') as f:
        shutil.copyfileobj(sys.stdin, f)

    officeApp = None
    
    if input_file_extension in ('doc', 'docx'):
        officeApp = win32com.client.gencache.EnsureDispatch('Word.Application')
        # officeApp = win32com.client.Dispatch('Word.Application')
        # officeApp.Visible = 1

        document = officeApp.Documents.Open(
            FileName = input_file_path,
            ConfirmConversions = False,
            ReadOnly = True,
            AddToRecentFiles = False,
        )
        
        document.ExportAsFixedFormat(
            OutputFileName = output_file_path,
            ExportFormat = win32constants.wdExportFormatPDF,
            OptimizeFor = win32constants.wdExportOptimizeForPrint,
            IncludeDocProps = True,
            DocStructureTags = True,
        )

        document.Close()
        document = None
    elif input_file_extension in ('xls', 'xlsx'):
        officeApp = win32com.client.gencache.EnsureDispatch('Excel.Application')
        # officeApp.Visible = 1

        workbook = officeApp.Workbooks.Open(
            Filename = input_file_path,
            ReadOnly = True,
        )

        workbook.ExportAsFixedFormat(
            Filename = output_file_path,
            Type = win32constants.xlTypePDF,
            Quality = win32constants.xlQualityStandard,
            IncludeDocProperties = True,
            IgnorePrintAreas = False,
        )

        workbook.Close()
        workbook = None
    elif input_file_extension in ('ppt', 'pptx'):
        # error when trying to access win32constants.mso{True, False} ?!?
        msoTrue = -1
        msoFalse = 0

        # officeApp = win32com.client.Dispatch('PowerPoint.Application')
        officeApp = win32com.client.gencache.EnsureDispatch('PowerPoint.Application')
        # officeApp.Visible = 1
        
        presentation = officeApp.Presentations.Open(
            FileName = input_file_path,
            ReadOnly = msoTrue,
            # MsoTriState Untitled, # what effect does it have??
            WithWindow = msoFalse,
        )
        
        # DON'T remove "PrintRange" parameter,
        # otherwise ExportAsFixedFormat() call will fail!
        # <http://sourceforge.net/p/pywin32/bugs/339/>
        presentation.ExportAsFixedFormat(
            Path = output_file_path,
            FixedFormatType = win32constants.ppFixedFormatTypePDF,
            Intent = win32constants.ppFixedFormatIntentPrint,
            # MsoTriState FrameSlides,  # draws a border around slides
            OutputType = win32constants.ppPrintOutputSlides,
            PrintHiddenSlides = msoTrue,
            IncludeDocProperties = True,
            DocStructureTags = True,
            PrintRange = None, 
        )

        presentation.Close()
        presentation = None
    
    print 'Content-Type: application/pdf' + CRLF
    print CRLF

    with open(output_file_path, 'rb') as f:
        shutil.copyfileobj(f, sys.stdout)
except Exception as exc:
    # writing to STDOUT fails if STDIN hasn't been read!?
    sys.stdin.read()
    
    print 'Status: 500 Internal Server Error' + CRLF
    print 'Content-Type: text/plain' + CRLF
    print CRLF
    print exc
finally:
    try:
        officeApp.Quit()
    except:
        pass

    try:
        shutil.rmtree(tmp_dir)
    except:
        pass
