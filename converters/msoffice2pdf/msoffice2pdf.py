import sys
import win32com.client

# WARNING 1 : saving Excel (and PowerPoint ?) documents as PDF requires to have a printer installed.
# (e.g. "Local printer" >> "Generic" manafacturer >> "Text only" model)
# <http://social.technet.microsoft.com/Forums/en-US/officesetupdeploylegacy/thread/6756e463-78e3-4bef-80f0-dc8438eceb1e/>

# WARNING 2 : Office API constants are made available with "COM Makepy utility" tool run from PythonWin.
# From the list, choose:
#   Microsoft Excel 14.0 Object Library
#   Microsoft PowerPoint 14.0 Object Library
#   Microsoft Word 14.0 Object Library
#   Microsoft Office 14.O Object Library
# => NOT required if using gencache.EnsureDispatch()

try:
    return_code = 0
    win32constants = win32com.client.constants

    input_file_extension = sys.argv[1]
    input_content_type = sys.argv[2]
    input_file_path = sys.argv[3]
    output_file_path = sys.argv[4]

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
        # officeApp = win32com.client.Dispatch('Excel.Application')
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
        # officeApp = win32com.client.Dispatch('PowerPoint.Application')
        officeApp = win32com.client.gencache.EnsureDispatch('PowerPoint.Application')
        # officeApp.Visible = 1
        
        presentation = officeApp.Presentations.Open(
            FileName = input_file_path,
            ReadOnly = win32constants.msoTrue,
            # MsoTriState Untitled, # what effect does it have??
            WithWindow = win32constants.msoFalse,
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
            PrintHiddenSlides = win32constants.msoTrue,
            IncludeDocProperties = True,
            DocStructureTags = True,
            PrintRange = None, 
        )

        presentation.Close()
        presentation = None
except Exception as exc:
    sys.stderr.write(exc)
    return_code = 1
finally:
    try:
        officeApp.Quit()
    except:
        pass

    sys.exit(return_code)