import os
import shutil
import subprocess
import sys
import tempfile
import uuid

CRLF = '\r\n'
EXTENSIONS_FOR_CONTENT_TYPE = {
    #
    # Microsoft Office
    #
    
    # Word
    'application/msword': 'doc',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document': 'docx',
    # Excel
    'application/vnd.ms-excel': 'xls',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': 'xlsx',
    # PowerPoint
    'application/vnd.ms-powerpoint': 'ppt',
    'application/vnd.openxmlformats-officedocument.presentationml.presentation': 'pptx',

    #
    # IGC Brava
    #

    # DWG
    'application/acad': 'dwg',
    'application/autocad.dwg': 'dwg',
    'application/dwg': 'dwg',
    'image/dwg': 'dwg',
    'image/vnd.dwg': 'dwg',
    'vector/x-dwg': 'dwg',
    # DWF
    'drawing/x-dwf': 'dwf',
    'image/vnd.dwf': 'dwf',
    'model/vnd.dwf': 'dwf',
    # DWFX
    'model/vnd.dwfx+xps': 'dwfx',
}

def _executable_is_python(path):
    head, tail = os.path.splitext(path)
    return tail.lower() in (".py", ".pyw")

try:   
    input_content_length = os.environ.get('CONTENT_LENGTH')
    if input_content_length is None:
        raise 'error: Content-Length is absent'
    input_content_length = int(input_content_length)

    input_content_type = os.environ.get('CONTENT_TYPE')
    if input_content_type is None:
        raise Exception('Content-Type header is absent.')
    elif input_content_type not in EXTENSIONS_FOR_CONTENT_TYPE:
        raise Exception('Invalid Content-Type header.')

    output_content_type = os.environ.get('HTTP_ACCEPT')
    if output_content_type is None:
        raise Exception('Accept header is absent.')
    output_content_type = output_content_type.strip()
    if output_content_type != 'application/pdf':
        raise Exception('Invalid Accept header.')
    
    input_file_extension = \
        EXTENSIONS_FOR_CONTENT_TYPE[input_content_type]
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

    executable = None

    if input_file_extension in ('doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx'): # rtf ?
        executable = 'msoffice2pdf.py'
    elif input_file_extension in ('dwg', 'dwf', 'dwfx'):
        executable = 'brava2pdf.exe'

    executable_path = os.path.join(
        os.path.dirname(__file__),
        executable
    )
    
    cmdline = [
        executable_path,
        input_file_extension,
        input_content_type,
        input_file_path,
        output_file_path
    ]

    # from CGIHTTPServer.py
    if _executable_is_python(executable_path):
        interpreter = sys.executable
        if interpreter.lower().endswith("w.exe"):
            # On Windows, use python.exe, not pythonw.exe
            interpreter = interpreter[:-5] + interpreter[-4:]
        cmdline = [interpreter, '-u'] + cmdline

    process = subprocess.Popen(
        cmdline, stderr = subprocess.PIPE
    )
    (stdout_data, stderr_data) = process.communicate()

    if process.returncode != 0:
        raise Exception('Error while running executable: ' + stderr_data)

    print 'Content-Type: ' + output_content_type + CRLF
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
        shutil.rmtree(tmp_dir)
    except:
        pass
