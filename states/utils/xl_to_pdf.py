# import clr
# import sys
import os.path

# clr.AddReference("Microsoft.Office.Interop.Excel")

from Microsoft.Office.Interop import Excel
from System import Type, GC

# See http://msdn.microsoft.com/en-us/library/bb407651.aspx for
# reference


class ExcelWorkbook(object):
    """
    Workbook abstraction.
    Handles workbook opening, closing and rendering to PDF.
    """

    def __init__(self, app, path):
        "Init with Excel.ApplicationClass object and input path"
        self._wb = None
        self._app = app
        self._path = os.path.abspath(path)

    def __enter__(self):
        "Opens the workbook"
        self._wb = self._app.Workbooks.Open(self._path, 0, True)
        return self

    def __exit__(self, type, value, traceback):
        "Closes the workbook"
        if self._wb is not None:
            self._wb.Close(False)
            self._wb = None

    def toPDF(self, path):
        "Exports the workbook to path as PDF"
        form = Excel.XlFixedFormatType.xlTypePDF
        self._wb.ExportAsFixedFormat(form, os.path.abspath(path))

class ExcelApp(object):
    """
    Excel application abstraction.
    Handles app creation, termination, cleanup and creates workbooks for
    path given.
    """

    def __init__(self):
        "Just init the members"
        self._app = None

    def __enter__(self):
        "Opens the Excel app"
        self._app = Excel.ApplicationClass()
        # Disable the "file not found" alerts and such.
        self._app.DisplayAlerts = False
        return self

    def __exit__(self, type, value, traceback):
        "Quit the app and perform garbage collection"
        if self._app is not None:
            self._app.Quit()
            self._app = None
        GC.Collect()
        GC.WaitForPendingFinalizers()
        GC.Collect()
        GC.WaitForPendingFinalizers()

    def wb(self, path):
        "Returns the workbook object for given path"
        return ExcelWorkbook(self._app, path)


def convert_to_pdf(inpath, outpath):
    result = 2
    try:
        with ExcelApp() as app:
            with app.wb(inpath) as wb:
                wb.toPDF(outpath)
        result = 0
    except Exception as e:
        print (e) 
    return result