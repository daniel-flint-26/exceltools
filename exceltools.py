"""
exceltools - providing more user-friendly access to the pywin32 library
=============================================================================

exceltools is a Python module acting as a friendlier interface to the pywin32 library which
in itself is an API to the Windows COM client API.

exceltools does not provide the full functionality of pywin32, it only seeks to simplify some commonly used code.

exceltools is intended to work alongside pandas and numpy and aids in creating and
populating spreadsheets programmatically.

"""

import datetime as dt
import os
import re
import sys
import shutil
import warnings
import ast
from time import sleep
from typing import Union
from pathlib import Path
import pandas as pd
import pandas.api.types as types
import numpy as np
from win32com import client
from win32com.client import constants as c
import pythoncom

def col2num(col_str: str) -> int:
    """
    Convert an Excel column reference to an integer
    e.g. "A" = 1, "B" = 2 e.t.c.
    """
    if not isinstance(col_str, str):
        raise ValueError("Invalid data type supplied. Must supply a scalar string")
    expn = 0
    col_num = 0
    for char in reversed(col_str):
        col_num += (ord(char) - ord("A") + 1) * (26 ** expn)
        expn += 1
    return col_num

def num2col(col_int: int) -> str:
    """
    Convert an Excel column index to a string
    e.g. 1 == "A", 27 == "AA" e.t.c.
    """
    if not isinstance(col_int, int):
        raise ValueError("Invalid data type supplied. Must supply an integer")
    col_str = ""
    while col_int > 0:
        col_int, remainder = divmod(col_int - 1, 26)
        col_str = chr(65 + remainder) + col_str
    return col_str

def rgb2hex(rgb: Union[list, tuple]) -> str:
    """
    Excel expects a hex value in order to fill cells
    This function allows you to supply standard RGB values to be converted to hex.
    """

    if not isinstance(rgb, (tuple, list)):
        raise TypeError("Argument supplied must be a tuple or list of RGB values")
    bgr = (rgb[2], rgb[1], rgb[0])
    str_value = "%02x%02x%02x" % bgr
    hexcode = int(str_value, 16)
    return hexcode

def excel_date(date1: Union[pd.Series, dt.datetime, dt.date]) -> float:
    """
    Convert a datetime.datetime or pandas.Series object into an Excel date float
    """
    if isinstance(date1, (dt.datetime, dt.date)):
        if isinstance(date1, dt.date):
            date1 = dt.datetime.combine(date1, dt.datetime.min.time())
        temp = dt.datetime(1899, 12, 30)    # Excels epoch. Note, not 31st Dec but 30th
        delta = date1 - temp
        return float(delta.days) + (float(delta.seconds) / 86400)
    elif isinstance(date1, pd.Series):
        temp  = pd.Timestamp(dt.datetime(1899, 12, 30))
        delta = date1 - temp
        return (delta.dt.days)+ (delta.dt.seconds / 86400)
    else:
        raise TypeError("Must supply datetime, date or pd.Series")

class Error(Exception):
    """Base Class for Exceptions"""
    def __init__(self, msg, *args):
        Exception.__init__()
        self.what = msg.format(*args)

    def __str__(self):
        return self.what

class NoWorkbookError(Exception):
    """Raised when there is no Excel Workbook Open/Active"""
    def __init__(self):
        Exception.__init__()
        self.what = "There are no active workbooks open"

    def __str__(self):
        return self.what

class InvalidSheetError(Error):
    """Raised when a worksheet does not exist/is invalid"""

class InvalidCellRefError(Error):
    """Raised when a cell reference is invalid"""

class InvalidRangeError(Error):
    """Raised when a range is invalid"""


class ExcelSpreadSheet():
    """
    A class built to simplify and streamline working with the win32client library.
    Example usage involves opening an existing workbook and saving a new copy without changing the original.
    New workbooks can also be created and originals can be overwritten.
    Example Usage:
        excel = ExcelSpreadSheet()
        excel.open("C:/Users/daflin/Documents/master_file.xlsx")
        excel.write_dataframe(data, sheet="Sheet 1", startcol=1, startrow=2, headers=True)
        excel.write_cell("SomeString", sheet=1, row=1, col="A")
        excel.save_xlsx("C:/Users/daflin/Documents/new_file.xlsx")
        excel.close(save_changes=False)
    """
    def __init__(self):
        global client
        try:
            self.excel = client.gencache.EnsureDispatch("Excel.Application")
        except Exception:
            # Remove cache and try again.
            module_list = [m.__name__ for m in sys.modules.values()]
            for module in module_list:
                if re.match(r"win32com\.gen_py\..+", module):
                    del sys.modules[module]
            shutil.rmtree(os.path.join(os.environ.get("LOCALAPPDATA"), "Temp", "gen_py"))
            from win32com import client
            self.excel = client.gencache.EnsureDispatch("Excel.Application")

        self.wb = None
        self.activesheet = None
        self.sheetnames = []
        self.null_arg = pythoncom.Empty
        self._wb_open = 0
        self._range_regex = re.compile(r"(^[a-zA-Z]{1,3})(\d+):([a-zA-Z]{1,3})(\d+$)|"
                                       r"(^[a-zA-Z]{1,3}:[a-zA-Z]{1,3}$)")
        self._cell_regex  = re.compile(r"^[a-zA-Z]{1}\d+$")
        self.format_args = {
            "Condition": {
                "logic": "logic_dict[logic]",
                "value": "value",
                "value2": "value2"
             },
            "Format": {
                "interior_colour": "Interior.Color = self.rgb2hex(kwargs['interior_colour'])",
                "number_format": "NumberFormat = kwargs['number_format']",
                "bold": "Font.Bold = kwargs['bold']",
                "font_colour": "Font.Color = self.rgb2hex(kwargs['font_colour'])",
                "font_size": "Font.Size = kwargs['font_size']",
                "font_name": "Font.Name = kwargs['font_name']",
                "orientation": "Orientation = kwargs['orientation']",
                "underline": "Font.Underline = kwargs['underline']",
                "merge": "MergeCells = kwargs['merge']",
                "wrap_text": "WrapText = kwargs['wrap_text']",
                "h_align": "HorizontalAlignment = kwargs['h_align']",
                "v_align": "VerticalAlignment = kwargs['v_align']",
                "border_left": {
                    "line_style": "Borders(c.xlEdgeLeft).LineStyle = kwargs['border_left']['line_style']",
                    "weight": "Borders(c.xlEdgeLeft).Weight = kwargs['border_left']['weight']",
                    "colour": "Borders(c.xlEdgeLeft).Color = self.rgb2hex(kwargs['border_left']['colour'])",
                    },
                "border_right": {
                    "line_style": "Borders(c.xlEdgeRight).LineStyle = kwargs['border_right']['line_style']",
                    "weight": "Borders(c.xlEdgeRight).Weight = kwargs['border_right']['weight']",
                    "colour": "Borders(c.xlEdgeRight).Color = self.rgb2hex(kwargs['border_right']['colour'])",
                    },
                "border_top": {
                    "line_style": "Borders(c.xlEdgeTop).LineStyle = kwargs['border_top']['line_style']",
                    "weight": "Borders(c.xlEdgeTop).Weight = kwargs['border_top']['weight']",
                    "colour": "Borders(c.xlEdgeTop).Color = self.rgb2hex(kwargs['border_top']['colour'])",
                    },
                "border_bot": {
                    "line_style": "Borders(c.xlEdgeBottom).LineStyle = kwargs['border_bot']['line_style']",
                    "weight": "Borders(c.xlEdgeBottom).Weight = kwargs['border_bot']['weight']",
                    "colour": "Borders(c.xlEdgeBottom).Color = self.rgb2hex(kwargs['border_bot']['colour'])",
                    },
                "border_inside_h": {
                    "line_style": "Borders(c.xlInsideHorizontal).LineStyle = kwargs['border_inside_h']['line_style']",
                    "weight": "Borders(c.xlInsideHorizontal).Weight = kwargs['border_inside_h']['weight']",
                    "colour": "Borders(c.xlInsideHorizontal).Color = self.rgb2hex(kwargs['border_inside_h']['colour'])",
                    },
                "border_inside_v": {
                    "line_style": "Borders(c.xlInsideVertical).LineStyle = kwargs['border_inside_v']['line_style']",
                    "weight": "Borders(c.xlInsideVertical).Weight = kwargs['border_inside_v']['weight']",
                    "colour": "Borders(c.xlInsideVertical).Color = self.rgb2hex(kwargs['border_inside_v']['colour'])",
                    }
            }
        }
    
 #   def __del__(self):
 #       self.excel.Quit()
    
    @staticmethod
    def col2num(col_str):
        """
        Convert an Excel column string to an integer -> "A" == 1, "AA" == 27 e.t.c.
        """
        expn = 0
        col_num = 0
        for char in reversed(col_str):
            col_num += (ord(char) - ord("A") + 1) * (26 ** expn)
            expn += 1
        return col_num

    @staticmethod
    def num2col(col_int):
        """
        Convert an Excel column index to a string -> 1 == "A", 27 == "AA" e.t.c.
        """
        col_str = ""
        while col_int > 0:
            col_int, remainder = divmod(col_int - 1, 26)
            col_str = chr(65 + remainder) + col_str
        return col_str

    @staticmethod
    def rgb2hex(rgb):
        """
        Excel expects a hex value in order to fill cells
        This function allows you to supply standard RGB values to be converted to hex.
        """
        if not isinstance(rgb, (tuple, list)):
            raise TypeError("Argument supplied must be a tuple or list of RGB values")
        bgr = (rgb[2], rgb[1], rgb[0])
        str_value = "%02x%02x%02x" % bgr
        hexcode = int(str_value, 16)
        return hexcode

    @staticmethod
    def excel_date(date1):
        """
        Convert a datetime.datetime or pandas.Series object into an Excel date float
        """
        if isinstance(date1, (dt.datetime, dt.date)):
            if isinstance(date1, dt.date):
                date1 = dt.datetime.combine(date1, dt.datetime.min.time())
            temp = dt.datetime(1899, 12, 30)    # Excels epoch. Note, not 31st Dec but 30th!
            delta = date1 - temp
            return float(delta.days) + (float(delta.seconds) / 86400)
        elif isinstance(date1, pd.Series):
            temp  = pd.Timestamp(dt.datetime(1899, 12, 30))
            delta = date1 - temp
            return (delta.dt.days)+ (delta.dt.seconds / 86400)
        else:
            raise TypeError("Must supply datetime, date or pd.Series")

    def _validate_column(self, col):
        """
        Checks that a column reference supplied is valid, and returns it if true.
        String references such as "AB" are returned as integers.
        """
        if col is None:
            return col
        if isinstance(col, str):
            if re.search(r"[^a-zA-Z0-9]", col):
                raise ValueError("Column reference must only contain alphanumeric characters"
                                 ", invalid column reference supplied")
            if len(col) > 3:
                raise ValueError("String must be no more than 3 characters")
            col = self.col2num(col)

        try:
            int(col)
        except ValueError as e:
            raise ValueError("Column reference could not be coerced to integer") from e

        if col > 18278:
            raise ValueError("Column reference is too large, 18278/\"ZZZ\" "
                             "is the maximum width accepted")

        return col

    def _validate_row(self, row):
        """
        Ensures the value supplied is a valid Excel row number
        """
        if row is None:
            return row
        try:
            row = int(row)
        except ValueError as e:
            raise ValueError("Could not coerce row value to integer") from e

        return row

    def _validate_workbook(self):
        """
        Ensure the current workbook is open and valid
        """
        if self._wb_open == 0:
            raise NoWorkbookError()

    def _validate_worksheet(self, sheet):
        """
        Make sure the sheet supplied is valid for the current open workbook
        """
        if isinstance(sheet, str):
            if sheet not in self.sheetnames:
                raise InvalidSheetError(f"A sheet with the name {sheet} does not exist")
        elif isinstance(sheet, int):
            if len(self.sheetnames) < sheet:
                raise InvalidSheetError(f"Invalid Sheet Index. Sheet index {sheet} is out of bounds.")

    def _validate_cellref(self, cellref, row, col):
        """
        Ensures the cellref supplied is a valid Excel cell reference -
        returns a tuple of row and col values to be used.
        """
        if all(value is None for value in (row, col, cellref)):
            raise InvalidCellRefError("Please supply co-ordinates to write to.")
        elif all(value is not None for value in (row, col, cellref)):
            raise InvalidCellRefError("Too many co-ordinates supplied."
                                      " Please supply either a cell reference or x and y values")
        elif (all(value is None for value in (row, col)) and cellref is not None
                and re.match(self._cell_regex, cellref) is None):
            raise InvalidCellRefError("Cell reference supplied is invalid.")
        else:
            if cellref is None:
                row = self._validate_row(row)
                col = self._validate_column(col)
                return row, col
            else:
                col, row = cellref[0], int(cellref[1])
                return row, col

    def _validate_range(self, _range, startrow, endrow, startcol, endcol):
        """
        Ensures the range supplied is a valid Excel range - 
        returns a string e.g. "A1:B2"
        """
        # Convert chars to ints
        [startcol, endcol] = [self._validate_column(i) for i in [startcol, endcol]]
        [startrow, endrow] = [self._validate_row(i)    for i in [startrow, endrow]]

        coords = [startrow, endrow, startcol, endcol]
        if _range is not None and all(coord is None for coord in coords):
            match = re.match(self._range_regex, _range)
            col_1 = match.group(1)
            row_1 = match.group(2)
            col_2 = match.group(3)
            row_2 = match.group(4)
            cols_only = match.group(5)

            if match is None:
                raise InvalidRangeError("range must be a valid Excel range string i.e. A1:B3 or A:A. Column references must be 3 chars max.")

            if cols_only is not None:
                return _range

            col_1 = self._validate_column(col_1)
            col_2 = self._validate_column(col_2)
            row_1 = self._validate_row(row_1)
            row_2 = self._validate_row(row_2)

            if col_1 > col_2:
                raise InvalidRangeError("Starting column cannot be greater than the ending column!")
            if row_1 > row_2:
                raise InvalidRangeError("Starting row cannot be greater than the ending row!")

            return _range
        else:
            if any(coord is not None for coord in coords) and any(coord is None for coord in coords):
                raise InvalidRangeError("All start and end col/row values must be supplied, only partial values detected.")
            if all(coord is not None for coord in coords) and _range is not None:
                raise InvalidRangeError("You cannot supply both an Excel range and start/end values. Please supply one or the other.")
            if startcol > endcol:
                raise InvalidRangeError("Starting column cannot be greater than the ending column!")
            if startrow > endrow:
                raise InvalidRangeError("Starting row cannot be greater than the ending row!")

            _range = str(self.num2col(startcol)) + str(startrow) + ":" + str(self.num2col(endcol)) + str(endrow)
            return _range

    def _cleanse_data(self, data):
        """
        Excel will print np.Nan as 65535.
        This function aims to cleanse any representations of NULL so that they print as expected to Excel.
        At this stage we also attempt to convert datetimes to a numeric value used by Excel.
        """
        if isinstance(data, pd.DataFrame):
            for column in data:

                _dtype = data[column].dtype

                if types.is_numeric_dtype(_dtype):
                    data.loc[:, column] = data[column].fillna(0)

                if types.is_string_dtype(_dtype):
                    data.loc[:, column] = data[column].fillna("")

                if types.is_datetime64_any_dtype(_dtype):
                    data.loc[:, column] = self.excel_date(data[column])

        elif isinstance(data, (pd.Series, list)):
            _dtype = pd.Series(data).dtype

            if types.is_numeric_dtype(_dtype):
                data = data.fillna(0)

            elif types.is_string_dtype(_dtype):
                data = data.fillna("")

            elif types.is_datetime64_any_dtype(_dtype):
                data = self.excel_date(data)

        return data

    def get_format_args(self):
        return list(self.format_args["Format"].keys())
        
    def open(self, file):
        """
        Open an Excel workbook. If the workbook is already open, no action is taken.
        If the workbook does not exist, a new one is created
        args:
            file : The name of the workbook to be open.
        """
        if self._wb_open == 1:
            raise Exception("Only one active workbook can be open at a time, please close the current workbook")
            
        file = os.path.normpath(Path(file))
        if not os.path.isfile(file):
            self.wb = self.excel.Workbooks.Add()
            self.wb.SaveAs(file)
            self.wb = self.excel.Workbooks.Open(file)
        else:
            try:
                self.wb = self.excel.Workbooks(file)
            except Exception:
                try:
                    self.wb = self.excel.Workbooks.Open(file)
                except Exception as e:
                    # Remove cache and try again.
                    module_list = [m.__name__ for m in sys.modules.values()]
                    for module in module_list:
                        if re.match(r"win32com\.gen_py\..+", module):
                            del sys.modules[module]
                    shutil.rmtree(os.path.join(os.environ.get("LOCALAPPDATA"), "Temp", "gen_py"))
                    from win32com import client
                    try:
                        self.excel = client.gencache.EnsureDispatch("Excel.Application")
                        self.wb = self.excel.Workbooks.Open(file)
                    except Exception as e:
                        print(str(e))
                        self.wb = None
                        raise Exception from e
        
        # Wait until application is ready and has opened the file
        sleep(1)

        if self.wb is not None:
            self.excel.Visible = False
            self.excel.Interactive = False
            self.excel.DisplayAlerts = False
            self.excel.ScreenUpdating = False
            self.excel.EnableEvents = False
            self.excel.DisplayStatusBar = False
            self._wb_open = 1
            for sheet in self.wb.Sheets:
                self.sheetnames.append(sheet.Name)

    def write_dataframe(self, data, sheet, cellref=None, startrow=None, startcol=None, headers=False):
        """
        Writes a pandas dataframe to an Excel worksheet. If the supplied data is not a dataframe the method will error.
        args:
            data : A pandas dataframe
            cellref: An Excel cell reference to start writing from - overrides any other arguements supplied
            startrow : the starting row to write data to (default=2)
            startcol : the starting column to write data to (default=1)
            headers  : Boolean flag, if true the column names of the dataframe will be printed (default=False)
            sheet : the name or index of the sheet to be populated, must be a valid sheet in the active Workbook.
        """
        self._validate_workbook()
        self._validate_worksheet(sheet)
        startrow, startcol = self._validate_cellref(cellref, startrow, startcol)

        if not isinstance(data, pd.DataFrame):
            raise ValueError("Data supplied must be a pandas dataframe")

        data = self._cleanse_data(data)

        startcol = self._validate_column(startcol)

        self.excel.Calculation=c.xlCalculationManual
        sheet = self.wb.Sheets(sheet)
        
        if sheet.ProtectContents:
            raise InvalidSheetError(f"The sheet '{sheet.Name}' is protected, please unprotect before attempting to write to it")

        # Write the data
        if headers:
            headers = np.array(data.columns)
            sheet.Range(sheet.Cells(startrow, startcol), sheet.Cells(startrow,len(headers)+(startcol-1))).Value = headers
            sheet.Range(sheet.Cells((startrow+1), startcol), sheet.Cells((len(data)+(startrow)),(len(headers)+(startcol-1)))).Value = data.values
        else:
            sheet.Range(sheet.Cells(startrow, startcol), sheet.Cells((len(data)+(startrow-1)),(len(data.columns)+(startcol-1)))).Value = data.values

        self.excel.Calculate()
        self.excel.Calculation = c.xlCalculationAutomatic

    def write_cell(self, data, sheet, row=None, col=None, cellref=None):
        """
        Write scalar data to a specific Cell in a workbook. Non-Scalar data will attempt to be coerced into a comma seperated string.
        A Set cannot be passed to this method as it is un-ordered data.
        row is a row number, col can be supplied as a string Excel Reference i.e. "A" or column index.
        An Error should be returned if the object passed cannot be written.
            args:
                data : Variable to write to the cell
                cellref : An Excel cell reference to write to
                x : A row number to write to
                y : A column name or index to write to
                sheet : The sheet name or index to write to
        """

        if isinstance(data, (pd.DataFrame, set)):
            raise ValueError("Data supplied cannot be a pandas dataframe or set")
        self._validate_workbook()
        self._validate_worksheet(sheet)
        row, col = self._validate_cellref(cellref, row, col)
        col = self._validate_column(col)
        row = self._validate_row(row)

        if isinstance(data, tuple):
            data = str(data).lstrip("(").rstrip(")")
        elif isinstance(data, list):
            data = str(data).lstrip("[").rstrip("]")
        elif isinstance(data, pd.Series):
            data = data.to_string(index=False).strip().replace("\n",",")

        data = self._cleanse_data(data)

        try:
            str(data)
        except:
            raise ValueError("Data could not be coerced to string - try supplying scalar, list, tuples or a series")

        sheet = self.wb.Sheets(sheet)

        if sheet.ProtectContents:
            raise InvalidSheetError(f"The sheet '{sheet.Name}' is protected, please unprotect before attempting to write to it")

        sheet.Cells(row, col).Value = data

    def write_row(self, data, sheet, cellref=None, startrow=None, startcol=None, endcol=None):
        """
        Write list-like data to a specific Range in a workbook.
        Data structures will be coerced into a series.
        startrow is a row number, startcol can be supplied as a string Excel Reference i.e. "A" or column index.
        endcol is optional - a warning will be returned if the series length does not match the range supplied.
        An Error should be returned if the object passed cannot be written.
            args:
                data : Variable to write to the range
                cellref : An Excel cell reference to write from
                startrow : A row number to write to
                startcol : A column name or index to write to
                endcol : A column name or index to truncate the data to
                sheet : The sheet name or index to write to
        """
        self._validate_workbook()
        self._validate_worksheet(sheet)
        startrow, startcol = self._validate_cellref(cellref, startrow, startcol)
        startcol = self._validate_column(startcol)

        if not isinstance(data, (pd.Series, list, tuple)):
            raise ValueError("Data supplied must be a list-like structure")

        data = self._cleanse_data(data)

        try:
            data_series = pd.Series(data)
        except ValueError as e:
            print(str(e)+"\nData could not be coerced to Series")
            raise Exception from e

        end = None
        if endcol is not None:
            end = (self._validate_column(endcol)) + (startcol-1)
            if end != len(data):
                warnings.warn("\nObject supplied differs in length to supplied range!\nExcess data will be truncated.", UserWarning)
                if end < len(data):
                    data = data[:end]
        else:
            end = len(data)+(startcol-1)

        sheet = self.wb.Sheets(sheet)

        if sheet.ProtectContents:
            raise InvalidSheetError(f"The sheet '{sheet.Name}' is protected, please unprotect before attempting to write to it")

        sheet.Range(sheet.Cells(startrow, startcol), sheet.Cells(startrow, end)).Value = data_series.values

    def delete_sheet(self, sheet):
        """
        Delete a worksheet
            args:
            sheet : The sheet to delete
        """
        self._validate_workbook()
        self._validate_worksheet(sheet)

        self.wb.Sheets(sheet).Delete()
        # Reset Sheets
        self.sheetnames = []
        for sheet in self.wb.Sheets:
            self.sheetnames.append(sheet.Name)

    def set_sheet_visibility(self, sheet, visibility):
        """
        Set a worksheets visibility.
            args:
            sheet : The sheet to change the visibility of
            visibility : The level of visibility to set. Provide either a string or int value 
                {"visible": -1, "hidden": 1, "very hidden": 2}
        """
        self._validate_workbook()
        self._validate_worksheet(sheet)

        visibilty_values = {"visible": -1, "hidden": 1, "very hidden": 2}

        if (isinstance(visibility, str) and visibility.lower() not in visibilty_values or
            isinstance(visibility, int) and visibility not in [visibilty_values.values()]):
            raise ValueError("Visibility value supplied is invalid, "
                             "please provide a valid string or int value. "
                             + str(visibilty_values))

        if visibility in visibilty_values:
            visibility = visibilty_values[visibility]

        self.wb.Sheets(sheet).Visible = visibility

    def protect_sheet(self, sheet, password=None, draw_objects=True, contents=True,
                      scenarios=True, allow_sort=False, allow_filter=False, enable_selection=True):
        """
        Protect a worksheet
        args:
            sheet : The name/index of the worksheet to protect
            password : A password to lock the sheets withb(Optional)
            draw_objects: Protect shapes (Optional: Default=True)
            contents: Protect contents (Optional: Default=True)
            scenarios: Protect scenatios (Optional: Default=True)
            allow_sort: Allow user to sort (Optional: Default=False)
            allow_filter: Allow user to filter (Optional: Default=False)
            enable_selection: Set to false to disable user selecting cells (Optional: Default=True)
        """
        self._validate_workbook()
        self._validate_worksheet(sheet)

        if self.wb.Sheets(sheet).ProtectContents:
            warnings.warn("\nSheet {} is already protected.".format(sheet), UserWarning)
        else:
            if not enable_selection:
                self.wb.Sheets(sheet).EnableSelection = c.xlNoSelection
                
            self.wb.Sheets(sheet).Protect(
                password, draw_objects, contents, scenarios,
                False, False, False, False, False, False, False, False, False, # Unimplemented positional args
                allow_filter, allow_sort, False
            )

    def unprotect_sheet(self, sheet, password=None):
        """
        Unprotect a worksheet
        args:
            sheet : The name/index of the worksheet to unprotect
            password : A password to unlock the sheets with (Optional)
        """
        self._validate_workbook()
        self._validate_worksheet(sheet)

        if not self.wb.Sheets(sheet).ProtectContents:
            warnings.warn("\nSheet {} is not protected.".format(sheet), UserWarning)
        else:
            self.wb.Sheets(sheet).Unprotect(password)

    def protect_workbook(self, password=None):
        """
        Protect the current workbook and it's structure
        args:
            password : A password to lock the workbook with (Optional)
        """
        self._validate_workbook()

        self.wb.Protect(password, True)

    def unprotect_workbook(self, password=None):
        """
        Unprotect the current workbook and it's structure
        args:
            password : A password to unlock the workbook with (Optional)
        """
        self._validate_workbook()

        self.wb.Unprotect(password)

    def get_sheetnames(self):
        """
        Return a list of worksheet names
        """
        self._validate_workbook()

        return self.sheetnames

    def refresh_all(self):
        """
        Refresh all workbook connections and pivot tables
        """
        self._validate_workbook()

        # Refreshes the DB connections
        self.wb.RefreshAll()
        
        # Refresh all pivots - should be lighter than a second full refresh
        for sheet in self.wb.WorkSheets:
            for pivot in sheet.PivotTables:
                pivot.RefreshTable()
                pivot.Update()

    def read_dataframe(self, sheet, header=True, excel_range=None, startrow=None, endrow=None, startcol=None, endcol=None):
        """
        Reads in a range of an Excel spreadsheet and attempts to return a pandas dataframe object.
        args:
            sheet: The sheet name/index to read from
            header: Does this range include column headers? Default == True
            range: An Excel range to read supplied as a string e.g. "A1:B5" -- Supply this instead of start and end row values
            startrow: The starting row to read from
            endrow: The final row to read from
            startcol: The starting column to read from
            endcol: The final column to read from
        """
        self._validate_workbook()
        self._validate_worksheet(sheet)
        excel_range = self._validate_range(excel_range, startrow, endrow, startcol, endcol)

        data = self.wb.Sheets(sheet).Range(excel_range).Value
        
        if hasattr(data, "__len__") and not isinstance(data, str):
            if header is True:
                data = pd.DataFrame(list(data[1:]), columns=data[0])
            else:
                data = pd.DataFrame(list(data))
        else:
            data = pd.DataFrame([data])

        return data

    def read_cell(self, sheet, cellref=None, col=None, row=None):
        """
        Reads in a range of an Excel spreadsheet and attempts to return a pandas dataframe object.
        args:
            sheet: The sheet name/index to read from
            cellref : An Excel cell reference to write from
            row : A row number to read from
            col : A column name or index to read from
        """
        self._validate_workbook()
        self._validate_worksheet(sheet)
        row, col = self._validate_cellref(cellref, row, col)
        col = self._validate_column(col)
        row = self._validate_row(row)
        
        sheet = self.wb.Sheets(sheet)

        data = sheet.Range(sheet.Cells(row, col), sheet.Cells(row, col)).Value

        return data

    def conditional_formatting(self, sheet, condition="cell_value", excel_range=None, startrow=None, endrow=None, startcol=None, endcol=None, **kwargs):
        """
        Add conditional formatting to a range. Current implementation means the formatting cannot be removed or modified once added.
        This is OK as it is not intended for use in interactive sessions at the moment.
        args:
            sheet: The sheet to add formatting to.
            range: The range to format
            startcol, endcol, startrow, endrow: Integers representing the co-ordinates of a range
            condition: Specifies whether the conditional format is based on a cell value or an expression. See https://docs.microsoft.com/en-us/office/vba/api/excel.xlformatconditiontype
            logic: The conditional format operator
            value: The value or formula associated with the condition i.e. Cells in range less than "0.99"
            value2: Only used when a 'between' logical operator is passed, i.e. Cells in range between 0.80 and "0.99"
            **kwargs: snake_case formatting arguments i.e. font_colour, font_size, merge, bold, interior_colour ... (see self.get_format_args() for valid values)
        """
        arg_list = [i for keys in self.format_args.values() for i in keys]
        for k in kwargs.keys():
            if k not in arg_list:
                warnings.warn("Unknown parameter {!r} ignored.".format(k))
                del kwargs[k]
        for i in [keys for keys in self.format_args["Condition"].keys()]: # Add required values
            if i not in kwargs.keys():
                try:
                    kwargs[i] = ast.literal_eval(i)
                except NameError:
                    kwargs[i] = self.null_arg
        self._validate_workbook()
        self._validate_worksheet(sheet)
        excel_range = self._validate_range(excel_range, startrow, endrow, startcol, endcol)

        condition_dict = {"above_average": c.xlAboveAverageCondition, "is_blank": c.xlBlanksCondition, "cell_value": c.xlCellValue,
                          "color_scale": c.xlColorScale, "data_bar": 4, "is_error": c.xlErrorsCondition, "expression": c.xlExpression,
                          "icon_set": 6, "no_blanks": c.xlNoBlanksCondition, "no_errors": c.xlNoErrorsCondition, "text": c.xlTextString,
                          "time_period": c.xlTimePeriod, "top_ten": c.xlTop10, "unique": c.xlUniqueValues}

        logic_dict = {"between": c.xlBetween, "equal_to": c.xlEqual, "greater_than": c.xlGreater, "greater_equal": c.xlGreaterEqual, "less_than": c.xlLess,
                      "less_equal": c.xlLessEqual, "not_between": c.xlNotBetween, "not_equal": c.xlNotEqual, self.null_arg: self.null_arg}

        try:
            logic_dict[kwargs["logic"]]
        except KeyError:
            raise ValueError("Invalid 'logic' value supplied.")

        try:
            condition_dict[condition]
        except KeyError:
            raise ValueError("Invalid 'condition' value supplied.")

        wb_range = self.wb.Sheets(sheet).Range(excel_range)
 
        _format = wb_range.FormatConditions.Add(Type=condition_dict[condition], Operator=logic_dict[kwargs["logic"]], Formula1=kwargs["value"], Formula2=kwargs["value2"])

        # Apply the actual formatting
        for key in kwargs.keys():
            if key in self.format_args["Format"]:
                if "border_" in key:
                    for k in kwargs[key].keys():
                        exec("_format.{}".format(self.format_args["Format"][key][k]))
                else:
                    exec("_format.{}".format(self.format_args["Format"][key]))

    def format_range(self, sheet, excel_range=None, startrow=None, endrow=None, startcol=None, endcol=None, **kwargs):
        """
        Add formatting to a range.
        args:
            sheet: The sheet to add formatting to.
            excel_range: The range to format
            startcol, endcol, startrow, endrow: Integers representing the co-ordinates of a range
            **kwargs: snake_case'd formatting arguments i.e. font_colour, font_size, merge, bold, interior_colour ... (see self.get_format_args for valid values)
        """
        arg_list = [keys for keys in self.format_args["Format"]]
        for k in kwargs.keys():
            if k not in arg_list:
                warnings.warn("Unknown parameter {!r} ignored.".format(k))
                del kwargs[k]
        self._validate_workbook()
        self._validate_worksheet(sheet)
        excel_range = self._validate_range(excel_range, startrow, endrow, startcol, endcol)
        wb_range = self.wb.Sheets(sheet).Range(excel_range)

        for key in kwargs.keys():
            if key in self.format_args["Format"]:
                if "border_" in key:
                    for k in kwargs[key].keys():
                        exec("wb_range.{}".format(self.format_args["Format"][key][k]))
                else:
                    exec("wb_range.{}".format(self.format_args["Format"][key]))

    def reset_cursor(self, sheet=1, cell="A1"):
        """
        Move the cursor to cell A1 on all sheets - useful to use before saving
        This is added to make reports open nicely
        args:
            sheet : The sheet to open on after saving
            cell  : The cell to select on each sheet on open
        """
        self._validate_workbook()
        self._validate_worksheet(sheet)

        self.excel.EnableEvents = True

        for ws in self.wb.Sheets:
            if ws.Visible == -1 and ws.EnableSelection == c.xlNoRestrictions:
                ws.Activate()
                ws.Range(cell).Select()
                ws.Range(cell).Activate()
                self.excel.ActiveWindow.ScrollRow = 1
                self.excel.ActiveWindow.ScrollColumn = 1

        self.wb.Sheets(sheet).Activate()
        self.excel.EnableEvents = False


    def save_xlsx(self, outfile):
        """
        Save the active workbook to a new .xlsx file.
        args:
            outfile : A valid Windows path to a file to export to.
        """
        self._validate_workbook()

        outfile = os.path.normpath(Path(outfile))
        self.wb.SaveAs(outfile, ConflictResolution=c.xlLocalSessionChanges)
        self.wb.Saved = True

    def save_pdf(self, outfile, sheet=None):
        """
        Save the active workbook to a new .pdf file.
        You can specify a particular worksheet or list of worksheets to export.
        args:
            outfile : A valid Windows path to a file to export to, must have .pdf extension.
            sheet : specify a single worksheet to export by name or index (Optional)
        """
        self._validate_workbook()
        if sheet is not None:
            self._validate_worksheet(sheet)

        outfile = os.path.normpath(Path(outfile))
        if not outfile.endswith(".pdf"):
            raise ValueError("outfile argument must have a .pdf file extension")

        if sheet is None:
            self.wb.ExportAsFixedFormat(0, outfile)
        else:
            self.wb.Sheets(sheet).ExportAsFixedFormat(0, outfile)

    def close(self, save_changes=False):
        """
        Close the current active workbook and the current Excel session after 1 second.
        args
            save_changes : Bool - If True save changes to the active workbook before closing (Defaults to False) (Optional)
        """
        sleep(1)
        if self._wb_open == 1:
            self.wb.Close(save_changes)
        self.excel.Quit()
        self.sheetnames = []
        self._wb_open = 0
        self.wb = None
