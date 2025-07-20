"""module: file_excel
input/output (reader/writer) functions 
    - For reading data from Excel to Python 
    - And writing data from Python to Excel

Classes
-------
- XW(book, sheet_name, ranges, attr_names) -> instance of class
    - abc
- Read_XW(book_name, sheet_name, rg_list) -> instance of class
    - For reading data from Excel ranges.
    
Functions
---------
- get_book(path) -> xw.Book
    - Returns `xw.Book` object from `path`
- get_sheet(book, sheet_name) -> xw.main.sheets
    - Returns `xw.main.sheets` object
- get_range(sheet, rg_name) -> xw.main.Range
    - Returns `xw.main.Range` object for the given range name in a worksheet.
- rg_to_dict(rg) -> dict
    - Returns `dict` with data read from `sheet.range(rg_name)`
- dict_to_excel(sheet, rg_name, d, keys, isRowHeader)
    - Writes `d[keys]` to `sheet.range(rg_name)`
- save_fig(fig, file_name)
    - Save plotly `fig` to `file_name`
"""
import datetime
from pathlib import Path
import xlwings as xw
import numpy as np

# CLASSES
class XW:
    """
    xlwings convenience class.

    This class provides a simplified interface for working with Excel files using the `xlwings` library.
    """
    
    def __init__(self, book:Path, sheet_name:str, ranges:list[str], attr_names:list[str]=None)->None:
        """
        Initializes a new instance of the `XW` class.

        Parameters
        ----------
        book : str or Path
            Path to the Excel file.
        sheet_name : str or xw.Sheet
            Name of the sheet where `ranges` are scoped.
        ranges : list[str]
            A list of range names.
        attr_names : list[str], optional
            A list of attribute names corresponding to `ranges`.

        Attributes
        ----------
        app : xw.app
            Excel application object. api: https://docs.xlwings.org/en/stable/api/app.html
        book : xw.Book
            Excel Workbook object. api: https://docs.xlwings.org/en/stable/api/book.html
        sheet : xw.Sheet
            Excel Worksheet object. api: https://docs.xlwings.org/en/stable/api/sheet.html
        ranges : dict[xw.Range]
            dict containing `xw.Range` objects. api: https://docs.xlwings.org/en/stable/api/range.html
        """
        if isinstance(sheet_name, xw.Sheet):
            self.sheet = sheet_name
            self.book = self.sheet.book
        else:
            self.book = get_book(book)
            self.sheet = get_sheet(self.book, sheet_name)
        self.app = self.book.app
        self.ranges = {}
        if attr_names is not None and len(ranges) == len(attr_names):
            for rg, attr_name in zip(ranges, attr_names):
                self.ranges[attr_name] = get_range(self.sheet, rg_name=rg, isValue=False)
        else:
            for rg in ranges:
                self.ranges[rg] = get_range(self.sheet, rg_name=rg, isValue=False)
                
    def to_dict(self)->dict:
        """abc."""
        d = self.__dict__
        return d
        
class Read_XW:
    
    def __init__(self, book_name:str or Path, sheet_name:str, rg_list:list[str], rg_args:dict=None, custom_rg_args:dict=None) -> None:
        """Returns an instance with attributes matching named ranges provided in `rg_list` on `sheet`.
        
        Parameters
        ----------
        book_name : str or Path
            Path to workbook.
        sheet_name : str
            Sheet name.
        rg_list : list[str]
            List of named ranges on `book.sheets(sheet)` to import. Ex: ["rg_name1", "rg_name2"]
        rg_args : dict
            Contains the keys `{isRowHeader, isUnit, isTrimNone, isLowerCaseKey}`.
        custom_rg_args : dict[dict]
            Contains keys corresponding to `<rg_list[i]>[rg_args]`. Ex: dict(rg_name1=rg_args1, rg_name2=rg_args2)
        
        Returns
        -------
        class instance:
            - *.book (xw.Book): Book object.
            - *.sheet (xw.main.Sheet): Sheet object.
            - *.rng (xw.main.Range): Range object.
            - *.dict (dict): values from `rng` object.
            
        Notes
        -----
        1. `rg_args` keys:
            - isRowHeader [True]: if named range has headers on row[0]
            - isUnit [True]: if named range has a unit entry.
            - isTrimNone [False]: if `None` keys and values are removed from return dict.
            - isLowerCaseKey [False]: if keys of return dict are forced to lower case.
        2. `custom_rg_args` example: `= dict(<rg_name>=dict(isRowHeader=True, isUnit=True))`
        """
        
        # Add `book, sheet, ranges` to instance
        self.book = get_book(book_name)
        self.sheet = get_sheet(self.book, sheet_name)
        for s in rg_list:
            setattr(self, s, get_range(self.sheet, s))

        # Pass values from `rg` object to `data` dict
        # Set default arguments
        if rg_args is None:
            rg_args = dict(isRowHeader=True, isUnit=True, isTrimNone=False, isLowerCaseKey=False)

        self.dict = {}
        for s in rg_list:
            # Get specific arguments for this range (if provided), else use default arguments
            s_args = custom_rg_args.get(s, {}) if custom_rg_args else {}
            args = {**rg_args, **s_args}

            d = rg_to_dict(getattr(self, s), **args)
            self.dict[s] = d

    def attr_to_dict(self, instance, attrs:list) -> dict:
        """Extracts specified `attrs` from an `instance` and returns them as a dict."""
        return {
            attr: getattr(instance, attr) for attr in attrs if hasattr(instance, attr)
        }
            
# FUNCTIONS
def get_caller_book(book:str or Path, return_active_sheet=False)->xw.Book:
    """Returns `xw.Book` object of caller workbook."""
    try:
        book = xw.Book.caller()   # handle if called from Excel
    except Exception:
        xw.Book.set_mock_caller(str(book))
        book = xw.Book(book)   # handle from debugger or command line
    if book is None:
        raise ValueError('No workbook found.')
    return (book, book.sheets.active) if return_active_sheet else book

def get_book(path:str or Path, return_path=False) -> xw.Book:
    """Returns `xw.Book` object.
    
    Parameters
    ----------
    path : str | Path | xw.Book
        Path to Excel workbook.
    return_path : bool, optional
        If True, returns tuple (`xw.Book`, `Path`), else returns `xw.Book`.
    """
    if isinstance(path, str):
        book = xw.Book(Path(path))
    elif isinstance(path, Path):
        book = xw.Book(path)
    elif isinstance(path, xw.Book):
        book = path
    else:
        raise ValueError("Input must be `str` or `pathlib.Path` to a valid workbook object.")
    return (book, Path(book.fullname)) if return_path else book
    
def get_sheet(book:xw.Book, sheet_name:str) -> xw.main.sheets:
    """Returns `xw.main.sheets` object."""
    return book.sheets(sheet_name)

def get_range(sheet:xw.main.sheets, rg_name:str, isValue=False) -> xw.main.Range:
    """Returns `xw.main.Range` object (or values if isValue=True)."""
    try:
        rg = sheet.range(rg_name).options(ndim=2)
    except Exception as e:
        try:
            rg = sheet.book.range(rg_name).options(ndim=2)
        except:
            raise TypeError(f'Error getting Range: `{rg_name}`! Range likely does not exist in `sheet`') from e
    return rg.value if isValue else rg

# reader functions
def _rg_to_2d(rg):
    """Returns 2d range array."""
    return [rg] if isinstance(rg, (list, tuple)) and not isinstance(rg[0], (list, tuple)) else rg
    
def _rg_to_val(rg, is2d=True):
    """Extract values from range."""
    if isinstance(rg, xw.main.Range):
        rg = rg.value
    elif not isinstance(rg, (list, tuple)):
        raise ValueError("`rg` must be of type xw.main.Range or list.")
    if rg is None:
        raise ValueError("`rg` argument is not a valid range.")
    return _rg_to_2d(rg) if is2d else rg

def rg_to_dict(xw_rg, isRowHeader=True, isUnit=True, isTrimNone=False, isLowerCaseKey=False) -> dict:
    """Read Excel range to dict.

    Parameters
    ----------
    xw_rg : xw.main.Range or (list, tuple)
        Range object where values are stored.
    isRowHeader : bool, optional
        If `True`, then keys are stored on `row[0]`, else `col[0]`.
    isUnit : bool, optional
        If `True`, then header keys contain `dict(unit=<unit>, value=<value>)` child keys.
    isTrimNone : bool, optional
        If `True`, then trims `None` entries from header and values.
    isLowerCaseKey : bool, optional
        If `True`, then converts all header keys into lower case strings.

    Returns
    -------
    d : dict
        Data from `xw.main.Range` object.
    
    Notes
    -----
    `rg` object must refer to an Excel range with a specific format where:
    - `row[0]` or `col[0]`:     "prop" key
    - `row[1]` or `col[1]`:     "unit" key (if isUnit=True)
    - `row[2:]` or `col[2:]`:   "value" key
    
    Return dict of format: `d = dict(<prop> = dict(unit=<unit>, value=<value>))`.
    - "prop" key is the string in `row[0]` or `col[0]`
    - "unit" key is the string in `row[1]` or `col[1]`
    - "value" key holds the values in `row[2:]` or `col[2:]` stored in a list.
    - If "unit" key is empty, then returns dict of format `d = dict(<prop>=<value>)`      
    """
    
    # extract values from `xw_rg`
    rg = _rg_to_val(xw_rg, is2d=True)
    
    # Trim `rg` to exclude rows (or cols) with all None values or where `header` is None
    def is_empty(value):
        """Check if a value is considered empty for the purposes of trimming."""
        return value is None or value == "" or value == "null"

    def remove_last_element(rg, row_or_col:str):
        """Remove last `row_or_col` from `rg`."""
        if row_or_col == "row":
            rg.pop(-1)
        else:
            for row in rg:
                row.pop(-1)
        return rg
    
    # get value idx of `rg`
    v_idx = 2 if isUnit else 1
    
    if isTrimNone:
        min_index = v_idx - 1

        if isRowHeader:
            while all(is_empty(v) for v in rg[-1]):  # Check last row
                rg = rg[:-1]  # Remove last row

            while len(rg) > 1 and is_empty(rg[1][-1]) and (len(rg) <= 2 or is_empty(rg[v_idx][-1])):
                # Check if the unit key and the value key are also empty
                rg = remove_last_element(rg, "col")
        else:
            while all(is_empty(row[-1]) for row in rg):  # Check last column
                rg = remove_last_element(rg, "col")  # Remove last col

            while len(rg) > 1 and is_empty(rg[-1][1]) and (len(rg) <= 2 or is_empty(rg[-1][v_idx])):
                # Check if the unit key and the value key are also empty
                rg = remove_last_element(rg, "row")  # Remove last row

    # Construct return dict
    def make_dict_key(d:dict, keys:list[str], value):
        """Set value in a nested dict based on the list of `keys`."""
        for key in keys[:-1]:
            d = d.setdefault(key, {})
        d[keys[-1]] = value
    
    def make_val_key(value, delim=","):
        """Returns formatted value."""
        def _convert_str_to_float(s:str):
            """Convert `s` to float, if possible, otherwise return `s`."""
            try:
                return float(s)
            except ValueError:
                return s
            
        if len(value) == 1:  # If value is a single element list
            v = value[0]
            if isinstance(v, str) and delim in v:  # Check if value is a CSV-string
                value = [_convert_str_to_float(v.strip()) if isinstance(v, str) else v for v in v.split(delim)]
            elif isinstance(v, datetime.datetime):
                value = v.strftime('%Y-%m-%d %H:%M:%S')
            else:
                value = v   # extract single value from list
        return value
        
    d = {}
    if isRowHeader:
        keys = rg[0]
        units = rg[1] if isUnit else None
        values = np.array(rg[v_idx:], dtype=object).T.tolist()
    else:
        keys = [row[0] for row in rg]
        units = [row[1] for row in rg] if isUnit else None          
        values = [row[v_idx:] for row in rg]

    for key, value in zip(keys, values):
        if key is None:  # Skip if key is None
            continue

        if isUnit:
            unit = units[keys.index(key)]

        value = make_val_key(value)
        if isinstance(key, float):
            return None

        # Remove leading/trailing whitespaces from values
        d_keys = key.lower().split('.') if isLowerCaseKey else key.split('.')
        d_value = dict(unit=unit, value=value) if isUnit and unit is not None else value
        make_dict_key(d, d_keys, d_value)

    return d

def range_to_dict(sheet:xw.Book.sheets, rg_name:str, isRowHeader=True, isUnit=True, isTrimNone=False, isLowerCaseKey=False) -> dict:
    """Refer to `rg_to_dict()` docstring."""
    rg = get_range(sheet, rg_name)
    return rg_to_dict(rg, isRowHeader, isUnit, isTrimNone, isLowerCaseKey)

# writer functions
def dict_to_excel(rg:xw.main.Range, d:dict, keys:list[str], isRowHeader=True, isUnit=True):
    """Writes values from a dictionary `d` to an Excel range `rg`.
    
    The function loops through the `keys` list and writes the corresponding value from `d` 
    to the specified Excel range (`rg`), in the order defined by `keys`.
    
    Parameters
    ----------
    rg : xw.main.Range
        Excel range to which data will be written.
    d : dict
        Dictionary containing data to be written.
    keys : list[str]
        List of keys specifying the order in which the data will be written.
    isRowHeader : bool, optional
        If True, keys are treated as row headers, otherwise as column headers
    isUnit : bool, optional
        If True, expects each dictionary entry to have "unit" and "value" keys.
    """
    
    v_idx = 2 if isUnit else 1

    # Clearing existing contents
    range_to_clear = rg[v_idx:, :] if isRowHeader else rg[:, v_idx:]
    range_to_clear.clear_contents()

    for i, key in enumerate(keys):
        unit, val = d[key]["unit"], d[key]["value"]

        # Write unit (if applicable)
        if isUnit:
            unit_cell = rg[1, i] if isRowHeader else rg[i, 1]
            unit_cell.options(transpose=isRowHeader).value = unit

        # Write value
        val_cell = rg[v_idx, i] if isRowHeader else rg[i, v_idx]
        if isinstance(val, (int, float)):
            val_cell.options(transpose=isRowHeader).value = val
        else:
            val_range = rg[v_idx:, i] if isRowHeader else rg[i, v_idx:]
            val_range.options(transpose=isRowHeader).value = val

# PLOTLY
# def save_fig(fig, file_name:str or Path):
#     """Write Plotly figure to `file_name`."""
#     pio.kaleido.scope.default_format = "png"
#     fig.write_image(fig, file_name)
    
# SCRIPT
