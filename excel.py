"""@package excel.py
Partial implementation of xlrd.book object interface.
"""

from __future__ import print_function

import os
import random
import json
import subprocess
import string

import filetype

####################################################################
def read_sheet_from_csv(filename):
    """
    Read an Excel CSV file into a Sheet object.

    @param filename (str) The name of the CSV file.

    @return (ExcelSheet object) The Excel sheet object containing the CSV data.
    """

    # Open the CSV file.
    f = None
    try:
        f = open(filename, 'rb')
    except Exception as e:
        print("ERROR: Cannot open CSV file. " + str(e))
        return None

    # Read in all the cells. Note that this only works for a single sheet.
    row = 0
    r = {}
    for line in f:

        # Escape ',' in cell values so the split works correctly.
        line = line.strip()
        in_str = False
        tmp = ""
        for c in line:
            if (isinstance(c, int)):
                c = chr(c)
            if (c == '"'):
                in_str = not in_str
            if (in_str and (c == ',')):
                tmp += "#A_COMMA!!#"
            else:
                tmp += c
        line = tmp

        # Break out the individual cell values.
        cells = line.split(",")
        col = 0
        for cell in cells:

            # Add back in escaped ','.
            cell = cell.replace("#A_COMMA!!#", ",")

            # Strip " from start and end of value.
            dat = str(cell)
            if (dat.startswith('"')):
                dat = dat[1:]
            if (dat.endswith('"')):
                dat = dat[:-1]
            r[(row, col)] = dat
            col += 1
        row += 1

    # Close file.
    f.close()

    # Make an object with a subset of the xlrd book methods.
    r = make_book(r)
    #print "EXCEL:\n"
    #print r
    return r

####################################################################
def load_excel_libreoffice(data):
    """
    Load the sheets from a given in-memory Excel file into a Workbook object.

    @param data (binary blob) The contents of an Excel file.

    @return (ExcelBook object) On success return a workbook object with the read in
    Excel workbook, on failure return None.
    """
    
    # Don't try this if it is not an Office file.
    if (not filetype.is_office_file(data, True)):
        print("WARNING: The file is not an Office file. Not extracting sheets with LibreOffice.")
        return None
    
    # Save the Excel data to a temporary file.
    out_dir = "/tmp/tmp_excel_file_" + str(random.randrange(0, 10000000000))
    f = open(out_dir, 'wb')
    f.write(data)
    f.close()
    
    # Dump all the sheets as CSV files using soffice.
    output = None
    _thismodule_dir = os.path.normpath(os.path.abspath(os.path.dirname(__file__)))
    try:
        output = subprocess.check_output(["python3", _thismodule_dir + "/export_all_excel_sheets.py", out_dir])
    except Exception as e:
        print("ERROR: Running export_all_excel_sheets.py failed. " + str(e))
        os.remove(out_dir)
        return None

    # Get the names of the sheet files, if there are any.
    try:
        sheet_files = json.loads(output.replace(b"'", b'"'))
    except Exception as e:
        print(e)
        os.remove(out_dir)
        return None
    if (len(sheet_files) == 0):
        os.remove(out_dir)
        return None

    # Load the CSV files into Excel objects.
    sheet_map = {}
    for sheet_file in sheet_files:

        # Read the CSV file into a single Excel workbook object.
        tmp_workbook = read_sheet_from_csv(sheet_file)

        # Pull the cell data for the current sheet.
        cell_data = tmp_workbook.sheet_by_name("Sheet1").cells
        
        # Pull out the name of the current sheet.
        start = sheet_file.index("--") + 2
        end = sheet_file.rindex(".")
        sheet_name = sheet_file[start : end]

        # Pull out the index of the current sheet.
        start = sheet_file.index("-") + 1
        end = sheet_file[start:].index("-") + start
        sheet_index = int(sheet_file[start : end])
        
        # Make a sheet with the current name and data.
        tmp_sheet = ExcelSheet(cell_data, sheet_name)

        # Map the sheet to its index.
        sheet_map[sheet_index] = tmp_sheet

    # Save the sheets in the proper order into a workbook.
    result_book = ExcelBook(None)
    for index in range(0, len(sheet_map)):
        result_book.sheets.append(sheet_map[index])

    # Delete the temp files with the CSV sheet data.
    for sheet_file in sheet_files:
        os.remove(sheet_file)

    # Delete the temporary Excel file.
    if os.path.isfile(out_dir):
        os.remove(out_dir)
        
    # Return the workbook.
    return result_book

####################################################################
def read_excel_sheets(fname):
    """
    Read all the sheets of a given Excel file as CSV and return them as a ExcelBook object. 
    Returns None on error.

    @param fname (str) The name of the Excel file.

    @return (ExcelBook object) On success return a workbook object with the read in
    Excel workbook, on failure return None.
    """

    # Read the sheets.
    #try:
    f = open(fname, 'rb')
    data = f.read()
    f.close()
    return load_excel_libreoffice(data)
    #except Exception as e:
    #    print(e)
    #    return None

####################################################################
class ExcelSheet(object):
    """
    Single Excel sheet.
    """
    
    def __init__(self, cells, name="Sheet1"):
        self.cells = cells
        self.name = name

    def __repr__(self):
        r = ""
        r += "Sheet: " + self.name + "\n\n"
        for cell in self.cells.keys():
            if (len(self.cells[cell]) == 0):
                continue
            cell_content = None
            try:
                cell_content = "\t'" + str(self.cells[cell]) + "'"
                r += str(cell) + "\t=" + cell_content + "\n"
            except UnicodeDecodeError:
                cell_content = "\t'" + ''.join(filter(lambda x:x in string.printable, self.cells[cell])) + "'"
                r += str(cell) + "\t=" + cell_content + "\n"
        return r
    
    def cell(self, row, col):
        if ((row, col) in self.cells):
            return self.cells[(row, col)]
        raise KeyError("Cell (" + str(row) + ", " + str(col) + ") not found.")

    def cell_value(self, row, col):
        return self.cell(row, col)

####################################################################    
class ExcelBook(object):
    """
    Excel workbook containing multiple ExcelSheet objects.
    """
    
    def __init__(self, cells=None, name="Sheet1"):

        # Create empty workbook to fill in later?
        self.sheets = []
        if (cells is None):
            return

        # Create single sheet workbook?
        self.sheets.append(ExcelSheet(cells, name))

    def __repr__(self):
        r = ""
        for sheet in self.sheets:
            r += str(sheet) + "\n"
        return r
        
    def sheet_names(self):
        r = []
        for sheet in self.sheets:
            r.append(sheet.name)
        return r

    def sheet_by_index(self, index):
        if (index < 0):
            raise ValueError("Sheet index " + str(index) + " is < 0")
        if (index >= len(self.sheets)):
            raise ValueError("Sheet index " + str(index) + " is > num sheets (" + str(len(self.sheets)) + ")")
        return self.sheets[index]

    def sheet_by_name(self, name):
        for sheet in self.sheets:
            if (sheet.name == name):
                return sheet
        raise ValueError("Sheet name '" + str(name) + "' not found.")

####################################################################
def make_book(cell_data):
    return ExcelBook(cell_data)
