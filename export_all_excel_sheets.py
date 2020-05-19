#!/usr/bin/env python3

# Export all of the sheets of an Excel file as separate CSV files.
# This is Python 3.

import sys
import os, signal
# sudo pip3 install psutil
import psutil
import subprocess
import time
import string

# sudo pip3 install unotools
# sudo apt install libreoffice-calc, python3-uno
from unotools import Socket, connect
from unotools.component.calc import Calc
from unotools.unohelper import convert_path_to_url
from unotools import ConnectionError

# Make sure libreoffice is installed.
soffice_exe = "/usr/lib/libreoffice/program/soffice.bin"
if (not os.path.isfile(soffice_exe)):
    print("ERROR: It looks like libreoffice is not installed. Aborting")
    sys.exit(101)

# Connection information for LibreOffice.
HOST = "127.0.0.1"
PORT = 2002

verbose = False

###################################################################################################
def is_excel_file(maldoc):
    """
    Check to see if the given file is an Excel file..

    @param name (str) The name of the file to check.

    @return (bool) True if the file is an Excel file, False if not.
    """
    typ = subprocess.check_output(["file", maldoc])
    if verbose:
        print("CHECK FILE TYPE: " + str(maldoc), file=sys.stderr)
        print(typ, file=sys.stderr)
        
    if (b"Excel" in typ):
        return True
    typ = subprocess.check_output(["exiftool", maldoc])
    if verbose:
        print(typ, file=sys.stderr)
    return ((b"ms-excel" in typ) or (b"Worksheets" in typ))

###################################################################################################
def wait_for_uno_api():
    """
    Sleeps until the libreoffice UNO api is available by the headless libreoffice process. Takes
    a bit to spin up even after the OS reports the process as running. Tries several times before giving
    up and throwing an Exception.
    """

    tries = 0

    while tries < 10:
        try:
            connect(Socket(HOST, PORT))
            return
        except ConnectionError:
            time.sleep(5)
            tries += 1

    raise Exception("libreoffice UNO API failed to start")

###################################################################################################
def get_office_proc():
    """
    Returns the process info for the headless libreoffice process. None if it's not running

    @return (psutil.Process)
    """

    for proc in psutil.process_iter():
        try:
            pinfo = proc.as_dict(attrs=['pid', 'name', 'username'])
        except psutil.NoSuchProcess:
            pass
        else:
            if (pinfo["name"].startswith("soffice")):
                if verbose:
                    print("SOFFICE RUNNING: " + str(pinfo), file=sys.stderr)
                return pinfo
    if verbose:
        print("SOFFICE NOT RUNNING", file=sys.stderr)
    return None

###################################################################################################
def is_office_running():
    """
    Check to see if the headless libreoffice process is running.

    @return (bool) True if running False otherwise
    """

    return True if get_office_proc() else False

###################################################################################################
def run_soffice():
    """
    Start the headless, UNO supporting, libreoffice process to access the API, if it is not already
    running.
    """

    # start the process
    if not is_office_running():

        # soffice is not running. Run it in listening mode.
        cmd = soffice_exe + " --headless --invisible " + \
              "--nocrashreport --nodefault --nofirststartwizard --nologo " + \
              "--norestore " + \
              '--accept="socket,host=' + HOST + ',port=' + str(PORT) + ',tcpNoDelay=1;urp;StarOffice.ComponentContext"'
        subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, shell=True)
        wait_for_uno_api()

def get_component(fname, context):
    """
    Load the object for the Excel spreadsheet.
    """
    url = convert_path_to_url(fname)
    component = Calc(context, url)
    return component

def convert_csv(fname):
    """
    Convert all of the sheets in a given Excel spreadsheet to CSV files.

    fname - The name of the file.
    return - A list of the names of the CSV sheet files.
    """

    # Make sure this is an Excel file.
    if (not is_excel_file(fname)):

        # Not Excel, so no sheets.
        if verbose:
            print("NOT EXCEL", file=sys.stderr)
        return []

    # Run soffice in listening mode if it is not already running.
    run_soffice()
    
    # TODO: Make sure soffice is running in listening mode.
    # 
    
    # Connect to the local LibreOffice server.
    context = connect(Socket(HOST, PORT))

    # Load the Excel sheet.
    component = get_component(fname, context)

    # Iterate on all the sheets in the spreadsheet.
    controller = component.getCurrentController()
    sheets = component.getSheets()
    enumeration = sheets.createEnumeration()
    r = []
    pos = 0
    if sheets.getCount() > 0:
        while enumeration.hasMoreElements():

            # Move to next sheet.
            sheet = enumeration.nextElement()
            name = sheet.getName()
            if (name.count(" ") > 10):
                name = name.replace(" ", "")
            if verbose:
                print("LOOKING AT SHEET " + str(name), file=sys.stderr)
            controller.setActiveSheet(sheet)

            # Set up the output URL.
            short_name = fname
            if (os.path.sep in short_name):
                short_name = short_name[short_name.rindex(os.path.sep) + 1:]
            outfilename =  "/tmp/sheet_%s-%s--%s.csv" % (short_name, str(pos), name.replace(' ', '_SPACE_'))
            outfilename = ''.join(filter(lambda x:x in string.printable, outfilename))

            pos += 1
            r.append(outfilename)
            url = convert_path_to_url(outfilename)

            # Export the CSV.
            component.store_to_url(url,'FilterName','Text - txt - csv (StarCalc)')
            if verbose:
                print("SAVED CSV to " + str(outfilename), file=sys.stderr)
            
    # Close the spreadsheet.
    component.close(True)

    # clean up
    os.kill(get_office_proc()["pid"], signal.SIGTERM)
    if verbose:
        print("KILLED SOFFICE", file=sys.stderr)
    
    # Done.
    if verbose:
        print("DONE. RETURN " + str(r), file=sys.stderr)
    return r

fname = sys.argv[1]
if ((len(sys.argv) > 1) and (sys.argv[1] == "-v")):
    verbose = True
    fname = sys.argv[2]
print(convert_csv(fname))
