"""@package filetype
Check for Office file types
"""

from __future__ import print_function

# Office magic numbers.
magic_nums = {
    "office97" : "D0 CF 11 E0 A1 B1 1A E1",    # Office 97
    "office2007" : "50 4B 3 4",                # Office 2007+ (PKZip)
}

def get_1st_8_bytes(fname, is_data):

    info = None
    is_data = (is_data or (len(fname) > 200))
    if (not is_data):
        try:
            tmp = open(fname, 'rb')
            tmp.close()
        except:
            is_data = True
    if (not is_data):
        with open(fname, 'rb') as f:
            info = f.read(8)
    else:
        info = fname[:9]

    curr_magic = ""
    for b in info:
        byte_val = b
        try:
            byte_val = ord(b)
        except TypeError:
            pass
        curr_magic += hex(byte_val).replace("0x", "").upper() + " "
        
    return curr_magic
    
def is_office_file(fname, is_data):
    """
    Check to see if the given file is a MS Office file format.

    return - True if it is an Office file, False if not.
    """

    # Read the 1st 8 bytes of the file.
    curr_magic = get_1st_8_bytes(fname, is_data)

    # See if we have 1 of the known magic #s.
    for typ in magic_nums.keys():
        magic = magic_nums[typ]
        if (curr_magic.startswith(magic)):
            return True
    return False

def is_office97_file(fname, is_data):

    # Read the 1st 8 bytes of the file.
    curr_magic = get_1st_8_bytes(fname, is_data)

    # See if we have the Office97 magic #.
    return (curr_magic.startswith(magic_nums["office97"]))

def is_office2007_file(fname, is_data):

    # Read the 1st 8 bytes of the file.
    curr_magic = get_1st_8_bytes(fname, is_data)

    # See if we have the Office 2007 magic #.
    return (curr_magic.startswith(magic_nums["office2007"]))
