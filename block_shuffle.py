import re
import pandas as pd
import comtypes.client
from comtypes import COMError
from comtypes.client import CreateObject, GetActiveObject
from os.path import dirname, basename, splitext, join
from time import sleep #TODO: is really required?
#import my_log_queue

def check_legal_mapping(land_use_codes, cellno_formats): #TODO: raise error to GUI, which will open window
    error = True
    if len(land_use_codes) > len(set(land_use_codes)):
        print("\nError: \"land use code\" values must be unique")
    elif len(cellno_formats) > len(set(cellno_formats)):
        print("\nError: \"cellno format\" values must be unique")
    elif len(land_use_codes) != len(cellno_formats):
        print("\nError: number of \"land use code\" values should be equal to those of \"cellno format\"")
    else:
        error = False

    if error:
        print("Exiting...")
        exit()

def trailing(s):
    return len(s) - len(s.rstrip('0'))

# according to the number of '0' digits in the "cellno format" - decide what is the max nummber of cellno's 
def get_max_cellnos(cellno_format):
    trailing_zeros = trailing(cellno_format)
    #print(f"{cellno_format} has {trailing_zeros} '0's")
    return 10*trailing_zeros


def open_acad(filepath):
    global doc
    try: #Get AutoCAD running instance
        print("\nChecking for an active AutoCAD app...\n")
        acad = GetActiveObject("AutoCAD.Application")
        state = True
    except(OSError,COMError): #If autocad isn't running, open it
        print("No active app - opening AutoCAD...\n")
        acad = CreateObject("AutoCAD.Application",dynamic=True)
        state = False
    acad.Visible = False #TODO: 1. how to get invisible AutoCAD right at opening? 2. should make invisible if already opened?

    if state: #If you have only 1 opened drawing
        print("Found an active app\n")
        doc = acad.Documents.Item(0)
    else:
        doc = acad.Documents.Open(filepath)
    return doc

def acad_command(command_str): #TODO: add command status checking, and passing errors back to GUI, maybe try few times before quitting with failure
    global doc
    print(f'[debug] Sending command:{command_str}')
    doc.SendCommand(command_str)

# Extract used cellno and codes from Autocad file 
def acad_ext_cellno_codes(template_filepath, ext_filepath):
    print("\nExtracting CELLNO data from Autocad...\n")
    # 1. Select all blocks
    acad_command('._select all  ') #Notice that the last SPACE is equivalent to hiting ENTER
    #You should separate the command's arguments also with SPACE
   
    # Suppress dialog box for following file read/write
    acad_command('._filedia 0 ')
   
    # 2. Attribute extraction (ATTEXT)
    acad_command('._-attext c ' + template_filepath + '\r' + ext_filepath + '\ry\r' )
   
    # Return file read/write dialog box
    acad_command('._filedia 1 ')

def acad_replace_cellno(old_cellno, new_cellno):
#    acad_command('._-attedit n\rn\rCellno\rCELLNO\r' + old_cellno + '\r' + old_cellno + '\r' + new_cellno + '\r')
    acad_command('._-attedit n\rn\rCellno\rCELLNO\r\r' + old_cellno + '\r' + new_cellno + '\r')
#    acad_command('._-attedit y\rCellno\rCELLNO\r' + old_cellno + '\rc\r' + x_pos + ',' + y_pos + '\r' + x_pos + ',' + y_pos + '\r' + new_cellno + '\r')
#    x_corner_0 = str(float(x_pos)-1)
#    y_corner_0 = str(float(y_pos)-1)
#    x_corner_1 = str(float(x_pos)+1)
#    y_corner_1 = str(float(y_pos)+1)
#    print(f"x_corner_0:{x_corner_0}")
#    print(f"y_corner_0:{y_corner_0}")
#    acad_command('._-attedit y\r\r\r' + old_cellno + '\rc\r' + x_corner_0 + ',' + y_corner_0 + '\r' + x_corner_1 + ',' + y_corner_1 + '\r\rv\rr\r' + new_cellno + '\rn\r')

def gen_template_file(acad_filepath):
    template_filepath = dirname(acad_filepath) + '/' + 'attr_extract_template.txt'
    with  open(template_filepath, "w") as f: 
        wstr = "BL:NAME C008000\nCELLNO N003000\nCODE N004000\nBL:X N012004\nBL:Y N012004\n"
        f.write(wstr)
    return template_filepath

def shuffle(acad_filepath, mapping_excel_filepath): 
    global doc
    #TODO: how to switch to invisible?
    # Open Autocad and dwg file
    doc = open_acad(acad_filepath)
    
    # Extract cellno and code data
    template_filepath = gen_template_file(acad_filepath)
    ext_filepath = dirname(acad_filepath) + '/' + splitext(basename(acad_filepath))[0] + '.txt'
    acad_ext_cellno_codes(template_filepath, ext_filepath)
    
    # Read mapping excel and create a formats dict
    mapping_sheet = 'mapping' #TODO: take mapping sheet name from user
    df = pd.read_excel(mapping_excel_filepath, mapping_sheet) 
    print(f"Reading excel file: {mapping_excel_filepath}\tsheet name: {mapping_sheet}\n{df}")
    #print(f"YYY - my_log_queue.logger:{my_log_queue.logger}")
    #my_log_queue.logger.debug('Clock started')
    land_use_codes = list(map(lambda x: str(int(x)), df['land use code'].dropna().tolist()))
    cellno_formats = list(map(lambda x: str(int(x)), df['cellno format'].dropna().tolist()))
    print(f"len(land_use_codes):{land_use_codes}")
    print(f"len(cellno_formats):{cellno_formats}")
    check_legal_mapping(land_use_codes, cellno_formats)
    
    formats_d = {}
    for i, code in enumerate(land_use_codes):
        formats_d.update({code : cellno_formats[i]})
    
    
    # Read original Cellno <-> Use Code pairs
    with  open(ext_filepath) as f: 
        file_str = f.readlines()
    
    # Create dictionary of Original {code : cellno} pairs
    data_d = {}
    new_data_d = {}
    print_new_data_d = {}
    for line in file_str:
        print(f"[debug] line:{line}") # debug #TODO: create debug_print
        data_l = line.split(",")
        data_strip_l = [re.sub("[\s\t\n]", "", x) for x in data_l]
        if "cellno" in data_strip_l[0].lower():
            data_d.setdefault(data_strip_l[2],  []).append(data_strip_l[1])
    
    
    # Create dictionary of New {code : cellno} pairs
    mid_char = 'A'
    for i, key in enumerate(data_d):
        mid_char = chr(ord(mid_char) + 1)
        print(f"mid_char:{mid_char}")
        if key not in formats_d.keys():
            print("\nError: Autodesk CODE was not found in excel list of codes\nExiting...") #TODO: raise error to GUI, which will open window
            exit()
        print(f"\nCODE {key}:\nOriginal CELLNO values:{data_d[key]}")
        format_max_cellnos = get_max_cellnos(formats_d[key]) # get max number of cellnos (according to the format)
        for j, val in enumerate(data_d[key]):
            if j > format_max_cellnos:
                print(f"\nError: Exceeded maximum number of possible cellno values allowed by format\nThere can be {format_max_cellnos} cellnos\nExiting...") #TODO: raise error to GUI, which will open window
                exit()
            mid_val = mid_char + str(j)
            new_val = str(int(formats_d[key])+j)
            new_data_d.update({mid_val : new_val})
            print_new_data_d.setdefault(key, []).append(new_val) # only for debug print
            # replace cellno in autocad - first by a unique temporary value
            print(f"Replacing {val} by unique mid value {mid_val}")
            acad_replace_cellno(old_cellno=val,  new_cellno=mid_val)
            sleep(0.2)
        print(f"Updated CELLNO values:{print_new_data_d[key]}") 
    #acad_replace_cellno(old_cellno=val,  new_cellno=new_val)
    
    # now replace cellno by new values
    print(f"new_data_d:{new_data_d}")
    for uniq in new_data_d:
        new_val = new_data_d[uniq]
        print(f"Replacing unique mid value {uniq} by {new_val}")
        acad_replace_cellno(old_cellno=uniq,  new_cellno=new_val)
        sleep(0.2)
    
    #TODO: at the end - save file

