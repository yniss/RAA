import re
import pandas as pd
import comtypes.client
from comtypes import COMError
from comtypes.client import CreateObject, GetActiveObject
from os.path import dirname, basename, splitext, join
from time import sleep #TODO: is really required?
import threading
import raa_logger
import logging

#TODO: all over the file - do we need to replace variables with self.<variable name>?
class block_shuffle:
    def __init__(self, event):
        self.logger = logging.getLogger("raa_logger")
        self.event = event

    def check_legal_mapping(self, land_use_codes, cellno_formats): #TODO: raise error to GUI, which will open window + exception to logger?
        error = True
        if len(land_use_codes) > len(set(land_use_codes)):
            self.logger.exception("\nError: \"land use code\" values must be unique")
        elif len(cellno_formats) > len(set(cellno_formats)):
            self.logger.exception("\nError: \"cellno format\" values must be unique")
        elif len(land_use_codes) != len(cellno_formats):
            self.logger.exception("\nError: number of \"land use code\" values should be equal to those of \"cellno format\"")
        else:
            error = False

        if error: #TODO: if self.logger.exception - then exit still needed? to check 
            self.logger.info("Exiting...")
            exit()

    def trailing(self, s):
        return len(s) - len(s.rstrip('0'))
    
    # according to the number of '0' digits in the "cellno format" - decide what is the max nummber of cellno's 
    def get_max_cellnos(self, cellno_format):
        trailing_zeros = self.trailing(cellno_format)
        #print(f"{cellno_format} has {trailing_zeros} '0's")
        return 10*trailing_zeros


    def open_acad(self, filepath):
        try: #Get AutoCAD running instance
            self.logger.info("\nChecking for an active AutoCAD app...\n")
            acad = GetActiveObject("AutoCAD.Application")
            state = True
        except(OSError,COMError): #If autocad isn't running, open it
            self.logger.info("No active app - opening AutoCAD...\n")
            acad = CreateObject("AutoCAD.Application",dynamic=True)
            state = False
        acad.Visible = False #TODO: 1. how to get invisible AutoCAD right at opening? 2. should make invisible if already opened?

        if state: #If you have only 1 opened drawing
            self.logger.info("Found an active app\n")
            self.doc = acad.Documents.Item(0)
        else:
            self.doc = acad.Documents.Open(filepath)
#        return doc

    def acad_command(self, command_str):
        for i in range (50):
            try:
                self.logger.debug(f'Sending command:{command_str}')
                self.doc.SendCommand(command_str) 
            except:
                self.failed = True
                self.logger.debug(f"\Did not succeed in sending AutoCAD command {i} times")
                sleep(0.1)
            else:
                self.failed = False
                break
        if self.failed:
            self.logger.exception(f"\nError: did not succeed in sending AutoCAD command {i} times")

    # Extract used cellno and codes from Autocad file 
    def acad_ext_cellno_codes(self, template_filepath, ext_filepath):
        self.logger.info("\nExtracting CELLNO data from Autocad...\n")
        # 1. Select all blocks
        self.acad_command('._select all  ') #Notice that the last SPACE is equivalent to hiting ENTER
        #You should separate the command's arguments also with SPACE
       
        # Suppress dialog box for following file read/write
        self.acad_command('._filedia 0 ')
       
        # 2. Attribute extraction (ATTEXT)
        self.acad_command('._-attext c ' + template_filepath + '\r' + ext_filepath + '\ry\r' )
       
        # Return file read/write dialog box
        self.acad_command('._filedia 1 ')

    def acad_replace_cellno(self, old_cellno, new_cellno):
    #    acad_command('._-attedit n\rn\rCellno\rCELLNO\r' + old_cellno + '\r' + old_cellno + '\r' + new_cellno + '\r')
        #self.acad_command('._-attedit n\rn\rCellno\rCELLNO\r\r' + old_cellno + '\r' + new_cellno + '\r')
        self.acad_command('._-attedit n\rn\rCellno\rCELLNO\r' + old_cellno + '\r' + old_cellno + '\r' + new_cellno + '\r')
    #    acad_command('._-attedit y\rCellno\rCELLNO\r' + old_cellno + '\rc\r' + x_pos + ',' + y_pos + '\r' + x_pos + ',' + y_pos + '\r' + new_cellno + '\r')
    #    x_corner_0 = str(float(x_pos)-1)
    #    y_corner_0 = str(float(y_pos)-1)
    #    x_corner_1 = str(float(x_pos)+1)
    #    y_corner_1 = str(float(y_pos)+1)
    #    print(f"x_corner_0:{x_corner_0}")
    #    print(f"y_corner_0:{y_corner_0}")
    #    acad_command('._-attedit y\r\r\r' + old_cellno + '\rc\r' + x_corner_0 + ',' + y_corner_0 + '\r' + x_corner_1 + ',' + y_corner_1 + '\r\rv\rr\r' + new_cellno + '\rn\r')

    def gen_template_file(self, acad_filepath):
        template_filepath = dirname(acad_filepath) + '/' + 'attr_extract_template.txt'
        self.logger.info(f"Extracting block names from .dwg file...\n")
        with  open(template_filepath, "w") as f: 
#            wstr = "BL:NAME C008000\nCELLNO N003000\nCODE N004000\nBL:X N012004\nBL:Y N012004\n"
            wstr = "BL:NAME C008000\nCELLNO C004000\nCODE N004000\nBL:X N012004\nBL:Y N012004\n"
            f.write(wstr)
        return template_filepath

    def shuffle(self, acad_filepath, mapping_excel_filepath):
        #logger = logging.getLogger("raa_logger")
        #global doc
        #TODO: how to switch to invisible?
        # Open Autocad and dwg file
        #doc = open_acad(acad_filepath)
        self.open_acad(acad_filepath)
        
        
        # Extract cellno and code data
        template_filepath = self.gen_template_file(acad_filepath)
        ext_filepath = dirname(acad_filepath) + '/' + splitext(basename(acad_filepath))[0] + '.txt'
        self.acad_ext_cellno_codes(template_filepath, ext_filepath)
        
        # Read mapping excel and create a formats dict
        mapping_sheet = 'mapping' #TODO: take mapping sheet name from user
        df = pd.read_excel(mapping_excel_filepath, mapping_sheet) 
        self.logger.debug(f"Reading excel file: {mapping_excel_filepath}\tsheet name: {mapping_sheet}\n{df}")
        land_use_codes = list(map(lambda x: str(int(x)), df['land use code'].dropna().tolist()))
        cellno_formats = list(map(lambda x: str(int(x)), df['cellno format'].dropna().tolist()))
        self.logger.debug(f"len(land_use_codes):{land_use_codes}")
        self.logger.debug(f"len(cellno_formats):{cellno_formats}")
        self.check_legal_mapping(land_use_codes, cellno_formats)
        
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
            self.logger.debug(f"line:{line}") 
            data_l = line.split(",")
            data_strip_l = [re.sub("[\s\t\n\']", "", x) for x in data_l]
            if "cellno" in data_strip_l[0].lower():
                self.logger.debug(f"data_strip_l[1]:{data_strip_l[1]}") 
                if data_strip_l[1].isdigit():
                    data_d.setdefault(data_strip_l[2],  []).append(data_strip_l[1])
                else:
                    self.logger.warning(f"Block name: {data_strip_l[1]} is an invalid name - should consist of digits only!\nIgnoring block and moving on to the next one")
        
        
        # Create dictionary of New {code : cellno} pairs
        mid_char = 'A'
        for i, key in enumerate(data_d):
            mid_char = chr(ord(mid_char) + 1)
            self.logger.debug(f"mid_char:{mid_char}")
            if key not in formats_d.keys():
                self.logger.exception("\nError: Autodesk CODE was not found in excel list of codes\nExiting...") #TODO: raise error to GUI, which will open window
    #            exit() #TODO: needed? 
            self.logger.debug(f"\nCODE {key}:\nOriginal CELLNO values:{data_d[key]}")
            format_max_cellnos = self.get_max_cellnos(formats_d[key]) # get max number of cellnos (according to the format)
            for j, val in enumerate(data_d[key]):
                if j > format_max_cellnos:
                    self.logger.exception(f"\nError: Exceeded maximum number of possible cellno values allowed by format\nThere can be {format_max_cellnos} cellnos\nExiting...") #TODO: raise error to GUI, which will open window
                    exit()
                mid_val = mid_char + str(j)
                new_val = str(int(formats_d[key])+j)
                new_data_d.update({mid_val : new_val})
                print_new_data_d.setdefault(key, []).append(new_val) # only for debug print
                # replace cellno in autocad - first by a unique temporary value
                self.logger.info(f"Replacing {val} by unique mid value {mid_val}")
                self.acad_replace_cellno(old_cellno=val,  new_cellno=mid_val)
                sleep(0.1)
            self.logger.info(f"Updated CELLNO values:{print_new_data_d[key]}") 
        #acad_replace_cellno(old_cellno=val,  new_cellno=new_val)
        
        # now replace cellno by new values
        self.logger.debug(f"new_data_d:{new_data_d}")
        for uniq in new_data_d:
            new_val = new_data_d[uniq]
            self.logger.info(f"Replacing unique mid value {uniq} by {new_val}")
            self.acad_replace_cellno(old_cellno=uniq,  new_cellno=new_val)
            sleep(0.1)
        
        #TODO: at the end - save file
        self.event.set() #TODO: what if fails?
        return True

