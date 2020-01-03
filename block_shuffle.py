import re
import pandas as pd
import comtypes.client
from comtypes import COMError
from comtypes.client import CreateObject, GetActiveObject
from os.path import dirname, basename, splitext, join
from time import sleep 
import threading #TODO: is required?
import raa_logger
import logging

class acad_block:
    def __init__(self, orig_name, x_cord, y_cord):
        self.orig_name = orig_name
        self.x_cord = x_cord
        self.y_cord = y_cord
        self.uniq_name = ''
        self.new_name = ''

#TODO: all over the file - do we need to replace variables with self.<variable name>?
class block_shuffle:
    def __init__(self, event):
        self.logger = logging.getLogger("raa_logger")
        self.event = event

    def check_legal_mapping(self, land_use_codes, cellno_formats): #TODO: exception or error? what happens in case of exception?
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
            wstr = "BL:NAME C008000\nCELLNO C004000\nCODE N004000\nBL:X N012004\nBL:Y N012004\n"
            f.write(wstr)
        return template_filepath

    def get_new_name(self, formats_d, code, idx):
        format_max_cellnos = self.get_max_cellnos(formats_d[code]) # get max number of cellnos (according to the format)
        if idx > format_max_cellnos:
            self.logger.exception(f"\nError: Exceeded maximum number of possible cellno values allowed by format\nThere can be {format_max_cellnos} cellnos\nExiting...") #TODO: raise error to GUI, which will open window
            exit()
        new_name = str(int(formats_d[code])+idx)
        return new_name

    def get_nearest_block_idx(self, block, blocks, skip_blocks, code):
        self.logger.debug(f"For code {code} there are {len(blocks) - len(skip_blocks)} blocks left")
        self.logger.debug(f"Getting block nearest to {block.orig_name}")
        candid_blocks = list(b for b in blocks if b not in skip_blocks and b is not block)
        min_dist = 0
        if len(candid_blocks) > 1:
            for candid_block in candid_blocks:
                dist = (float(block.x_cord) - float(candid_block.x_cord))**2 + (float(block.y_cord) - float(candid_block.y_cord))**2
                self.logger.debug(f"Distance between {block.orig_name} and {candid_block.orig_name} is {dist}")
                if dist < min_dist or min_dist == 0:
                    min_dist = dist
                    nearest_block_idx = blocks.index(candid_block)
        else:
            nearest_block_idx = blocks.index(block)
        return nearest_block_idx


    def shuffle(self, acad_filepath, mapping_excel_filepath):
        # Open Autocad and dwg file
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
        
        # Create dictionary with:
        # keys: block codes
        # values: arrays of all acad_block classes of that code
        blocks_data = {}
        for line in file_str:
            self.logger.debug(f"\nline:{line}") 
            data_l = line.split(",")
            data_strip_l = [re.sub("[\s\t\n\']", "", x) for x in data_l]
            line_desc = data_strip_l[0].lower()
            block_name = data_strip_l[1]
            block_code = data_strip_l[2]
            x_cord = data_strip_l[3]
            y_cord = data_strip_l[4]
            if "cellno" in line_desc:
                self.logger.debug(f"block_code:{block_code}")
                self.logger.debug(f"block_name:{block_name}")
                # check valid name - only consisting digits
                if block_name.isdigit():
                    block = acad_block(block_name, x_cord, y_cord)
                    if block_code not in blocks_data:
                        blocks_data.update({block_code : []})
                    blocks_data[block_code].append(block)
                    self.logger.debug(f"length of blocks_data:{len(blocks_data)}")
                    #data_d.setdefault(block_code,  []).append(block_name)
                else:
                    self.logger.warning(f"Block name: {block_name} is an invalid name - should consist of digits only!\nIgnoring block and moving on to the next one")
            #TODO: OLD VERSION: before removing - verify that all checks and exceptions were taken to new version
#        # Create dictionary of Original {code : cellno} pairs
#        data_d = {}
#        new_data_d = {}
#        print_new_data_d = {}
#        for line in file_str:
#            self.logger.debug(f"line:{line}") 
#            data_l = line.split(",")
#            data_strip_l = [re.sub("[\s\t\n\']", "", x) for x in data_l]
#            if "cellno" in data_strip_l[0].lower():
#                self.logger.debug(f"data_strip_l[1]:{data_strip_l[1]}") 
#                if data_strip_l[1].isdigit():
#                    data_d.setdefault(data_strip_l[2],  []).append(data_strip_l[1])
#                else:
#                    self.logger.warning(f"Block name: {data_strip_l[1]} is an invalid name - should consist of digits only!\nIgnoring block and moving on to the next one")


        # Add unique temp name and a new name to each block
        # Naming should be done according to geografical distance between each group of block_code blocks,
        # such that the result should be that a group of same type blocks that are close to each other should have similar names, e.g. 200, 201, 202
        # To do this, we go through following steps:
        # 1. go over each code in the blocks_data dict
        uniq_char = 'A'
        for code, blocks in blocks_data.items():
            # 2. re-order blocks list by geographic proximity:
            # pick a random block and find its geografically nearest block (while ignoring blocks already picked),
            # then add picked block to skip list, rename the nearest with new and uniq names.
            # then move on to the nearest and look for the (now) closest to it.
            skip_blocks = []
            idx = 0
            cnt = 0
            if code not in formats_d.keys():
                self.logger.exception("\nError: Autodesk CODE was not found in excel list of codes\nExiting...") #TODO: raise error to GUI, which will open window
                exit() #TODO: needed? 
            while len(skip_blocks) < len(blocks):
                self.logger.debug(f"len of skip_blocks:{len(skip_blocks)}   len of blocks:{len(blocks)}")
                # rename with uniq and new names
                block = blocks[idx]
                block.uniq_name = uniq_char + str(cnt)
                block.new_name = self.get_new_name(formats_d, code, idx)
                self.logger.debug(f"Block {block.orig_name} - added unique temp name: {block.uniq_name}")
                self.logger.info(f"Block {block.orig_name} - added new name: {block.new_name}")
                idx = self.get_nearest_block_idx(block, blocks, skip_blocks, code)
                self.logger.debug(f"Nearest to block {block.orig_name} is {blocks[idx].orig_name}")
                cnt += 1
                skip_blocks.append(block)
            
            uniq_char = chr(ord(uniq_char) + 1) # advance uniq char for the next code
        #TODO: after assembling new and uniq names for all - need to actually replace

        #TODO: OLD VERSION: before removing - verify that all checks and exceptions were taken to new version
##        # Create dictionary of New {code : cellno} pairs
##        mid_char = 'A'
##        for i, key in enumerate(data_d):
##            mid_char = chr(ord(mid_char) + 1)
##            self.logger.debug(f"mid_char:{mid_char}")
##            if key not in formats_d.keys():
##                self.logger.exception("\nError: Autodesk CODE was not found in excel list of codes\nExiting...") #TODO: raise error to GUI, which will open window
##    #            exit() #TODO: needed? 
##            self.logger.debug(f"\nCODE {key}:\nOriginal CELLNO values:{data_d[key]}")
##            format_max_cellnos = self.get_max_cellnos(formats_d[key]) # get max number of cellnos (according to the format)
##            for j, val in enumerate(data_d[key]):
##                if j > format_max_cellnos:
##                    self.logger.exception(f"\nError: Exceeded maximum number of possible cellno values allowed by format\nThere can be {format_max_cellnos} cellnos\nExiting...") #TODO: raise error to GUI, which will open window
##                    exit()
##                mid_val = mid_char + str(j)
##                new_val = str(int(formats_d[key])+j)
##                new_data_d.update({mid_val : new_val})
##                print_new_data_d.setdefault(key, []).append(new_val) # only for debug print
##                # replace cellno in autocad - first by a unique temporary value
##                self.logger.info(f"Replacing {val} by unique mid value {mid_val}")
##                self.acad_replace_cellno(old_cellno=val,  new_cellno=mid_val)
##                sleep(0.1)
##            self.logger.info(f"Updated CELLNO values:{print_new_data_d[key]}") 
##        #acad_replace_cellno(old_cellno=val,  new_cellno=new_val)
##

            # now replace block names: first by uniq temp name, then by new name
            for i in range (2):
                for block in blocks:
                    old_name   = block.orig_name if i == 0 else block.uniq_name
                    new_name = block.uniq_name if i == 0 else block.new_name
                    if i == 1:
                        self.logger.info(f"Replacing block {block.orig_name} by {block.new_name} (going through middle temporary name {block.uniq_name})")
                    self.acad_replace_cellno(old_cellno=old_name,  new_cellno=new_name)
                    sleep(0.1)
            
#TODO: OLD VERSION
##        # now replace cellno by new values
##        self.logger.debug(f"new_data_d:{new_data_d}")
##        for uniq in new_data_d:
##            new_val = new_data_d[uniq]
##            self.logger.info(f"Replacing unique mid value {uniq} by {new_val}")
##            self.acad_replace_cellno(old_cellno=uniq,  new_cellno=new_val)
##            sleep(0.1)
        
        #TODO: at the end - save file
        self.event.set() #TODO: what if fails?
        return True

