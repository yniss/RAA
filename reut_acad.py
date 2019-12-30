from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkinter import Tk, Text, BOTH, W, N, E, S, messagebox, scrolledtext
from tkinter.ttk import Frame, Button, Label, Style
import re
import sys #TODO: is required for new logging?
from os.path import basename, splitext, isfile
import threading
import queue #TODO: required here?
import logging
#import concurrent.futures
import raa_logger
import block_shuffle as bshuf #TODO: temp removed
#import test_code #TODO: temp for logging debug



class Root(Tk):
    def __init__(self):
        super(Root, self).__init__()
        self.title("RAA") 
        self.wm_iconbitmap('C:/Users/reuts/Desktop/חן/קרית גת/arch.ico')
        self.acadFilename = ""
        self.excelFilename = ""
        self.event = threading.Event()

        lbl = Label(self, text="Title Place Holder", font=("Century Gothic Bold", 12), width=70, anchor='w')
        lbl.grid(column=0, row=1,  pady=10)

        # Label frame
        self.labelFrame1 = ttk.LabelFrame(self, text = "Block Name Shuffle:")
        self.labelFrame1.grid(row = 2, column = 0, padx = 20, pady = 20)
        self.labelFrame2 = ttk.LabelFrame(self, text = "Log:")
        self.labelFrame2.grid(row = 4, column = 0, padx = 20, pady = 20)

        # Autocad file: path text entry & file browse button
        self.acadLbl = ttk.Label(self.labelFrame1, text="Select AutoCAD File:")
        self.acadLbl.grid(row=1, column=0, sticky='E', padx=5, pady=2)
        self.acadEntry = ttk.Entry(self.labelFrame1, width=70)
        self.acadEntry.grid(row = 1, column = 1, pady = 3)
        self.acadButton = ttk.Button(self.labelFrame1, text = "Browse...", command = self.acadFileDialog)
        self.acadButton.grid(row = 1, column = 2, pady = 2)

        # Excel file: path text entry & file browse button
        self.excelLbl = ttk.Label(self.labelFrame1, text="Select Excel File:")
        self.excelLbl.grid(row=2, column=0, sticky='E', padx=5, pady=2)
        self.excelEntry = ttk.Entry(self.labelFrame1, width=70)
        self.excelEntry.grid(row = 2, column = 1, pady = 2)
        self.excelButton = ttk.Button(self.labelFrame1, text = "Browse...", command = self.excelFileDialog)
        self.excelButton.grid(row = 2, column = 2, pady = 2)

        # Run button
        self.excelButton = ttk.Button(self.labelFrame1, text = "Rename Blocks", command = self.runShuffle) 
        self.excelButton.grid(row = 3, column = 1, columnspan = 3, sticky="nesw",  pady = 2)

        # Logging to console
        self.console = raa_logger.ConsoleUi(self.labelFrame2, self.event)
        self.logger = logging.getLogger("raa_logger")
        self.logVar = BooleanVar()
        self.logButton = Checkbutton(self.labelFrame1, text="Debug Prints", variable=self.logVar, onvalue=True, offvalue=False)
        self.logButton.grid(row = 3, column = 0, sticky="nesw",  pady = 2)


    def acadFileDialog(self): 
        self.acadFilename = filedialog.askopenfilename(initialdir =  "/", title = "Select A File", filetype =
        (("AutoCAD files","*.dwg"),("all files","*.*")) )
        self.acadEntry.delete(0,END)
        self.acadEntry.insert(0,self.acadFilename)

    def excelFileDialog(self):
        self.excelFilename = filedialog.askopenfilename(initialdir =  "/", title = "Select A File", filetype =
        (("Excel files","*.xlsx"),("all files","*.*")) )
        self.excelEntry.delete(0,END)
        self.excelEntry.insert(0,self.excelFilename)

    def checkValidFile(self, filetype, filename, entrytext):
        # take filename from text
        filename = entrytext
        # check if no text at Entry or wrong file type, then shoot a warning message window
        fileext = ".dwg" if filetype == "AutoCAD" else ".xlsx"
        if filename == "" or splitext(basename(filename))[1] != fileext: 
            messagebox.showwarning(filetype + ' file not entered', filetype + ' file was not entered...\nPlease enter a valid ' + filetype + ' file')
            fileValid = False
        # check that file exist
        elif isfile(filename) == False:
            messagebox.showwarning('File does not exist', 'The file you entered does not exist...\nPlease enter a valid ' + filetype + ' file')
            fileValid = False
        else:
            fileValid = True
        return fileValid
    
        
    def runShuffle(self):
        acadFileValid = self.checkValidFile("AutoCAD", self.acadFilename, self.acadEntry.get()) 
        excelFileValid = self.checkValidFile("Excel", self.excelFilename, self.excelEntry.get())  
        
        if acadFileValid and excelFileValid:
            if self.logVar.get():
                logging.getLogger("raa_logger").setLevel(logging.DEBUG)
            my_bshuf = bshuf.block_shuffle(self.event)
            block_shuffle_thread = threading.Thread(target=my_bshuf.shuffle, args=(self.acadFilename, self.excelFilename), daemon=True) 
            try:
                block_shuffle_thread.start()
            except Exception as e:
                self.logger.exception("Error: could not run block shuffle thread")
            #else:
            #    return_value = block_shuffle_thread.join()
            #with concurrent.futures.ThreadPoolExecutor() as executor:
            #    block_shuffle_thread = executor.submit(my_bshuf.shuffle, (self.acadFilename, self.excelFilename))
            #    return_value = block_shuffle_thread.result()
            #if return_value:
            #    messagebox.showinfo('Job Successfully finished', 'Successfully renamed blocks')
            #else:
            #    messagebox.showwarning('Job Failed', 'Ohhh...\nCould not renamed blocks')
            #block_shuffle.shuffle(self.acadFilename, self.excelFilename)


#TODO: add status returned from shuffle and print success or failure by status

##  class TextRedirector(object): #TODO: update for logger use
##      def __init__(self, widget, tag="stdout"):
##          self.widget = widget
##          self.tag = tag
##  
##      def write(self, str):
##          self.widget.configure(state="normal")
##          self.widget.insert("end", str, (self.tag,))
##          self.widget.configure(state="disabled")


# ---------------------------------------------------------------------------------


root = Root()
root.mainloop()
