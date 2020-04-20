import urllib.request
import os
import threading
import tkinter as tk
import json
import getpass
import atexit
import glob

from threading import Thread
from tkinter import ttk
from getpass import getuser
from tkinter import filedialog
from tkinter import messagebox
from tkinter import simpledialog

# Note:
# If you are having issues with excel "DDDEOpen <filename>'. The macro may not be available in
# this workbook or all macros may be disabled
#       Try: http://www.noelpulis.com/fix-edocs-cannot-run-macro-deeopen-when-opening-an-excel-file/

# Also: The trusted locations set must include or cover this location
#       C:\Users\{YOURUSERNAMEHERE}\AppData\Roaming\OpenText\DM\Temp\


FULL_SCREEN = False 
FLOATING_WINDOW = True # Main window floats above all windows

doc_type = ['template','work_instruction','checklist','form','other'] # base 0 index

# IMPORTANT, this is only default sample, 
# Edit only the contents of the .json file. 
# Run this script once, and it will generate the json file for you to edit. 
DEFAULT_EDOCS_DATA = [
        {'info': 'CHK35v18 Non-EGM', 'doc_num': 1750506,'type': doc_type[2]},
        {'info': 'CHK39v10 LTFO', 'doc_num': 1805674,'type': doc_type[2]},
        {'info': 'CHK44v3 OK to LTFO', 'doc_num': 1813905,'type': doc_type[2]},
        {'info': 'WI10 Database Procedure', 'doc_num': 1739471,'type': doc_type[1]},
        {'info': 'FM14 Central Contact', 'doc_num': 1792324,'type': doc_type[3]},
        {'info': 'FM55 TU Key Decision Register', 'doc_num': 1751672,'type': doc_type[3]},
        {'type': "blank"}, 
        {'info': 'TP18 Memo', 'doc_num': 1822369,'type': doc_type[0]},
        {'info': 'GSB Memo (Director++)', 'doc_num': 1393000,'type': doc_type[0]},
        {'type': "blank"},  
        {'info': 'TP06 EGM Game Approval', 'doc_num': 1914608,'type': doc_type[0]},
        {'info': 'TP07 EGM Hardware Approval', 'doc_num': 1914354,'type': doc_type[0]},
        {'info': 'TP27v40 Systems Approval', 'doc_num': 1848746,'type': doc_type[0]},
        {'type': "blank"},  
        {'info': 'GSB Leave Calendar', 'doc_num': 1418148, 'type': doc_type[4]},
]

VERSION = '1.0'

class Preferences:

    def __init__(self, fname):
        self.fname = fname
        self.preferences = dict()
        self.floating_window = True
        self.full_screen = False
        
        if not os.path.isfile(self.fname):
            print("no preference file read, generating default pref files...")
            self.preferences = { 'eDOCS_data' : DEFAULT_EDOCS_DATA,
                                 'geometry': '',
                                 'user': getuser(),
                                 'prefs_fname' : self.fname,
                                 'floating_window': self.floating_window, # Main window floats above all windows
                                 'full_screen': self.full_screen, } # For those with real estate to spare
            self.writePreferences()

        self.preferences = self.readPreferences(self.fname)

    def readPreferences(self, fname):
        preferences = dict()
        
        with open(fname, 'r') as f:
            preferences = json.load(f)

        self.eDOCS_data = preferences['eDOCS_data']
        self.geometry = preferences['geometry']
        self.user = preferences['user']
        self.floating_window = preferences['floating_window']
        self.full_screen = preferences['full_screen']
        self.prefs_fname = preferences['prefs_fname']        

    def writePreferences(self):
        with open(self.fname,'w') as f:
            json.dump(self.preferences, f, sort_keys=True, indent=4)

class eDocsLaunch(threading.Thread):

    def __init__(self, fname):
        self.start(fname) 

    def start(self, fname): 
        self.root = tk.Tk()        
        self.preferences = Preferences(fname)
        self.cleanUp()
        
        Thread(self.setupGUI()).start()         

    def refresh(self, event): 
        self.root.destroy()
        self.start(self.preferences.prefs_fname)

    def open_config(self, event): 
        my_title='Select eDocsLaunch Config File'
        tmp = filedialog.askopenfilename(initialdir='.', title=my_title,
                                         filetypes = (("JSON files","*.json"),("all files","*.*")))

        if tmp:
            self.preferences.prefs_fname = tmp
            self.refresh('opening new file')

    def edit_config(self, event): 
        os.startfile(self.preferences.prefs_fname)

    def prompt_edocs_num(self, event): 
        edocs_num = simpledialog.askinteger("Open File", "Enter eDocs Reference Number?",
                                 parent=self.root)
        if edocs_num:
            self.handleButtonPress(edocs_num)
                
    def cleanUp(self): 
        drf_filelist = [x for x in os.listdir('.') if x.endswith('.drf')]
        for file in drf_filelist: 
          os.remove(file) # warning deletes files
        
    def setupGUI(self):
        self.root.wm_title("eDocsLaunch: v" + VERSION)
        self.root.resizable(1,1)
        # bindings
        self.root.bind("<Control-r>", self.refresh) # refresh window/buttons
        self.root.bind("<Control-o>", self.open_config) # open .json config
        self.root.bind("<Control-e>", self.edit_config) # open .json config
        self.root.bind("<Control-d>", self.prompt_edocs_num) # prompt for eDOCs doc num
        if self.preferences.floating_window == True:
            self.root.wm_attributes('-topmost', True)
        
        if self.preferences.full_screen == True:
            self.root.wm_attributes('-fullscreen','true')

            
        tk.Label(self.root, justify=tk.CENTER, text=self.preferences.prefs_fname).pack(side=tk.TOP,
                                                                          fill=tk.BOTH, expand=True)
        mbutton = tk.Button(
                    self.root,
                    text = "Launch eDocs Document...",
                    fg = 'red',
                    command = lambda event='__button_click__': self.prompt_edocs_num(event))
        mbutton.pack(side=tk.TOP,fill=tk.BOTH, expand=True)

        for entry in self.preferences.eDOCS_data:
            if entry['type'] in doc_type:
                # make buttons
                button = tk.Button(
                    self.root,
                    text = entry['info'],
                    # demo color
                    fg = 'blue' if entry['type'] == 'other' else 'black', 
                    command = lambda doc_num=str(entry['doc_num']): self.handleButtonPress(doc_num))
                button.pack(side=tk.TOP,fill=tk.BOTH, expand=True)
            else:
                # make a divider
                tk.Label(self.root, justify=tk.CENTER, text="-----").pack(side=tk.TOP,
                                                                          fill=tk.BOTH, expand=True)

        # adjust window size/position
        ws = self.root.winfo_screenwidth()
        hs = self.root.winfo_screenheight()
        if self.preferences.geometry == '':
            # position right of the screen
            self.root.geometry('+%d+%d' % (ws-375, hs/2-300)) 
        else:
            # print(self.root.geometry())            
            self.root.geometry(self.preferences.geometry)
            
        self.root.mainloop()

    def handleButtonPress(self, doc_num):
        fname_out = 'edoc_file_' + str(doc_num) + '.drf'
        url = 'https://jdcolgrconprd01.justice.qld.gov.au:8443/reports/open-edocs-attachment.aspx?docref=' + str(doc_num) + '&docver=R'

        if os.path.isfile(fname_out) == False: 
            urllib.request.urlretrieve(url, fname_out)
            print("generating: edoc_file_" + str(doc_num) + ".drf file")
        
        os.startfile(fname_out)

def main():
    app = None
    prefs_flist = [x for x in os.listdir('.') if x.endswith('.json')]
    if len(prefs_flist) > 0:
        # open most recently edited file
        # min is for Windows, max maybe for linux?
        latest_file = min(prefs_flist, key=os.path.getctime) 

        # read only first file in list
        preferences = Preferences(latest_file) 
        app = eDocsLaunch(preferences.prefs_fname)
    else:
        # default pref filename
        app = eDocsLaunch('prefs.json') 

if __name__ == "__main__": main()
