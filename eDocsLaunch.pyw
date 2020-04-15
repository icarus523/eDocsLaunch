import urllib.request
import os
import threading
import tkinter as tk

from threading import Thread
from tkinter import *
from tkinter import ttk

# Note:
# If you are having issues with excel "DDDEOpen <filename>'. The macro may not be available in
# this workbook or all macros may be disabled
#       Try: http://www.noelpulis.com/fix-edocs-cannot-run-macro-deeopen-when-opening-an-excel-file/

# Also: The trusted locations set must include or cover this location
#       C:\Users\{YOURUSERNAMEHERE}\AppData\Roaming\OpenText\DM\Temp\


FULL_SCREEN = False # For those with real estate to spare
FLOATING_WINDOW = True # Main window floats above all windows

doc_type = ['template','work_instruction','checklist','form','other'] # base 0 index

# Change this content to suit you, the order of this list, will be the order of the buttons.
# use {'type': "blank"} to add breaks
# Obviously if a file changes you have to confirm the document reference number
eDOCS_data = [
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
        {'info': 'CNP - Category 3 OEM Schedule', 'doc_num': 1845826,'type': doc_type[3]},
        {'info': 'The Star - Jackpot Transfer Failure Incident Log ', 'doc_num': 1852525,'type': doc_type[3]},
        {'type': "blank"},         
        {'info': 'File Review', 'doc_num': 1779838, 'type': doc_type[3]},
        {'info': 'GSB Leave Calendar', 'doc_num': 1418148, 'type': doc_type[3]},
        {'type': "blank"},
        {'info': 'Lotteries RNS', 'doc_num': 1832832, 'type': doc_type[4]},        
]

VERSION = '1.0'

class eDocsLaunch(threading.Thread):

    def __init__(self):
        self.root = tk.Tk()
        self.cleanUp() 
        
        Thread(self.setupGUI()).start() 

    def cleanUp(self): 
        drf_filelist = [x for x in os.listdir('.') if x.endswith('.drf')]
        
        for file in drf_filelist: 
           os.remove(file) # warning deletes files
        
    def setupGUI(self):
        self.root.wm_title("eDocsLaunch v" + VERSION)
        self.root.resizable(1,1)

        if FLOATING_WINDOW:
            self.root.wm_attributes('-topmost', True)
        
        if FULL_SCREEN:
            self.root.wm_attributes('-fullscreen','true') # full screen
            
        for entry in eDOCS_data:
            if entry['type'] in doc_type:
                # make buttons
                button = ttk.Button(
                    self.root,
                    text = entry['info'],
                    command = lambda doc_num=str(entry['doc_num']): self.handleButtonPress(doc_num))
                button.pack(side=TOP,fill=BOTH, expand=True)
            else:
                # make a divider
                tk.Label(self.root, justify=CENTER, text="-----").pack(side=TOP, fill=BOTH, expand=True)

        # adjust window size/position
        ws = self.root.winfo_screenwidth()
        hs = self.root.winfo_screenheight()
        # self.root.geometry('+%d+%d' % (ws-250, hs/2)) # position on screen
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
    app = eDocsLaunch()

if __name__ == "__main__": main()
