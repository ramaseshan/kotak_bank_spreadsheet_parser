#! /usr/bin/env python
from Tkinter import Tk
from Tkinter import *
from tkMessageBox import *
from tkFileDialog import askopenfilename

import pyexcel as pe
import pyexcel.ext.xls
import sys, time, os, unicodedata
# To get the current users home directory
from os.path import expanduser

home = expanduser("~/txt_files/")
if not os.path.exists(home):
    os.makedirs(home)
def convert_excel_to_txt(filename):
    fileout = home + (filename.split("/")[-1]).split('.')[0]+".txt"
    print fileout
    #print "Reading file ",filename

    records = pe.get_array(file_name=filename)
    f = open(fileout,'w')
    #print "Starting to process data. Hold your breath"

    for count,rec in enumerate(records[1:]):
        rec[0] = "DATALIFE"
        rec[1] = "RPAY"
        rec[5] = "04182010000104"
        rec[4] = time.strftime("%d/%m/%Y")
        line = ""
        for value in rec:
            if value and type(value) is unicode:	
                value = unicodedata.normalize('NFKD', value).encode('ascii','ignore')
            if rec[6] % 2 == 0:
                rec[6] = int(rec[6])
	        # Cross check payment types with mahesh
            if rec[2] == "NEFT" or rec[2] == "IFT":
                line = line + str(value)+"~"
            else:
                showerror("Error","Your Payment Type is Wrong in column %d. Please correct it and run the script again."%(count+2))
			#print "Exiting Script"
                delete_content(f)
                f.close()
                root.quit()
                #sys.exit()
        f.write(line[:-1])
        f.write("\n")
    f.close()
    showinfo("Final Status","File converted. Please see this path %s"%(fileout))
    root.quit()
#print "Finished writing ",fileout


def quit():
    root.quit()

def file_selector():
    Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
    filename = askopenfilename(parent=root,initialdir=home,title='Please select a spreadsheet for converting')
		#,filetypes = (("Spreadsheet Files","*.xls;"),)) #*.xlsx;*.odt"),("HTML files", "*.html;*.htm")))
    if filename:
        convert_excel_to_txt(filename)
    #showinfo("Answer", filename)

root = Tk()
root.wm_title("Kodak Excel Parser")
frame = Frame(root,height=300,width=300)
Button(text='Select Spreadsheet File', command=file_selector).pack(fill=X)
Button(text='Quit', command=quit).pack(fill=X)
root.mainloop()

