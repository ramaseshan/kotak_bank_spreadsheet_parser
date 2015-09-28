"""
Code : Remove the dependency for Kodak Bank, default excel parser macros and just GNU/Linux to acheive it. 
Authors : Ramaseshan, Anandhamoorthy , Engineers, Fractalio Data  Pvt Ltd, Magadi, Karnataka.
Licence : GNU GPL v3.
Code Repo URL : https://github.com/ramaseshan/kodak_bank_excel_parser
"""
import pyexcel as pe
import pyexcel.ext.xls
import unicodedata
import sys
import time

def delete_content(pfile):
	pfile.seek(0)
	pfile.truncate()

filename = sys.argv[1]
fileout = filename.split('.')[0]+".txt"

print "Reading file ",filename

records = pe.get_array(file_name=filename)
f = open(fileout,'w')
print "Starting to process data. Hold your breath"

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
			print "Your Payment Type is Wrong in column %d. Please correct it and run the script again."%(count+2)
			print "Exiting Script"
			delete_content(f)
			f.close()
			sys.exit()
	f.write(line[:-1])
	f.write("\n")
f.close()
print "Finished writing ",fileout
