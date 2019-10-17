import pyqrcode 
from pyqrcode import QRCode 
import xlrd 
import pdfkit
import os
import shutil

workbookName = input("Enter Workbook Name: ")
loc = ("./Sheets/"+workbookName+".xlsx") 

filename = input("Enter Sheetname/PDF name: ")
print("Opening Workbook, please wait")
# To open Workbook 
try:
    wb = xlrd.open_workbook(loc)
except:
    print("Workbook not found")
    exit()
try:
    userSheet = int(input("Enter Sheet no: ")) 
    sheet = wb.sheet_by_index(userSheet - 1) 
except:
    print("Sheet doesn't exist")
    exit()

names = []

path = "QRcodes/" + str(filename)

try:
    print("Creating Directory %s" % path)
    os.makedirs(path)
except OSError:
    pass
print("=========================")
print("Generating QR Codes")

try:
    for i in range(0,1001):
        vals = sheet.row_values(i)
        vals = vals[0:7]
        fileName = vals[0]
        names.append(fileName[4:])
        addString = ""
        for j in vals:
            if j == 0.0:
                j = 0
            addString += str(j)
            addString += ','
        addString = addString[:-1]
        # print(addString)
        # Comment from here
        url = pyqrcode.create(addString) 
        # url.svg("./QRcodes/" + str(filename) + "/" + vals[0] + ".svg", scale = 2) 
        url.svg("./QRcodes/" + str(filename) + "/" + str(i) + ".svg", scale = 2) 
except IndexError:
    pass

print("Done")
print("=========================")
print("Generating Template")
allFilenames = names[::]
# Temp data
for i in range(len(names)):
    names[i] = 'QRcodes/' + filename + '/' + str(names[i]) + '.svg'

# Read in the file
with open('./lib/template.html', 'r') as file :
  filedata = file.read()

# Write the file with template
with open('./QRcodes/' + filename + '/' + filename + '.html', 'w') as file:
  file.write(filedata)

# Add divs
with open('./QRcodes/' + filename + '/' + filename + '.html', 'a') as file:
    
    finalString = ""
    for i in range(len(names)):
        repDiv = "<div id = 'im'> \r\n \
            <img style = 'float: left;' src='"+str(i)+".svg'> \r\n \
            <pre id = 'im' style='font-size: 28px'>  "+ allFilenames[i]+"</pre> \r\n \
            </div> \r\n"
        finalString += repDiv
    file.write(finalString)
  
# End HTML
with open('./QRcodes/' + filename + '/' + filename + '.html', 'a') as file:
  file.write("<body>\r\n<html>")

print("Done")
print("=========================")
print("Generating PDF")
# Generate PDF from HTML
try:
    path = "PDFs/"
    os.makedirs(path)
except OSError:
    pass

options = {
    'page-size': 'A4',
    'margin-top': '0.75in',
    'margin-right': '0.75in',
    'margin-bottom': '0.75in',
    'margin-left': '0.75in'
}

pdfkit.from_file('./QRcodes/'+filename+'/'+filename+'.html', './PDFs/'+filename+'.pdf', options = options)
print("=========================")


# Clean Up
print("Cleaning Up")

try:
    path = 'QRcodes'
    shutil.rmtree(path)
except OSError:
    print ("Deletion of the directory %s failed" % path)

print("Done")
print("=========================")
