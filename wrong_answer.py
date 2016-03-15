# -*- coding: utf-8 -*-

#Created on Wed Mar 09 17:38:20 2016
#absite pdf parse and store in excel file wrong answers
#by clancy clark

import os
import re
import xlsxwriter
from cStringIO import StringIO
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage

#open and read pdf utility
def convert(fname, pages=None):
    if not pages:
        pagenums = set()
    else:
        pagenums = set(pages)

    output = StringIO()
    manager = PDFResourceManager()
    converter = TextConverter(manager, output, laparams=LAParams())
    interpreter = PDFPageInterpreter(manager, converter)

    infile = file(fname, 'rb')
    for page in PDFPage.get_pages(infile, pagenums):
        interpreter.process_page(page)
    infile.close()
    converter.close()
    text = output.getvalue()
    output.close
    return text 
            
path = os.getcwd()
dirs = os.listdir(path)   

#create excel file
workbook = xlsxwriter.Workbook('resident_answers.xlsx')
worksheet = workbook.add_worksheet('Wrong')

#add headers
worksheet.write(0,0,'Resident Name')
worksheet.write(0,1,'Resident Year')
worksheet.write(0,2,'Section1')
worksheet.write(0,3,'Section2')
worksheet.write(0,4,'Section3')
worksheet.write(0,5,'Section4')
worksheet.write(0,6,'Answer')

row = 1  
linenumber = 1

for gradefile in dirs:

  if gradefile.endswith('.pdf'):
    #add wrong answers   
    print 'working on', gradefile
    pdftext = convert(gradefile)
    #analyze text by line
    lines = pdftext.splitlines()
    
    #loop through lines
    for line in lines: 
      linenumber = linenumber + 1 #count lines
    
      match = re.search(r'(.+)Level:\s(\w+)',line)  #find name and resident level
      if match: 
        #grab resident name and year level        
        resident = match.group(1)
        resident = resident.strip()
        level = match.group(2)
    
      if linenumber > 8: #after header information check for answers        
        redoline = line.decode('ascii','replace')
          
        answerline = re.search(r'\b(?!SCORE\b)[A-Z]{3}.+',redoline) #find answers  
        if answerline:  
          section = answerline.group()
          sections = re.split('\s-\s',section) 

          row = row + 1
          #write excel cells
          worksheet.write(row,0,resident)
          worksheet.write(row,1,level) 
          worksheet.write(row,6,redoline)
          #write sections of answer
          column = 2          
          for phrase in sections:  
            worksheet.write(row,column,phrase)
            column = column + 1


  linenumber = 1 #reset line counting 
print 'Complete'
workbook.close()
