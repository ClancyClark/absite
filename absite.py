# -*- coding: utf-8 -*-
"""
Created on Wed Mar 09 17:38:20 2016

@author: cjclark
"""

#absite pdf parse and store in excel file
#by clancy clark
# WFBH
# 3.9.2016

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

def processtext(filename):
    pdftext = convert(filename)
    #analyze text by line
    lines = pdftext.splitlines()

    #set defaults
    linenumber = 0
    PCscore = 0
    SSscore = 0
    Percscore = 0
    grades = []
    #loop through lines
    for line in lines:
        linenumber += 1  #track line numbers    
        match = re.search(r'(.+)Level:\s(\w+)',line)  #find name and resident level
        if match: 
            #grab resident name and year level        
            resident = match.group(1)
            resident = resident.strip()
            level = match.group(2)
    
        #grab percent correct
        PCline = re.search(r'Percent Correct',line)
        if PCline:
            PCscore = linenumber
        if linenumber == PCscore + 1:
            apsc = re.search(r'..',line)
            if apsc:
                appliedscience = apsc.group()
        if linenumber == PCscore + 2:
            cm = re.search(r'..',line)
            if cm:
                clinicalmgm = cm.group()
        if linenumber == PCscore + 3:
            mk = re.search(r'..',line)
            if mk:
                medknow = mk.group()
        if linenumber == PCscore + 4:
            pc = re.search(r'..',line)
            if pc:
                patientcare = pc.group()
        if linenumber == PCscore + 5:
            tt = re.search(r'..',line)
            if tt:
                totaltest = tt.group()
        #grab standard score
        SSline = re.search(r'Standard Score',line)
        if SSline:
            SSscore = linenumber
        if linenumber == SSscore + 1:
            apscs = re.search(r'.+',line)
            if apscs:
                appliedsciences = apscs.group()
        if linenumber == SSscore + 2:
            cms = re.search(r'.+',line)
            if cms:
                clinicalmgms = cms.group()
        if linenumber == SSscore + 3:
            mks = re.search(r'.+',line)
            if mks:
                medknows = mks.group()
        if linenumber == SSscore + 4:
            pcs = re.search(r'.+',line)
            if pcs:
                patientcares = pcs.group()
        if linenumber == SSscore + 5:
            tts = re.search(r'.+',line)
            if tts:
                totaltests = tts.group()
        #grab percentile
        Percline = re.search(r'Percentile',line)
        if Percline:
            Percscore = linenumber
        if linenumber == Percscore + 2:
            ptile = re.search(r'.+',line)
            if ptile:
                studentptile = ptile.group()
                if studentptile.isdigit():
                    studentptile = studentptile
                else: studentptile = 'Missing'
    #fill array
                
    grades = [resident, level, appliedscience, clinicalmgm, medknow, patientcare, totaltest,appliedsciences, clinicalmgms, medknows, patientcares, totaltests,studentptile]     
    return grades

    
#create excel file
workbook = xlsxwriter.Workbook('resident_scores.xlsx')
worksheet = workbook.add_worksheet('Grades')

row = 0

#add headers
worksheet.write(0,0,'Resident Name')
worksheet.write(0,1,'Resident Year')
worksheet.write(0,2,'Percent Correct: Applied Science')
worksheet.write(0,3,'Percent Correct: Clinical Mgm')
worksheet.write(0,4,'Percent Correct: Medical Knowledge')
worksheet.write(0,5,'Percent Correct: Patient Care')
worksheet.write(0,6,'Percent Correct: Total Test')
worksheet.write(0,7,'Standard Score: Applied Science')
worksheet.write(0,8,'Standard Score: Clinical Mgm')
worksheet.write(0,9,'Standard Score: Medical Knowledge')
worksheet.write(0,10,'Standard Score: Patient Care')
worksheet.write(0,11,'Standard Score: Total Test')
worksheet.write(0,12,'Percentile')

path = os.getcwd()
dirs = os.listdir(path)          
for gradefile in dirs:
  print 'working'
  if gradefile.endswith('.pdf'):
    grading = processtext(gradefile)
    print 'processing'
    row += 1      
    worksheet.write(row,0,grading[0])
    worksheet.write(row,1,grading[1])
    worksheet.write(row,2,grading[2])
    worksheet.write(row,3,grading[3])
    worksheet.write(row,4,grading[4])
    worksheet.write(row,5,grading[5])
    worksheet.write(row,6,grading[6])
    worksheet.write(row,7,grading[7])
    worksheet.write(row,8,grading[8])
    worksheet.write(row,9,grading[9])
    worksheet.write(row,10,grading[10])
    worksheet.write(row,11,grading[11])
    worksheet.write(row,12,grading[12])  
    workbook.close
    
print 'Complete'
