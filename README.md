# absite

Files:  absite.py and wrong_answers.py
python 2.7.11
by clancy j clark
3/15/2016
http://www.clancyclark.com

Scrape 2016 ABSITE scoring PDFs to extract data and import into Excel file for analysis.

Two python programs were created in simplistic fashion to process data stuck in PDFs.  This work is in part an introduction to processing string text and a practical effort to extract data for purposes of surgery resident education.

PDFminer is used to grab text as long string.

regex expressions are used to grab text from string.

xlsxwriter is used to write to Excel.  

Run each program in a folder that contains PDF files.

Excel file will be created in the same folder.

Note:  This is not fancy but does the job.
