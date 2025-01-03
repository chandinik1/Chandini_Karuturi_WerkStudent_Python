# **WerkStudent_Python**

## Overview
This repository contains the solution for the given task, including:
- requirements.txt : file listing the necessary packages.
- data_extraction.py : A python script to extarct the data from the sample_invoice pdfs
- data_extraction : An executable file to run the program 
- sample_invoice_1.pdf, sample_invoice_2.pdf - Input files for the script.

## Extracting Information from PDF Files
- The script uses a library called Fitz to read the content of PDF files.
- Each PDF file is scanned to find the required information, such as: Date, Amount.
- A list is used to store the extracted information, with the file name and the extracted data (date and amount) as the values.

## Creating an Excel File
- The extracted data is written into an Excel file with the following:
  - Sheet 1 Contains a table with three columns - FileName, Date, Amount.
-Sheet 2 (Pivot table)
- A Pivot Table is created to summarize the data from Sheet 1, added as a second sheet in the same Excel file.
- Can be summarized using sum function , by default its mean.

## Creating a CSV File
- The data from the Excel file is saved as a CSV file, with "," as a separator.

## Creating an executable from pyinstaller
- the excetuable runs as a script and creates the output file i.e, xlsx and csv file here.