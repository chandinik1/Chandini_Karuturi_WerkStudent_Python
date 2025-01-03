# **WerkStudent_Python**

## Overview
This repository contains the solution for the given task, including:
- An executable file to run the program.
- A `requirements.txt` file listing the necessary packages.

## Extracting Information from PDF Files**
- The script uses a library called Fitz to read the content of PDF files.
- Each PDF file is scanned to find the required information, such as: Date, Amount.
- A list is used to store the extracted information, with the file name and the extracted data (date and amount) as the values.

## Creating an Excel File**
- The extracted data is written into an Excel file with the following:
  - Sheet 1 Contains a table with three columns - FileName, Date, Amount.

## Generating a Pivot Table**
- A Pivot Table is created to summarize the data from Sheet 1, added as a second sheet in the same Excel file.
- Can be summarized using sum function , by default its mean.

## Creating a CSV File**
- The data from the Excel file is saved as a CSV file, with "," as a separator.


