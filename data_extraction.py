import fitz
import re
import xlsxwriter
import pandas as pd

# Define file patterns and headers
files_data = {
    "sample_invoice_1.pdf": {
        "patterns": {
            "Amount": r"Gross Amount incl\. VAT\s*[\r\n]+(\d{1,3}(?:[.,]\d{2})?\s*)",
            "Date": r"Date\s*[\r\n]+\s*(\d+\.\s+\w+\s+\d{4})"
        }
    },
    "sample_invoice_2.pdf": {
        "patterns": {
            "Amount": r"Total[\s\n]*USD \$([\d,]+\.\d{2})",
            "Date": r"Invoice date:\s*(\w{3}\s\d{1,2},\s\d{4})"
        }
    }
}
headers = ['File Name', 'Amount', 'Date']
output_xlsx = 'data_extracted.xlsx'
output_csv = 'data_extracted.csv'

#Extracts data from the provided PDF files based on patterns.
def extract_data(files_data):
    data = []
    #Go through Dictionary containing file names and patterns to extract data.
    for file_name, data_patterns in files_data.items():
        try:
            doc = fitz.open(file_name)
            text = ""
            for page in doc:
                text += page.get_text()
            #text contains all the data from the PDF
            row = [file_name]
            for key, pattern in data_patterns["patterns"].items():
                match = re.search(pattern, text)
                row.append(match.group(1) if match else None)
            #List of extracted rows (file name, amount, date)
            data.append(row)
        except Exception as e:
            print(f"Error processing file {file_name}: {e}")
    return data


#Writes data to an Excel file and creates a pivot table.
def write_to_excel(data, headers, output_xlsx):
    workbook = xlsxwriter.Workbook(output_xlsx)
    worksheet = workbook.add_worksheet("Sheet1")
    #headers is list of column headers, put on row 0
    for col_num, header in enumerate(headers):
        worksheet.write(0, col_num, header)
    #data is list of extracted values, date
    for row_num, row_data in enumerate(data, start=1):
        for col_num, value in enumerate(row_data):
            worksheet.write(row_num, col_num, value)
    workbook.close()

    df = pd.DataFrame(data, columns=headers)
    #creating a pivot table - uisng Date as index and values As amount
    pivot_table = pd.pivot_table(df, 
                                 values='Amount', 
                                 index=['Date'], 
                                 columns=['File Name'], 
                                 aggfunc='sum', 
                                 fill_value=0)
    with pd.ExcelWriter(output_xlsx, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Sheet1', index=False)
        pivot_table.to_excel(writer, sheet_name='PivotTable', index=True)

#Converts the Excel file to a CSV format.
def write_to_csv(output_xlsx, output_csv):
    df = pd.read_excel(output_xlsx, engine='openpyxl')
    df.to_csv(output_csv, index=False, sep=',')


# Main Workflow
if __name__ == "__main__":
    extracted_data = extract_data(files_data)
    #we can access the individual values through the extracted_data list, using idices
    write_to_excel(extracted_data, headers, output_xlsx)
    write_to_csv(output_xlsx, output_csv)

