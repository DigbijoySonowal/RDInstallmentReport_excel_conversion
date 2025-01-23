# change file path name at line 67 between inverted commas " " eg: "D:/Work/1"
# change the \ to / in between the file names to not show errors
# at line 52 change the output folder name according to the month eg: January = 1

import xlwings as xw
import os #importing os library to handle file paths and directories
from PyPDF2 import PdfMerger

# function to convert excel to pdf
def convert_to_pdf(file_path):
    
    # opening the excel app without displaying it
    app = xw.App(visible=False)
    
    # opening the workbook from the specified file_path
    wb = app.books.open(file_path)
    

    # output path for the pdf
    pdf_file_name = os.path.basename(file_path).replace(".xls",".pdf")
    pdf_file_path = os.path.join(output_folder, pdf_file_name)
    
    # Debugging: Print the generated PDF file path
    print(f"Attempting to save PDF to: {pdf_file_path}")

    # saving the excel file as PDF
    wb.save()
    try:
        wb.to_pdf(pdf_file_path)
    except Exception as e:
        print(f"Error during PDF conversion: {e}")
        return
    
    # verifying if the PDF file was created and print the file path
    if os.path.exists(pdf_file_path):
        print(f"PDF created successfully: {pdf_file_path}")
        # adding pdf to the merger queue
        pdf_merger.append(pdf_file_path)
        print(f"Converted {file_path} to PDF and added to merger queue \n")
    else:
        print(f"Error: PDF {file_path} not added to merger queue \n")
        
    
    # Closing the workbook and quit the Excel application
    wb.close()
    app.quit()
    
    

# Directory containing the Excel files
directory = "RD files/1"

# Output folder where PDF file will be saved
output_folder = os.path.join(directory, "1")

# Creating the output folder if it doesn't exist
if not os.path.exists(output_folder):
    os.makedirs(output_folder)
    

# Initializing PDF merger    
pdf_merger = PdfMerger()

# Looping through each file in the directory
for filename in os.listdir(directory):
    if filename.endswith(".xls"):
        file_path = os.path.join(directory, filename)
        convert_to_pdf(file_path)
        
# Merging all PDFs into one file
merged_pdf_path = os.path.join(output_folder, "Merged_pdf.pdf")
pdf_merger.write(merged_pdf_path)
pdf_merger.close()

print("All files conversion complete")