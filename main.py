# change file path name at line 67 between inverted commas " " eg: "D:/Work/1"
# change the \ to / in between the file names to not show errors
# at line 71 change the output folder name according to the month eg: January = 1

import xlwings as xw
import os #importing os library to handle file paths and directories

# function to process each excel file
def process_file(file_path):
    
    # opening the excel app without displaying it
    app = xw.App(visible=False)
    
    # opening the workbook from the specified file_path
    wb = app.books.open(file_path)
    
    # accessing the active sheet in the workbook
    sheet = wb.sheets.active

    # changing the font size to 10 and font weight to normal to print in one page 
    cell = sheet.range("A2:Z50")
    cell.api.Font.Size=10
    cell.api.Font.Bold=False

    # deleting columns P to T which contains cheque details
    sheet.range("P:T").delete()
    
    # changing row height of the 2nd row
    sheet.range("2:2").row_height = 13.20

    # changing the row height of the 14th row (where names start) to the row where the names end (last row - 2)
    last_row = sheet.range("C2").end("down").row
    start_row = 14
    sheet.range(f"{start_row}:{last_row - 2}").row_height = 45

    # changing the column width of each column accordingly
    sheet.range("G:G").column_width = 12
    sheet.range("I:I").column_width = 6.33
    sheet.range("J:J").column_width = 0
    sheet.range("L:L").column_width = 3.67
    sheet.range("M:M").column_width = 5
    sheet.range("N:N").column_width = 3.67
    sheet.range("O:O").column_width = 3.67
    sheet.range("P:P").column_width = 7.33
    sheet.range("Q:Q").column_width = 5.33
    sheet.range("R:R").column_width = 10.33
    sheet.range("S:T").column_width = 0


    #accessing the first shape (image) on the sheet
    image = sheet.shapes[0]
    #checking if the shape is an image and resizing it
    if image.type == "picture": #the image type should be in lowercase (picture)
        image.width = 126.72
        image.height = 54
    #changing the row height of the row containing the image    
    sheet.range("1:1").row_height = 55    

    # output path for the excel file and saving the file
    edited_file_path = os.path.join(output_folder, os.path.basename(file_path))
    wb.save(edited_file_path)

    # Closing the workbook and quiting the Excel application
    wb.close()
    app.quit()
    
    print(f"Processed {file_path}")

directory = "RD files"

# Define the output folder inside the same directory
output_folder = os.path.join(directory, "1")

# Create the output folder if it doesn't exist
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

for filename in os.listdir(directory):
    if filename.endswith(".xls"):
        file_path = os.path.join(directory, filename)
        process_file(file_path)

print("All files conversion complete")