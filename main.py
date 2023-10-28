import tkinter as tk
from tkinter import filedialog
import openpyxl
from PyPDF2 import PdfReader


# Function to open a file dialog and select a PDF file
def choose_pdf_file():
    pdf_file = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    pdf_file_entry.delete(0, tk.END)
    pdf_file_entry.insert(0, pdf_file)


# Function to process the selected PDF file and save data to an Excel file
def process_pdf_to_excel():
    pdf_file = pdf_file_entry.get()
    excel_file = excel_file_entry.get()

    # Automatically append ".xlsx" extension if not provided by the user
    if not excel_file.lower().endswith('.xlsx'):
        excel_file += '.xlsx'

    try:
        # Try to open and read the PDF file
        with open(pdf_file, 'rb') as file:
            content = PdfReader(file)
            fields = content.get_form_text_fields()
    except FileNotFoundError:
        # Handle the case where the PDF file is not found
        result_label.config(text="Error: The PDF file not found.")
        return

    try:
        # Try to open an existing Excel file and select the active worksheet
        workbook = openpyxl.load_workbook(excel_file)
        worksheet = workbook.active
    except FileNotFoundError:
        # Handle the case where the Excel file is not found, create a new file and select the active worksheet
        workbook = openpyxl.Workbook()
        worksheet = workbook.active

    for question, answer in fields.items():
        if answer:
            if answer.strip() and answer.strip().lower() != 'no':

                if answer.isnumeric():
                    answer = int(answer)  # Convert numeric text to int
                elif answer.replace(".", "", 1).isdigit():
                    answer = float(answer)  # Convert numeric text to float

                # Append question and answer to the Excel worksheet
                worksheet.append([question, answer])

    # Save the Excel file
    workbook.save(excel_file)
    result_label.config(text=f"Data from PDF file appended to {excel_file} successfully.")


unique_questions_texts = set()

# Create the GUI
root = tk.Tk()
root.title("PDF to Excel Converter")

# Label and input field to select a PDF file
tk.Label(root, text="Select PDF File:").pack()
pdf_file_entry = tk.Entry(root)
pdf_file_entry.pack()
tk.Button(root, text="Browse", command=choose_pdf_file).pack()

# Label and input field to enter the Excel file name
tk.Label(root, text="Enter Excel File Name:").pack()
excel_file_entry = tk.Entry(root)
excel_file_entry.pack()

# Button to initiate the conversion process
tk.Button(root, text="Convert", command=process_pdf_to_excel).pack()

# Label to display the result of the conversion
result_label = tk.Label(root, text="")
result_label.pack()

# Start the Tkinter main loop
root.mainloop()
