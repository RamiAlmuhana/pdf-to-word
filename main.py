import tkinter as tk
from tkinter import filedialog
from pdf2docx import Converter

def convert_pdf_to_docx(pdf_path, docx_path):
    cv = Converter(pdf_path)
    cv.convert(docx_path, start=0, end=None)
    cv.close()

def browse_pdf_file():
    pdf_file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    pdf_entry.delete(0, "end")
    pdf_entry.insert(0, pdf_file_path)

def convert_pdf():
    pdf_path = pdf_entry.get()
    if not pdf_path:
        status_label.config(text="Please select a PDF file first.")
        return

    docx_path = pdf_path.replace(".pdf", ".docx")
    convert_pdf_to_docx(pdf_path, docx_path)
    status_label.config(text=f"Conversion complete. The Word document is saved as {docx_path}")

# Create the main window
window = tk.Tk()
window.title("PDF to Word Converter")

# Create widgets
pdf_label = tk.Label(window, text="Select a PDF file:")
pdf_entry = tk.Entry(window)
pdf_browse_button = tk.Button(window, text="Browse", command=browse_pdf_file)
convert_button = tk.Button(window, text="Convert to Word", command=convert_pdf)
status_label = tk.Label(window, text="")

# Place widgets on the window
pdf_label.grid(row=0, column=0)
pdf_entry.grid(row=0, column=1)
pdf_browse_button.grid(row=0, column=2)
convert_button.grid(row=1, column=0, columnspan=3)
status_label.grid(row=2, column=0, columnspan=3)

window.mainloop()
