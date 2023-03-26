import os
import tkinter
from tkinter.filedialog import askopenfile, asksaveasfilename
from tkinter.messagebox import showinfo
import customtkinter
from docx2pdf import convert
import pdf2docx


customtkinter.set_appearance_mode('light')

root = customtkinter.CTk()
root.geometry('400x450')



def openfile():
    file = askopenfile(filetypes=[('Word Files', '*.docx')])
    # Ask the user to choose the destination folder and filename for the converted PDF file
    dest_path = asksaveasfilename(
        initialdir=os.path.expanduser("~"),
        defaultextension='.pdf',
        filetypes=[('PDF Files', '*.pdf')]
)
   # Convert the file and save it to the destination path
    convert(file.name, dest_path)
    showinfo("Done", "File successfully converted ")
def pdfToWord():
    # Ask the user to choose the PDF file to convert
    pdf_file_path = askopenfile(filetypes=[('PDF Files', '*.pdf')])

    if not pdf_file_path:
        return  # User cancelled the file dialog

    # Ask the user to choose the destination folder and filename for the converted Word file
    dest_path = asksaveasfilename(
        initialdir=os.path.expanduser("~"),
        defaultextension='.docx',
        filetypes=[('Word Files', '*.docx')]
    )

    # Convert the PDF file to Word
    pdf2docx.parse(pdf_file_path, dest_path)

    showinfo("Done", "File successfully converted to Word format.")

frame = customtkinter.CTkFrame(master=root)
frame.pack(pady=20, padx=60, fill='both', expand=True)

label = customtkinter.CTkLabel(master=frame, text="converter", font = (("Arial"), 22))
label.pack(pady = 12, padx = 10)

word_button = customtkinter.CTkButton(master=frame, text="Convert Word to PDF", command=openfile)
word_button.place(relx=0.5, rely=0.4, anchor=tkinter.CENTER)

pdf_button = customtkinter.CTkButton(master=frame, text="Convert PDF to Word", command=pdfToWord)
pdf_button.place(relx=0.5, rely=0.6, anchor=tkinter.CENTER)

root.mainloop()