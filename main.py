import os
import tkinter
from tkinter.filedialog import askopenfile, asksaveasfilename
from tkinter.messagebox import showinfo
import customtkinter
from docx2pdf import convert


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
frame = customtkinter.CTkFrame(master=root)
frame.pack(pady=20, padx=60, fill='both', expand=True)

label = customtkinter.CTkLabel(master=frame, text="convert word to pdf", font = (("Arial"), 22))
label.pack(pady = 12, padx = 10)

button = customtkinter.CTkButton(master=frame, text="select", command = openfile)
button.place(relx=0.5, rely=0.5, anchor=tkinter.CENTER)


root.mainloop()