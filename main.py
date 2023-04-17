from pdf2docx import parse, Converter
import tkinter as tk
import sys
from tkinter import filedialog
from tkinter import ttk
import os

# global variables
pdf_file_path = ""
current_dir = os.path.dirname(__file__)
docx_file_path = ""

# functions


def finish_convert():
    label_result.config(text="Successful!\nFile in: " +
                        docx_file_path, fg="yellowgreen")


def open_file():
    global pdf_file_path
    pdf_file_path = filedialog.askopenfilename()

    label_file_path.config(text="File: " + pdf_file_path, fg="cyan")


def convert():
    global pdf_file_path
    global current_dir
    global docx_file_path

    try:
        file_extension = pdf_file_path[pdf_file_path.rindex(".")+1:]

        if file_extension == "pdf":
            label_result.config(text="Converting...", fg="yellow")

            docx_file_path = current_dir + "/result/" + pdf_file_path[pdf_file_path.rindex(
                "/")+1:pdf_file_path.rindex(".")] + ".docx"

            converter = Converter(pdf_file_path)
            converter.convert(docx_file_path)
            converter.close()

            label_result.after(5000, finish_convert)
        else:
            label_result.config(text="Error: only PDFs", fg="red")
    except Exception as error:
        print(error)


# *****GUI*****
root = tk.Tk()
root.title("PDF to DOCX Converter")

# frame
frame = tk.Frame(root, bd=2, bg="black")
frame.pack()

# menu bar
menu_bar = tk.Menu(root, bg="#e5e4e2")
file_sub_menu = tk.Menu(menu_bar, tearoff=0)
file_sub_menu.add_command(label="Open...", command=lambda: open_file())
file_sub_menu.add_separator()
file_sub_menu.add_command(label="Exit", command=lambda: root.quit())

menu_bar.add_cascade(label="File", menu=file_sub_menu)
root.config(menu=menu_bar)

# labels
label_title = tk.Label(
    frame, text="Wellcome to PDF to DOCX converter!", fg="yellowgreen", bg="black", font=("Arial", 18))
label_result = tk.Label(frame, text="", bg="black")
label_file_path = tk.Label(frame, text="", bg="black")

# buttons
button_convert = tk.Button(frame, text="Convert!",
                           command=lambda: convert(), bg="yellowgreen", borderwidth=2, highlightbackground="red", font=("Arial", 11, "bold"))

# grid
label_title.grid(row=1, columnspan=3)
label_file_path.grid(row=2, columnspan=3)
label_result.grid(row=3, columnspan=3)
button_convert.grid(row=4, column=2)

# display GUI
root.mainloop()

# exit
print("\nThanks for trying!")
sys.exit()
