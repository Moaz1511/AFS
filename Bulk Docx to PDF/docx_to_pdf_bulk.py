import os
import tkinter as tk
from tkinter import filedialog, messagebox
import win32com.client

def convert_files(input_folder, output_folder):
    word = win32com.client.Dispatch('Word.Application')
    ppt = win32com.client.Dispatch('PowerPoint.Application')
    word.Visible = False
    # ppt.Visible = False  # REMOVE THIS LINE

    count_docx, count_pptx = 0, 0

    for filename in os.listdir(input_folder):
        in_path = os.path.normpath(os.path.join(input_folder, filename))
        if filename.lower().endswith('.docx'):
            out_path = os.path.normpath(os.path.join(output_folder, filename[:-5] + '.pdf'))
            try:
                print(f"Converting Word: {in_path} → {out_path}")
                doc = word.Documents.Open(in_path)
                doc.SaveAs(out_path, FileFormat=17)
                doc.Close()
                count_docx += 1
            except Exception as e:
                print(f"Error converting {filename}: {e}")
        elif filename.lower().endswith('.pptx'):
            out_path = os.path.normpath(os.path.join(output_folder, filename[:-5] + '.pdf'))
            try:
                print(f"Converting PPT: {in_path} → {out_path}")
                presentation = ppt.Presentations.Open(in_path, WithWindow=False)
                presentation.SaveAs(out_path, 32)  # 32 = PDF
                presentation.Close()
                count_pptx += 1
            except Exception as e:
                print(f"Error converting {filename}: {e}")
    word.Quit()
    ppt.Quit()
    return count_docx, count_pptx

def start_conversion():
    input_folder = input_folder_var.get()
    output_folder = output_folder_var.get()
    if not input_folder or not output_folder:
        messagebox.showerror("Error", "Please select both folders.")
        return
    msg.set("Converting, please wait...")
    root.update_idletasks()
    docx, pptx = convert_files(input_folder, output_folder)
    msg.set(f"Done! Converted {docx} DOCX and {pptx} PPTX files.")

def browse_input():
    folder = filedialog.askdirectory()
    if folder:
        input_folder_var.set(folder)

def browse_output():
    folder = filedialog.askdirectory()
    if folder:
        output_folder_var.set(folder)

root = tk.Tk()
root.title("Bulk DOCX & PPTX to PDF Converter (MS Office)")

input_folder_var = tk.StringVar()
output_folder_var = tk.StringVar()
msg = tk.StringVar()

tk.Label(root, text="Input Folder (contains .docx/.pptx):").grid(row=0, column=0, sticky="e")
tk.Entry(root, textvariable=input_folder_var, width=45).grid(row=0, column=1)
tk.Button(root, text="Browse", command=browse_input).grid(row=0, column=2)

tk.Label(root, text="Output Folder (PDFs will be saved here):").grid(row=1, column=0, sticky="e")
tk.Entry(root, textvariable=output_folder_var, width=45).grid(row=1, column=1)
tk.Button(root, text="Browse", command=browse_output).grid(row=1, column=2)

tk.Button(root, text="Convert All", command=start_conversion, bg="#2ecc71", fg="white", font=("Arial", 12, "bold")).grid(row=2, column=0, columnspan=3, pady=15)

tk.Label(root, textvariable=msg, fg="blue", font=("Arial", 11)).grid(row=3, column=0, columnspan=3)

root.resizable(False, False)
root.mainloop()
