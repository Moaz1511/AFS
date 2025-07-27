import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from PyPDF2 import PdfReader, PdfWriter
from io import BytesIO
import os

PAGE_WIDTH = 612  # 8.5 inch
PAGE_HEIGHT = 766.8  # 10.65 inch

def create_overlay_with_image(image_path, position, y_value):
    packet = BytesIO()
    can = canvas.Canvas(packet, pagesize=(PAGE_WIDTH, PAGE_HEIGHT))

    img = ImageReader(image_path)
    img_width, img_height = img.getSize()
    scale = 0.2
    img_width *= scale
    img_height *= scale

    if position == 'right':
        x = PAGE_WIDTH - img_width
    else:
        x = 0
    y = y_value
    can.drawImage(img, x, y, width=img_width, height=img_height, mask='auto')
    can.save()
    packet.seek(0)
    return PdfReader(packet)

def stamp_pdf(input_pdf, output_pdf, left_image_path, right_image_path, y_value):
    reader = PdfReader(input_pdf)
    writer = PdfWriter()
    for i, page in enumerate(reader.pages):
        pos = 'right' if (i + 1) % 2 == 1 else 'left'
        image_path = right_image_path if pos == 'right' else left_image_path
        overlay = create_overlay_with_image(image_path, pos, y_value)
        page.merge_page(overlay.pages[0])
        writer.add_page(page)
    with open(output_pdf, "wb") as f_out:
        writer.write(f_out)

# GUI starts here
class PDFStamperApp:
    def __init__(self, master):
        self.master = master
        master.title("PDF Stamper")

        tk.Label(master, text="Select PDF file:").grid(row=0, column=0)
        self.pdf_entry = tk.Entry(master, width=40)
        self.pdf_entry.grid(row=0, column=1)
        tk.Button(master, text="Browse", command=self.browse_pdf).grid(row=0, column=2)

        tk.Label(master, text="Select LEFT image:").grid(row=1, column=0)
        self.left_img_entry = tk.Entry(master, width=40)
        self.left_img_entry.grid(row=1, column=1)
        tk.Button(master, text="Browse", command=self.browse_left_img).grid(row=1, column=2)

        tk.Label(master, text="Select RIGHT image:").grid(row=2, column=0)
        self.right_img_entry = tk.Entry(master, width=40)
        self.right_img_entry.grid(row=2, column=1)
        tk.Button(master, text="Browse", command=self.browse_right_img).grid(row=2, column=2)

        tk.Label(master, text="Chapter:").grid(row=3, column=0)
        self.chapter_var = tk.IntVar(value=1)
        self.chapter_dropdown = ttk.Combobox(master, textvariable=self.chapter_var, values=[i for i in range(1, 21)], width=5, state="readonly")
        self.chapter_dropdown.grid(row=3, column=1, sticky='w')

        tk.Button(master, text="Stamp PDF", command=self.stamp_pdf).grid(row=4, column=1, pady=10)

    def browse_pdf(self):
        file = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if file:
            self.pdf_entry.delete(0, tk.END)
            self.pdf_entry.insert(0, file)

    def browse_left_img(self):
        file = filedialog.askopenfilename(filetypes=[("Image files", "*.png;*.jpg;*.jpeg;*.bmp")])
        if file:
            self.left_img_entry.delete(0, tk.END)
            self.left_img_entry.insert(0, file)

    def browse_right_img(self):
        file = filedialog.askopenfilename(filetypes=[("Image files", "*.png;*.jpg;*.jpeg;*.bmp")])
        if file:
            self.right_img_entry.delete(0, tk.END)
            self.right_img_entry.insert(0, file)

    def stamp_pdf(self):
        input_pdf = self.pdf_entry.get()
        left_img = self.left_img_entry.get()
        right_img = self.right_img_entry.get()
        chapter = self.chapter_var.get()
        if not all([input_pdf, left_img, right_img, chapter]):
            messagebox.showerror("Error", "Please fill all fields!")
            return

        y_value = 610 - (chapter * 31.68)
        output_pdf = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if not output_pdf:
            return

        try:
            stamp_pdf(input_pdf, output_pdf, left_img, right_img, y_value)
            messagebox.showinfo("Success", f"Stamped PDF saved to {os.path.basename(output_pdf)}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

if __name__ == '__main__':
    root = tk.Tk()
    app = PDFStamperApp(root)
    root.mainloop()
