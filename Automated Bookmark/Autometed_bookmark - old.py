from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from PyPDF2 import PdfReader, PdfWriter
from io import BytesIO

PAGE_WIDTH = 612  # 8.5 inch
PAGE_HEIGHT = 766.8  # 10.65 inch

def create_overlay_with_image(image_path, position):
    packet = BytesIO()
    can = canvas.Canvas(packet, pagesize=(PAGE_WIDTH, PAGE_HEIGHT))

    # Load image with ImageReader to get actual size
    img = ImageReader(image_path)
    img_width, img_height = img.getSize()

    # Example scaling if needed (optional):
    scale = 0.2  # or compute scale dynamically if you want
    img_width *= scale
    img_height *= scale

    if position == 'right':
        x = PAGE_WIDTH - img_width - 0  # 10 pts padding from right edge
        y = 356.56 # in pptx 1st chapter bookmark horizontal position in points 610, (10.65-0.7)*72 [1 inch = 72 points], next chapter position y-31.68=610-31.68 points. so, for chapter 16 should be 610-(31.68*16) 
        can.drawImage(img, x, y, width=img_width, height=img_height, mask='auto')

    elif position == 'left':
        x = 0  # 10 pts padding from left edge
        y = 356.56 # in pptx 1st chapter bookmark horizontal position (10.65-0.7)*72 [1 inch = 72 points], next chapter position y-31.68=610-31.68 points
        can.drawImage(img, x, y, width=img_width, height=img_height, mask='auto')

    can.save()
    packet.seek(0)
    return PdfReader(packet)

def stamp_pdf(input_pdf, output_pdf, left_image_path, right_image_path):
    reader = PdfReader(input_pdf)
    writer = PdfWriter()

    for i, page in enumerate(reader.pages):
        pos = 'right' if (i + 1) % 2 == 1 else 'left'
        image_path = right_image_path if pos == 'right' else left_image_path
        overlay = create_overlay_with_image(image_path, pos)
        page.merge_page(overlay.pages[0])
        writer.add_page(page)

    with open(output_pdf, "wb") as f_out:
        writer.write(f_out)

# Example call:
stamp_pdf(
    r"D:\Office Files\Developer\Automated Bookmark\AP - Class 9 - General Mathematics - অধ্যায় ০৮ - বৃত্ত - Preparation Book.pdf",
    r"D:\Office Files\Developer\Automated Bookmark\your_output.pdf",
    r"D:\Office Files\Developer\Automated Bookmark\left.png",
    r"D:\Office Files\Developer\Automated Bookmark\right.png"
)
