import fitz                   # PyMuPDF
import numpy as np

def brighten_grayscale(pix: fitz.Pixmap, factor: float = 1.4) -> fitz.Pixmap:
    """
    Return a new Pixmap that is the original pixmap brightened in-place.
    `factor` > 1.0 lightens, < 1.0 darkens.
    Works only on GRAY pixmaps.
    """
    if pix.n != 1:
        raise ValueError("Expected a grayscale pixmap (n == 1)")

    # Copy pixel bytes to NumPy, scale, clip back to 0-255
    buf = np.frombuffer(pix.samples, dtype=np.uint8)
    lighter = np.clip(buf * factor, 0, 255).astype(np.uint8)

    # Re-create a Pixmap from the modified bytes  (note the final `False` = no alpha)
    return fitz.Pixmap(fitz.csGRAY, pix.width, pix.height, lighter.tobytes(), False)

def convert_pdf_to_light_grayscale(src_pdf: str,
                                   dst_pdf: str,
                                   brightness_factor: float = 1.4,
                                   dpi: int = 144):
    doc_in  = fitz.open(src_pdf)
    doc_out = fitz.open()              # empty PDF to collect pages

    for pno in range(doc_in.page_count):
        page = doc_in.load_page(pno)

        # Render the page to a *grayscale* pixmap at chosen resolution
        pix  = page.get_pixmap(dpi=dpi, colorspace=fitz.csGRAY)

        # Brighten it
        pix_lighter = brighten_grayscale(pix, brightness_factor)

        # Wrap the pixmap into a one-page PDF
        img_pdf = fitz.open()
        rect    = fitz.Rect(0, 0, pix_lighter.width, pix_lighter.height)
        img_page = img_pdf.new_page(width=rect.width, height=rect.height)
        img_page.insert_image(rect, pixmap=pix_lighter)

        doc_out.insert_pdf(img_pdf)

    doc_out.save(dst_pdf)
    doc_in.close()
    doc_out.close()
    print(f"✓ Saved lighter-grayscale PDF as '{dst_pdf}'")

# ---- Example usage ----
convert_pdf_to_light_grayscale(
    "input.pdf",
    "output_light_grayscale.pdf",
    brightness_factor=1.9   # tweak to taste
)
