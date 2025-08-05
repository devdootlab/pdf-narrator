import tkinter as tk
from tkinter import ttk
import pymupdf
from PIL import Image, ImageTk # Import Pillow

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF Narrator")
        self.geometry("850x1100") # Adjust size for a typical page

        # Create a canvas to display the PDF page
        self.canvas = tk.Canvas(self, bg="gray")
        self.canvas.pack(fill=tk.BOTH, expand=True)

        self.doc = self.load_pdf("sample.pdf")
        if self.doc:
            # Define a rectangle for the top-left quadrant of the page to zoom into
            zoom_area = pymupdf.Rect(0, 0, 400, 200) 
            # Call display_page with the zoom area
            self.display_page(page_num=0, zoom_rect=zoom_area, zoom_factor=2.5)

    def load_pdf(self, filepath):
        try:
            doc = pymupdf.open(filepath)
            return doc
        except Exception as e:
            print(f"Error loading PDF: {e}")
            # You might want to display this error on the canvas as well
            self.canvas.create_text(425, 550, text=f"Error: Could not load '{filepath}'", font=("Arial", 16))
            return None

    def display_page(self, page_num, zoom_rect=None, zoom_factor=2.0):
        if not self.doc or page_num >= self.doc.page_count:
            return

        page = self.doc.load_page(page_num)

        # If no zoom rectangle is provided, use the full page
        if zoom_rect is None:
            clip_rect = page.rect
        else:
            # Ensure the provided rect is a pymupdf.Rect object
            clip_rect = pymupdf.Rect(zoom_rect)

        # Create a zoom matrix
        mat = pymupdf.Matrix(zoom_factor, zoom_factor)

        # Render the clipped and zoomed region
        pix = page.get_pixmap(matrix=mat, clip=clip_rect)

        # Convert the pixmap to a format Tkinter can use
        mode = "RGBA" if pix.alpha else "RGB"
        img = Image.frombytes(mode, [pix.width, pix.height], pix.samples)
        
        # NOTE: self.tk_img must be an instance variable to prevent garbage collection
        self.tk_img = ImageTk.PhotoImage(img)
        
        self.canvas.delete("all")
        self.canvas.create_image(0, 0, anchor='nw', image=self.tk_img)


if __name__ == "__main__":
    app = App()
    app.mainloop() # Start the event loop