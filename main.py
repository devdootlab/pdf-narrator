import tkinter as tk
from tkinter import ttk
import pymupdf
from PIL import Image, ImageTk
import json
import pyttsx3 # Add pyttsx3 to imports [cite: 501]

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF Narrator")
        self.geometry("850x1100")

        # --- Layout Frames ---
        # Main frame for the PDF canvas
        main_frame = ttk.Frame(self)
        main_frame.pack(fill=tk.BOTH, expand=True) # [cite: 529, 530]

        # Control frame for buttons at the bottom
        control_frame = ttk.Frame(self)
        control_frame.pack(fill=tk.X) # [cite: 534]

        # --- Widgets ---
        self.canvas = tk.Canvas(main_frame, bg="gray")
        self.canvas.pack(fill=tk.BOTH, expand=True) # [cite: 531, 532]

        # "Start Narration" button in the control frame
        self.start_button = ttk.Button(
            control_frame,
            text="Start Narration",
            command=self.start_narration
        ) # [cite: 535, 536]
        self.start_button.pack(side=tk.LEFT, padx=5, pady=5) # [cite: 537]

        # --- Initial Setup ---
        # Initialize the TTS engine [cite: 505, 506]
        self.tts_engine = pyttsx3.init()

        self.doc = self.load_pdf("sample.pdf")
        self.script_data = self.load_script("script.json")

        # Display the first page of the PDF initially (not zoomed) [cite: 539, 540]
        if self.doc:
            self.display_page(0)
        else:
            error_text = "Failed to load sample.pdf"
            self.canvas.create_text(425, 550, text=error_text, font=("Arial", 16))

    def load_pdf(self, filepath):
        try:
            doc = pymupdf.open(filepath)
            return doc
        except Exception as e:
            print(f"Error loading PDF: {e}")
            return None

    def load_script(self, filepath):
        try:
            with open(filepath, 'r') as f:
                data = json.load(f)
                return data
        except Exception as e:
            print(f"Error loading script: {e}")
            return None

    def display_page(self, page_num, zoom_rect=None, zoom_factor=2.0):
        if not self.doc or page_num >= self.doc.page_count:
            return

        page = self.doc.load_page(page_num)

        if zoom_rect is None:
            clip_rect = page.rect
        else:
            clip_rect = pymupdf.Rect(zoom_rect)

        mat = pymupdf.Matrix(zoom_factor, zoom_factor)
        pix = page.get_pixmap(matrix=mat, clip=clip_rect)

        mode = "RGBA" if pix.alpha else "RGB"
        img = Image.frombytes(mode, [pix.width, pix.height], pix.samples)
        self.tk_img = ImageTk.PhotoImage(img)
        self.canvas.delete("all")
        self.canvas.create_image(0, 0, anchor='nw', image=self.tk_img)

    def speak(self, text):
        """Speaks the given text using the TTS engine."""
        print(f"Narrating: {text}") # [cite: 511]
        self.tts_engine.say(text) # [cite: 512]
        self.tts_engine.runAndWait() # This is a blocking call [cite: 513]

    def start_narration(self):
        """Starts the narration process."""
        if not self.script_data:
            self.speak("No script is loaded.") # [cite: 543, 544]
            return # [cite: 545]

        # For now, just speak the text from the first step [cite: 547]
        first_step_text = self.script_data['narration_steps'][0]['narration_text'] # [cite: 548]
        self.speak(first_step_text) # [cite: 548]


if __name__ == "__main__":
    app = App()
    app.mainloop()