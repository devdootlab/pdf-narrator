# main.py
# A desktop application to display a PDF and narrate a corresponding script
# with synchronized zooming, built as part of a 5-day coding sprint.

import tkinter as tk
from tkinter import ttk
import json
import pymupdf  # PyMuPDF library for PDF manipulation
from PIL import Image, ImageTk  # Pillow library for image handling
import threading
import time  # Time module imported for sleep functionality
import asyncio
import edge_tts
import sounddevice as sd
import soundfile as sf
import win32com.client
import tkinter.filedialog
from screeninfo import get_monitors

class App(tk.Tk):
    """
    A dynamic PDF narrator application that synchronizes audio narration
    with zoomed views of a PDF document.
    """
    def __init__(self):
        """Initializes the main application window and all its components."""
        super().__init__()

        # --- Find second monitor (or use primary if only one) ---
        monitors = get_monitors()
        if len(monitors) > 1:
            mon = monitors[1]  # Second monitor
        else:
            mon = monitors[0]  # Primary monitor

        # --- Set window size and position ---
        default_width = 850
        default_height = 1100
        double_width = default_width * 2
        x = mon.x
        y = mon.y

        self.geometry(f"{double_width}x{default_height}+{x}+{y}")
        self.title("PDF Narrator")

        # --- Application State Variables ---
        self.doc = None  # Will hold the loaded PyMuPDF document object.
        self.script_data = None  # Will hold the parsed JSON narration script.
        self.tk_img = None  # A reference to the current page image to prevent garbage collection.
        self.current_step = 0  # The index of the current step in the narration script.
        self.is_narrating = False  # A flag to prevent multiple narrations from starting.

        # --- Layout Frames ---
        main_frame = ttk.Frame(self)
        main_frame.pack(fill=tk.BOTH, expand=True)

        control_frame = ttk.Frame(self)
        control_frame.pack(fill=tk.X)

        # --- Widgets ---
        self.canvas = tk.Canvas(main_frame, bg="gray")
        self.canvas.pack(fill=tk.BOTH, expand=True)

        self.start_button = ttk.Button(
            control_frame,
            text="Start Narration",
            command=self.start_narration
        )
        self.start_button.pack(side=tk.LEFT, padx=5, pady=5)

        self.skip_button = ttk.Button(
            control_frame,
            text="Skip Narration",
            command=self.skip_narration
        )
        self.skip_button.pack(side=tk.LEFT, padx=5, pady=5)

        self.next_button = ttk.Button(
            control_frame,
            text="Next Figure",
            command=self.next_figure
        )
        self.next_button.pack(side=tk.LEFT, padx=5, pady=5)

        self.open_button = ttk.Button(
            control_frame,
            text="Open PDF+Script",
            command=self.open_new_pdf_and_script
        )
        self.open_button.pack(side=tk.LEFT, padx=5, pady=5)

        self.speed_label = ttk.Label(control_frame, text="Speech Speed")
        self.speed_label.pack(side=tk.LEFT, padx=5, pady=5)

        self.speed_var = tk.DoubleVar(value=1.5)
        self.speed_slider = ttk.Scale(
            control_frame,
            from_=0.5, to=2.0, variable=self.speed_var,
            orient=tk.HORIZONTAL
        )
        self.speed_slider.pack(side=tk.LEFT, padx=5, pady=5)

        self.narration_enabled = tk.BooleanVar(value=True)
        self.narration_checkbox = ttk.Checkbutton(
            control_frame,
            text="Enable Narration",
            variable=self.narration_enabled
        )
        self.narration_checkbox.pack(side=tk.LEFT, padx=5, pady=5)

    def open_new_pdf_and_script(self):
        """
        Prompts the user to select a new PDF and JSON script, then reloads them.
        """
        pdf_path = tkinter.filedialog.askopenfilename(
            title="Select PDF file",
            filetypes=[("PDF files", "*.pdf")]
        )
        if not pdf_path:
            print("No PDF selected.")
            return

        script_path = tkinter.filedialog.askopenfilename(
            title="Select narration script (JSON)",
            filetypes=[("JSON files", "*.json")]
        )
        if not script_path:
            print("No script selected.")
            return

        self.load_files(pdf_path, script_path)

    def load_files(self, pdf_path, script_path):
        """Loads the PDF and script files."""
        self.doc = self.load_pdf(pdf_path)
        self.script_data = self.load_script(script_path)
        self.current_step = 0
        self.is_narrating = False
        if self.doc:
            self.display_page(0)
        else:
            self.canvas.delete("all")
            error_text = f"Failed to load {pdf_path}"
            self.canvas.create_text(425, 550, text=error_text, font=("Arial", 16))

    def load_pdf(self, filepath):
        """
        Opens a PDF file using PyMuPDF and returns the document object.
        """
        try:
            doc = pymupdf.open(filepath)
            print(f"Successfully loaded PDF: {filepath}")
            return doc
        except Exception as e:
            print(f"Error loading PDF: {e}")
            return None

    def load_script(self, filepath):
        """
        Opens and parses a JSON narration script.
        """
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                data = json.load(f)
                print(f"Successfully loaded script: {filepath}")
                return data
        except Exception as e:
            print(f"Error loading script: {e}")
            return None

    def display_page(self, page_num, zoom_rect=None):
        """
        Renders a specific region of a PDF page onto the main canvas,
        dynamically calculating the zoom to fit the canvas.
        """
        if not self.doc or page_num >= self.doc.page_count:
            return

        page = self.doc.load_page(page_num)
        
        # Use the full page if no zoom_rect is provided
        clip_rect = pymupdf.Rect(zoom_rect) if zoom_rect else page.rect
        
        # --- DYNAMIC ZOOM CALCULATION ---
        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas.winfo_height()

        # To avoid division by zero on first render before canvas is sized
        if canvas_width == 1 or canvas_height == 1:
            self.after(100, lambda: self.display_page(page_num, zoom_rect))
            return

        rect_width = clip_rect.width
        rect_height = clip_rect.height

        if rect_width <= 0 or rect_height <= 0:
             # Fallback to full page if rect is invalid
            clip_rect = page.rect
            rect_width = clip_rect.width
            rect_height = clip_rect.height

        zoom_x = canvas_width / rect_width
        zoom_y = canvas_height / rect_height
        
        # Use the smaller zoom factor to maintain aspect ratio and fit the whole area
        zoom_factor = min(zoom_x, zoom_y)
        
        mat = pymupdf.Matrix(zoom_factor, zoom_factor)
        pix = page.get_pixmap(matrix=mat, clip=clip_rect)

        mode = "RGBA" if pix.alpha else "RGB"
        img = Image.frombytes(mode, [pix.width, pix.height], pix.samples)
        self.tk_img = ImageTk.PhotoImage(img)

        self.canvas.delete("all")
        # Center the image on the canvas
        x_pos = (canvas_width - pix.width) / 2
        y_pos = (canvas_height - pix.height) / 2
        self.canvas.create_image(x_pos, y_pos, anchor='nw', image=self.tk_img)

    def speak(self, text):
        """
        Speaks the given text using Microsoft Speech Platform (SAPI).
        Speech rate is controlled by self.speed_var.
        """
        try:
            speaker = win32com.client.Dispatch("SAPI.SpVoice")
            voices = speaker.GetVoices()
            if voices.Count > 0:
                speaker.Voice = voices.Item(0)
            
            rate = int((self.speed_var.get() - 1.25) * 8)
            speaker.Rate = rate
            
            speaker.Speak(text)
        except Exception as e:
            print(f"Error initializing speech engine: {e}")

    def start_narration(self):
        """
        Begins the narration process from the first step.
        """
        if self.is_narrating or not self.script_data:
            return

        print("--- Starting Narration ---")
        self.is_narrating = True
        self.current_step = 0
        self.process_narration_step()

    def process_narration_step(self):
        """
        Processes a single narration step and schedules the next one.
        """
        if not self.is_narrating or self.current_step >= len(self.script_data['narration_steps']):
            print("--- Narration Finished ---")
            self.is_narrating = False
            self.display_page(0)
            return

        step_data = self.script_data['narration_steps'][self.current_step]
        print(f"Processing step {self.current_step}: Zooming to page {step_data['page_number']}")

        self.display_page(
            page_num=step_data['page_number'],
            zoom_rect=step_data.get('zoom_rect')
        )
        self.update_idletasks()

        self.after(
            step_data['pre_speech_delay_ms'],
            lambda: self._speak_and_continue(step_data['narration_text'])
        )

    def _speak_and_continue(self, text):
        if self.narration_enabled.get():
            print(f"Narrating: {text[:80]}...")
            self.narration_thread = threading.Thread(target=self._narrate_step, args=(text,), daemon=True)
            self.narration_thread.start()
        else:
            self.after(1000, self._advance_narration_step)

    def _narrate_step(self, text):
        self.speak(text)
        self.after(0, self._advance_narration_step)

    def _advance_narration_step(self):
        if self.is_narrating:
            self.current_step += 1
            self.process_narration_step()

    def skip_narration(self):
        """Skips the current speaking part and moves to the next step."""
        if self.is_narrating:
            self.current_step += 1
            self.stop_narration()
            self.after(100, self.process_narration_step)

    def stop_narration(self):
        """Stops the narration loop."""
        print("--- Narration Stopped ---")
        self.is_narrating = False

    def next_figure(self):
        """Advances to the next step that involves a page change."""
        if not self.script_data:
            return
        
        current_page = self.script_data['narration_steps'][self.current_step]['page_number']
        
        # Find the next step with a different page number that is not the title page
        next_step_index = self.current_step + 1
        while next_step_index < len(self.script_data['narration_steps']):
            next_page = self.script_data['narration_steps'][next_step_index]['page_number']
            if next_page != current_page and next_page != 0:
                self.current_step = next_step_index
                if self.is_narrating:
                    self.stop_narration()
                    self.after(100, self.process_narration_step)
                else:
                    self.process_narration_step()
                return
            next_step_index += 1
        
        print("No subsequent figure found.")


if __name__ == "__main__":
    app = App()
    try:
        app.load_files("2024.10.10.617658v3.full.pdf", "orthrus1.json")
    except Exception as e:
        print(f"Could not auto-load files on startup: {e}")
    app.mainloop()
