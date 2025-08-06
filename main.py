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

class App(tk.Tk):
    """
    A dynamic PDF narrator application that synchronizes audio narration
    with zoomed views of a PDF document.
    """
    def __init__(self):
        """Initializes the main application window and all its components."""
        super().__init__()

        # --- Configure the root window ---
        self.title("PDF Narrator")
        self.geometry("850x1100")

        # --- Application State Variables ---
        self.doc = None  # Will hold the loaded PyMuPDF document object.
        self.script_data = None  # Will hold the parsed JSON narration script.
        self.tk_img = None  # A reference to the current page image to prevent garbage collection.
        self.current_step = 0  # The index of the current step in the narration script.
        self.is_narrating = False  # A flag to prevent multiple narrations from starting.

        # --- Layout Frames ---
        # Main frame for the PDF canvas, allowing it to expand.
        main_frame = ttk.Frame(self)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Control frame for buttons at the bottom.
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

        self.speed_label = ttk.Label(control_frame, text="Speech Speed")
        self.speed_label.pack(side=tk.LEFT, padx=5, pady=5)

        self.speed_var = tk.DoubleVar(value=1.5)  # This maps to +5
        self.speed_slider = ttk.Scale(
            control_frame,
            from_=0.5, to=2.0, variable=self.speed_var,
            orient=tk.HORIZONTAL
        )
        self.speed_slider.pack(side=tk.LEFT, padx=5, pady=5)

        # --- Initial Setup ---
        # Prompt for PDF file
        pdf_path = tkinter.filedialog.askopenfilename(
            title="Select PDF file",
            filetypes=[("PDF files", "*.pdf")]
        )
        if not pdf_path:
            print("No PDF selected. Exiting.")
            self.destroy()
            return

        # Prompt for JSON script file
        script_path = tkinter.filedialog.askopenfilename(
            title="Select narration script (JSON)",
            filetypes=[("JSON files", "*.json")]
        )
        if not script_path:
            print("No script selected. Exiting.")
            self.destroy()
            return

        # Load the required files
        self.doc = self.load_pdf(pdf_path)
        self.script_data = self.load_script(script_path)

        # Display the first page of the PDF initially (not zoomed).
        if self.doc:
            self.display_page(0)
        else:
            # Display an error if the PDF could not be loaded.
            error_text = f"Failed to load {pdf_path}"
            self.canvas.create_text(425, 550, text=error_text, font=("Arial", 16))

    def load_pdf(self, filepath):
        """
        Opens a PDF file using PyMuPDF and returns the document object.

        Args:
            filepath (str): The path to the PDF file.

        Returns:
            pymupdf.Document or None: The opened document object, or None if an error occurs.
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

        Args:
            filepath (str): The path to the JSON script file.

        Returns:
            dict or None: A dictionary containing the parsed JSON data, or None if an error occurs.
        """
        try:
            with open(filepath, 'r') as f:
                data = json.load(f)
                print(f"Successfully loaded script: {filepath}")
                return data
        except Exception as e:
            print(f"Error loading script: {e}")
            return None

    def display_page(self, page_num, zoom_rect=None, zoom_factor=2.0):
        """
        Renders a specific region of a PDF page onto the main canvas.

        Args:
            page_num (int): The 0-indexed page number to display.
            zoom_rect (list, optional): A list of four coordinates [x0, y0, x1, y1]
                defining the clip box. Defaults to None (full page).
            zoom_factor (float, optional): The scaling factor for the zoom.
                Defaults to 2.0.
        """
        if not self.doc or page_num >= self.doc.page_count:
            return

        page = self.doc.load_page(page_num)

        # Determine the clipping rectangle for the zoom.
        if zoom_rect is None:
            clip_rect = page.rect
        else:
            clip_rect = pymupdf.Rect(zoom_rect)

        # Create a zoom matrix from the zoom factor.
        mat = pymupdf.Matrix(zoom_factor, zoom_factor)

        # Render the specified page region to a pixmap (an image).
        pix = page.get_pixmap(matrix=mat, clip=clip_rect)

        # Convert the pixmap to a format Tkinter can use via the Pillow library.
        mode = "RGBA" if pix.alpha else "RGB"
        img = Image.frombytes(mode, [pix.width, pix.height], pix.samples)

        # This must be an instance variable to prevent Python's garbage collector
        # from discarding the image before it's displayed.
        self.tk_img = ImageTk.PhotoImage(img)

        # Update the canvas with the new page image.
        self.canvas.delete("all")
        self.canvas.create_image(0, 0, anchor='nw', image=self.tk_img)

    def speak(self, text):
        """
        Speaks the given text using Microsoft Speech Platform (SAPI).
        Speech rate is controlled by self.speed_var.
        """
        speaker = win32com.client.Dispatch("SAPI.SpVoice")
        # Example: select the second installed voice
        voices = speaker.GetVoices()
        # You can use a dropdown to let the user choose
        speaker.Voice = voices.Item(0)  # Change index for different voices
        rate = int((self.speed_var.get() - 1.0) * 10)
        speaker.Rate = rate
        speaker.Speak(text)

    def start_narration(self):
        """
        Begins the narration process from the first step.
        Runs the narration loop using Tkinter's event loop.
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
            self.display_page(0)  # Reset to the full first page view
            return

        step_data = self.script_data['narration_steps'][self.current_step]
        print(f"Processing step {self.current_step}: Zooming to page {step_data['page_number']}")

        self.display_page(
            page_num=step_data['page_number'],
            zoom_rect=step_data['zoom_rect']
        )
        self.update_idletasks()

        # Schedule the narration after the pre-speech delay
        self.after(
            step_data['pre_speech_delay_ms'],
            lambda: self._speak_and_continue(step_data['narration_text'])
        )

    def _speak_and_continue(self, text):
        print(f"Narrating: {text}")
        self.narration_thread = threading.Thread(target=self._narrate_step, args=(text,), daemon=True)
        self.narration_thread.start()

    def _narrate_step(self, text):
        self.speak(text)
        # Schedule next step on main thread after narration finishes
        self.after(0, self._advance_narration_step)

    def _advance_narration_step(self):
        self.current_step += 1
        self.process_narration_step()

    def skip_narration(self):
        # Try to stop current narration thread (not possible with playsound, but you can skip to next step)
        self.current_step += 1
        self.process_narration_step()

    def next_figure(self):
        self.current_step += 1
        self.process_narration_step()

    def list_voices(self):
        speaker = win32com.client.Dispatch("SAPI.SpVoice")
        for voice in speaker.GetVoices():
            print(voice.GetDescription())

# The standard Python entry point. This block ensures the code inside
# only runs when the script is executed directly from the command line.
if __name__ == "__main__":
    app = App()
    app.mainloop()