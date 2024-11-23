import tkinter as tk
from tkinter import filedialog, messagebox
import pyttsx3
import threading
import zipfile
import xml.etree.ElementTree as ET

def select_file():
    filepath = filedialog.askopenfilename(
        filetypes=[("Word Documents", "*.docx")])
    if filepath:
        docx_path.set(filepath)

def select_output_file():
    filepath = filedialog.asksaveasfilename(
        defaultextension=".mp3",
        filetypes=[("MP3 Audio Files", "*.mp3"), ("WAV Audio Files", "*.wav")])
    if filepath:
        output_path.set(filepath)

def convert_to_audio():
    if not docx_path.get():
        messagebox.showerror("Error", "Please select a .docx file.")
        return
    if not output_path.get():
        messagebox.showerror("Error", "Please select an output file.")
        return
    threading.Thread(target=process_conversion).start()

def extract_text_from_docx(docx_file):
    with zipfile.ZipFile(docx_file) as docx_zip:
        xml_content = docx_zip.read('word/document.xml')
    tree = ET.fromstring(xml_content)
    paragraphs = []
    for paragraph in tree.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p'):
        texts = [node.text
                 for node in paragraph.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
                 if node.text]
        if texts:
            paragraphs.append(''.join(texts))
    return '\n'.join(paragraphs)

def process_conversion():
    try:
        btn_convert.config(state='disabled')
        status.set("Reading .docx file...")
        text = extract_text_from_docx(docx_path.get())
        if not text.strip():
            raise ValueError("The selected .docx file contains no text.")
        status.set("Converting to audio...")
        engine = pyttsx3.init()
        engine.setProperty('rate', rate.get())
        engine.setProperty('volume', volume.get())
        selected_voice_name = voice_id.get()
        selected_voice_id = voice_dict[selected_voice_name]
        engine.setProperty('voice', selected_voice_id)
        engine.save_to_file(text, output_path.get())
        engine.runAndWait()
        status.set("Conversion completed!")
        messagebox.showinfo("Success", "Audiobook created successfully.")
    except Exception as e:
        messagebox.showerror("Error", str(e))
    finally:
        btn_convert.config(state='normal')

# Initialize TTS engine and get available voices
engine = pyttsx3.init()
voices = engine.getProperty('voices')
voice_dict = {v.name: v.id for v in voices}
voice_names = list(voice_dict.keys())

# Set default voice
default_voice = voice_names[0] if voice_names else ''

# Create the main window
root = tk.Tk()
root.title(".docx to Audiobook Converter")
root.geometry("500x400")
root.resizable(False, False)

# Tkinter variables (must be created after root)
docx_path = tk.StringVar()
output_path = tk.StringVar()
status = tk.StringVar()
rate = tk.IntVar(value=200)
volume = tk.DoubleVar(value=1.0)
voice_id = tk.StringVar(value=default_voice)

# DOCX file selection
tk.Label(root, text="Select .docx File:").pack(pady=5)
tk.Entry(root, textvariable=docx_path, width=60).pack()
tk.Button(root, text="Browse", command=select_file).pack(pady=5)

# Output file selection
tk.Label(root, text="Select Output File:").pack(pady=5)
tk.Entry(root, textvariable=output_path, width=60).pack()
tk.Button(root, text="Save As", command=select_output_file).pack(pady=5)

# Voice options frame
options_frame = tk.Frame(root)
options_frame.pack(pady=10)

# Rate
tk.Label(options_frame, text="Rate:").grid(row=0, column=0, padx=5, pady=5)
tk.Scale(options_frame, from_=50, to=300, orient=tk.HORIZONTAL, variable=rate).grid(row=0, column=1, padx=5, pady=5)

# Volume
tk.Label(options_frame, text="Volume:").grid(row=1, column=0, padx=5, pady=5)
tk.Scale(options_frame, from_=0.0, to=1.0, resolution=0.1, orient=tk.HORIZONTAL, variable=volume).grid(row=1, column=1, padx=5, pady=5)

# Voice selection
tk.Label(options_frame, text="Voice:").grid(row=2, column=0, padx=5, pady=5)
tk.OptionMenu(options_frame, voice_id, *voice_names).grid(row=2, column=1, padx=5, pady=5)

# Convert button
btn_convert = tk.Button(root, text="Convert to Audiobook", command=convert_to_audio)
btn_convert.pack(pady=20)

# Status label
tk.Label(root, textvariable=status).pack(pady=5)

# Run the application
root.mainloop()
