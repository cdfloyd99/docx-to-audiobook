# DOCX to Audiobook Converter

A user-friendly application that converts Microsoft Word documents (`.docx` files) into audiobooks. With a simple graphical interface built using Tkinter, you can select a `.docx` file and generate an audiobook in just a few clicks. The application supports customization options like voice selection, speech rate, and volume control.

## **Features**

- **Easy-to-Use GUI:** Simple and intuitive interface for selecting files and adjusting settings.
- **Text Extraction:** Efficiently extracts text from `.docx` files without the need for external libraries.
- **Text-to-Speech Conversion:**
  - **Local TTS Engine:** Uses `pyttsx3` for offline conversion with available system voices.
  - **gTTS Integration:** Optionally use Google Text-to-Speech for a wider range of voices and languages.
- **Voice Customization:**
  - Adjust speech **rate** and **volume**.
  - Select from available **voices** on your system or through gTTS.
- **Multilingual Support:** Convert documents in different languages using gTTS.
- **Status Updates:** Real-time updates during the conversion process.
- **Threading:** Responsive GUI that remains usable during long operations.

## **Installation**

### **Prerequisites**

- **Python 3.6 or higher** installed on your system.
- Required Python packages:
  - `pyttsx3`
  - `gTTS` (for Google Text-to-Speech version)
  - `simpleaudio` (for voice preview in gTTS version)
  
### **Install Dependencies**

Open a terminal or command prompt and run:

```bash
pip install pyttsx3 gTTS simpleaudio
