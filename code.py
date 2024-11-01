import tkinter as tk
from tkinter import filedialog
from PIL import Image
import img2pdf
import os
from pathlib import Path

try:
    import comtypes.client as comtypes
except ImportError:
    comtypes = None  # Ensure comtypes is installed if working with .doc files

# Initialize the Tkinter file dialog
def open_file_dialog():
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    file_path = filedialog.askopenfilename(filetypes=[("Word Files", "*.doc *.docx")])
    return file_path

# Convert Word file to images and return list of image paths
def convert_word_to_images(file_path):
    images = []
    output_dir = Path(file_path).parent / "temp_images"
    output_dir.mkdir(exist_ok=True)

    if file_path.endswith('.docx'):
        from docx2pdf import convert
        pdf_path = output_dir / "temp_pdf.pdf"
        convert(file_path, pdf_path)
    elif file_path.endswith('.doc') and comtypes:
        word = comtypes.CreateObject("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(file_path)
        pdf_path = output_dir / "temp_pdf.pdf"
        doc.SaveAs(pdf_path, FileFormat=17)
        doc.Close()
        word.Quit()
    else:
        raise ValueError("Only .doc and .docx files are supported.")

    # Convert PDF pages to images
    from pdf2image import convert_from_path
    images = convert_from_path(pdf_path, dpi=300, output_folder=output_dir, fmt="png")
    return [str(image.filename) for image in images]

# Convert images to a single PDF
def images_to_pdf(images, output_path):
    with open(output_path, "wb") as f:
        f.write(img2pdf.convert(images))

# Main function to execute all steps
def main():
    # Step 1: Open file dialog to choose a Word file
    word_file = open_file_dialog()
    if not word_file:
        print("No file selected.")
        return

    # Step 2: Convert Word to images and get paths of images
    images = convert_word_to_images(word_file)
    
    # Step 3: Convert all images to a single PDF
    original_file = Path(word_file)
    output_pdf = original_file.with_name(original_file.stem + "_new.pdf")
    images_to_pdf(images, output_pdf)

    print(f"PDF saved as: {output_pdf}")

# Run the script
if __name__ == "__main__":
    main()