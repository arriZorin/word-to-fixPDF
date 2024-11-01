import tkinter as tk
from tkinter import filedialog
from PIL import Image
import img2pdf
import os
from time import sleep
from pathlib import Path
from pdf2image import convert_from_path

try:
    import comtypes.client as comtypes
except ImportError:
    comtypes = None  # Ensure comtypes is installed if working with .doc files

# Initialize the Tkinter file dialog
def open_file_dialog():
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    file_path = filedialog.askopenfilenames(filetypes=[("Word Files", "*.doc *.docx")])
    return list(file_path)

# Convert Word file to images and return list of image paths
def convert_word_to_images(file_path):
    images = []
    output_dir = Path(file_path).parent / "temp_images"
    output_dir.mkdir(exist_ok=True)
    
    # Ensure the file path is in the correct format for comtypes and Windows
    file_path = Path(file_path).resolve(strict=True)

    if file_path.suffix == '.docx':
        from docx2pdf import convert
        pdf_path = output_dir / "temp_pdf.pdf"
        convert(str(file_path), str(pdf_path))
    elif file_path.suffix == '.doc' and comtypes:
        word = comtypes.CreateObject("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(str(file_path))  # Convert Path to string
        pdf_path = output_dir / "temp_pdf.pdf"
        doc.SaveAs(str(pdf_path), FileFormat=17)
        doc.Close()
        word.Quit()
    else:
        raise ValueError("Only .doc and .docx files are supported.")

    # Convert PDF pages to images
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
    # print(word_file)
    # sleep(10)
    # return
    
    if not word_file:
        print("No file selected.")
        return

    for x, y in enumerate(word_file):
        # Step 2: Convert Word to images and get paths of images
        images = convert_word_to_images(y)
        
        # Step 3: Convert all images to a single PDF
        original_file = Path(y)
        output_pdf = original_file.with_name(original_file.stem + ".pdf")
        images_to_pdf(images, output_pdf)

        print(f"{x+1}/{len(word_file)} PDF saved as: {output_pdf}")
    print("\n--------- Finished ---------")
    sleep(2)

# Run the script
if __name__ == "__main__":
    main()