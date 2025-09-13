from docx import Document
from docx.shared import Inches
import os
import pytesseract
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"


# 📂 Folder containing images
img_folder = r"C:\Users\Boyet\Pictures\flipbooks"
output_folder = r"C:\Users\Boyet\Pictures\flipbooks\pdf"

# Make sure output folder exists
os.makedirs(output_folder, exist_ok=True)

# Get all image files (sorted)
images = [f for f in os.listdir(img_folder) if f.lower().endswith(('.png', '.jpg', '.jpeg'))]
images.sort()

# Create one Word file per image
for img in images:
    img_path = os.path.join(img_folder, img)
    
    # Create new Word document
    doc = Document()
    doc.add_picture(img_path, width=Inches(6))
    
    # Save with same name but .docx
    base_name = os.path.splitext(img)[0]
    output_file = os.path.join(output_folder, f"{base_name}.docx")
    doc.save(output_file)
    print(f"✅ Saved: {output_file}")
