from PyPDF2 import PdfReader, PdfWriter
from docx import Document
import pytesseract
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

pdf_file = r"C:\Users\Boyet\Downloads\Telegram Desktop\history\history_outline 2.pdf"
output_folder = r"C:\Users\Boyet\Downloads\Telegram Desktop\history\output"

reader = PdfReader(pdf_file)

for i, page in enumerate(reader.pages, start=1):
    # Extract text per page
    text = page.extract_text()
    
    # Create Word document
    doc = Document()
    doc.add_paragraph(text)
    
    # Save Word file per page
    word_path = f"{output_folder}\\page_{i}.docx"
    doc.save(word_path)

print("✅ Done! Each page saved as a separate Word doc.")
