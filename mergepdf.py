import os

try:
    from pypdf import PdfMerger
except ImportError:
    from PyPDF2 import PdfMerger  # fallback

# === Settings ===
input_folder = r"C:\Users\Boyet\Pictures\flipbooks\word_files\output"
output_file = r"C:\Users\Boyet\Pictures\flipbooks\word_files\merged_output.pdf"

merger = PdfMerger()

for file in sorted(os.listdir(input_folder)):
    if file.lower().endswith(".pdf"):
        merger.append(os.path.join(input_folder, file))
        print(f"📄 Added: {file}")

merger.write(output_file)
merger.close()

print(f"🎉 All PDFs merged into: {output_file}")
