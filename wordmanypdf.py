import os
import comtypes.client
import subprocess
import time

# === Settings ===
input_folder = r"C:\Users\Boyet\Pictures\flipbooks\word_files"
output_folder = r"C:\Users\Boyet\Pictures\flipbooks\word_files\output"

# Make sure output folder exists
os.makedirs(output_folder, exist_ok=True)

# 1. Kill any running WINWORD.EXE to avoid RPC errors
subprocess.call("taskkill /f /im WINWORD.EXE", shell=True)

# 2. Start Word fresh
word = comtypes.client.CreateObject("Word.Application")
time.sleep(2)  # give Word some time to initialize
word.Visible = False

# 3. Loop through files
for file in os.listdir(input_folder):
    if file.lower().endswith((".doc", ".docx")):
        doc_path = os.path.join(input_folder, file)
        pdf_path = os.path.join(output_folder, os.path.splitext(file)[0] + ".pdf")

        try:
            doc = word.Documents.Open(doc_path)
            doc.SaveAs(pdf_path, FileFormat=17)  # 17 = PDF
            doc.Close()
            print(f"✅ Converted: {file} → {os.path.basename(pdf_path)}")
        except Exception as e:
            print(f"❌ Failed: {file} ({e})")

# 4. Quit Word
word.Quit()
print("🎉 All conversions finished!")
