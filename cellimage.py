from openpyxl import Workbook
from openpyxl.drawing.image import Image
import os

# === Settings ===
img_folder = r"C:\Users\Boyet\Downloads\Telegram Desktop"   # folder containing your images
output_file = r"C:\Users\Boyet\Downloads\Telegram Desktop\output.xlsx"

# Create workbook and sheet
wb = Workbook()
ws = wb.active
ws.title = "Images"

# Optional: widen column for images
ws.column_dimensions["A"].width = 25
ws.row_dimensions[1].height = 100  # adjust row height (for first row)

# Loop through image files
row = 1
for file in os.listdir(img_folder):
    if file.lower().endswith((".png", ".jpg", ".jpeg")):
        img_path = os.path.join(img_folder, file)

        # Create Image object
        img = Image(img_path)

        # Resize image (optional)
        img.width = 100
        img.height = 100

        # Insert into cell
        cell = f"A{row}"
        ws.add_image(img, cell)

        # Adjust row height for image
        ws.row_dimensions[row].height = 80

        row += 1

# Save Excel file
wb.save(output_file)
print(f"Saved Excel file with images: {output_file}")
