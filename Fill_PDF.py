import pandas as pd
import fitz  # PyMuPDF
import cv2
import os
import numpy as np
from io import BytesIO

# Specify the path to your Excel file
excel_file_path = r"C:\Users\NaradaJayasuriyaBIST\Desktop\Dev\Python 2\BBOX2\src\T10 test data.xlsx"

# Read the Excel file into a Pandas DataFrame, skipping the first row (header row)
df = pd.read_excel(excel_file_path, header=0)

# Remove rows with NaN values
df = df.dropna()

# Remove special characters from column headers
df.columns = df.columns.str.replace(r'[^\w\s]', '', regex=True)

# Identify merged cells (empty cells in this example)
merged_cells = df.applymap(lambda x: x == '')

# Clean up by replacing original merged cells with NaN
df = df.mask(merged_cells)

# Reset the index
df.reset_index(drop=True, inplace=True)

# Specify the indices of the columns to keep
columns_to_keep = [0, 1, 2, 3, 4,5,6,7,8,9] 

# Select the desired columns and exclude merged columns
desired_frame = df.iloc[:, columns_to_keep]

# Extract rows as separate datasets without headers
row_datasets = [row for row in desired_frame.values]

# Specify the path to the PDF file
pdf_file_path = r"C:\Users\NaradaJayasuriyaBIST\Desktop\Dev\Python 2\BBOX2\src\APIT_T10_2324_EST.pdf"

# Open the PDF and select the first page
pdf_document = fitz.open(pdf_file_path)
page = pdf_document.load_page(0)

# Get the original resolution of the page
original_resolution = page.get_pixmap().width, page.get_pixmap().height

# Print the original resolution
print("Original Resolution:", original_resolution)

# Convert the page to an image
pix = page.get_pixmap()
img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

# Load the image using OpenCV
image = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR)

# Define the order of text boxes and the corresponding row values
textbox_order = ["Employee Number", "Full Name", "NIC", "Total Gross Remuneration", "Cash Benefits",
                 "Deducted Tax1", "Deducted Tax2", "Remitted to the Inland Revenue Department"]

# Create a directory to store output PDFs
output_directory = "output_directory"
os.makedirs(output_directory, exist_ok=True)

# Define a function to save an image with text boxes as a PDF
def save_image_as_pdf(image, row_data, output_directory, index):
    # Create a unique filename for each PDF
    pdf_filename = f"output_{index}.pdf"
    pdf_path = os.path.join(output_directory, pdf_filename)

    # Create a new PDF document
    pdf_document = fitz.open()

    # Create a new page with the original resolution
    pdf_page = pdf_document.new_page(width=original_resolution[0], height=original_resolution[1])

    # Convert the OpenCV image to a PyMuPDF pixmap
    pdf_pix = fitz.Pixmap(fitz.csRGB, pdf_page.rect)
    pdf_pix.pil_dpi = (300, 300)
    
    # Convert the OpenCV image to a PIL image
    pil_image = cv2.cvtColor(image, cv2.COLOR_BGR2RGB)
    
    # Save the PIL image to a BytesIO object
    pil_image_io = BytesIO()
    Image.fromarray(pil_image).save(pil_image_io, 'JPEG')
    
    # Draw the image on the PDF page
    pdf_page.insert_image(pdf_page.rect, stream=pil_image_io.getvalue())

    # Define coordinates for the text boxes
    textbox_coordinates = [
        ((519, 668), (1407, 712)),
        ((634, 791), (915, 831)),
        ((1163, 798), (1478, 836)),
        ((822, 919), (1042, 952)),
        ((1208, 921), (1381, 952)),
        ((1106, 1016), (1538, 1054)),
        ((560, 1131), (1228, 1171)),
        ((1018, 1446), (1535, 1479)),
        ((880, 1517), (1542, 1543)),
        ((1048, 1596), (1541, 1636))
    ]

    # Define the order of text boxes and fill with corresponding row values
    for label, ((x1, y1), (x2, y2)), value in zip(textbox_order, textbox_coordinates, row_data):
        # Subtract 2 pixels from y1 and y2 to move the box 2 pixels up
        y1 -= 0
        y2 -= 0

        # Convert the value to a string
        if isinstance(value, float):
            value = f'{value:.2f}'
        else:
            value = str(value)

        # Add the value text to the PDF page
        pdf_page.insert_text((x1, y1), value, fontname="helv", fontsize=10)

    # Save the PDF document
    pdf_document.save(pdf_path)

    # Close the PDF document
    pdf_document.close()

    return pdf_path

# Select the row you want to generate a PDF for (e.g., row index 1)
selected_row_index = 1
selected_row_data = row_datasets[selected_row_index]

# Generate a PDF for the selected row
pdf_path = save_image_as_pdf(image, selected_row_data, output_directory, selected_row_index)

# Print the path to the generated PDF
print("PDF created successfully:", pdf_path)
