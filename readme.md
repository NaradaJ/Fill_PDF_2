# PDF Data Extraction and Image to PDF Converter

This Python script allows you to extract data from an Excel file and generate PDF documents from an existing PDF template. It's useful for automating the creation of customized PDF reports from Excel data.

## Features

- Extracts specific columns from an Excel file.
- Identifies and handles merged cells in the Excel data.
- Maps the extracted data to predefined text boxes in a PDF template.
- Generates individual PDF reports for each row of data.

## Prerequisites

Before running the script, ensure you have the following:

- Python 3.x installed.
- Required Python packages installed (see the `requirements.txt` file).
- An Excel file (`excel_file_path`) containing the data you want to process.
- A PDF template file (`pdf_file_path`) with pre-defined text boxes where data will be inserted.

## Installation

1. Clone this repository to your local machine or download the code files.


2. Navigate to the project directory.


3. Install the required Python packages using pip.


## Usage

1. Update the `excel_file_path` variable with the path to your Excel file.
2. Update the `pdf_file_path` variable with the path to your PDF template.
3. Customize the `textbox_order` variable to match the order of text boxes in your PDF template.
4. Run the script.


5. The script will process the data, create individual PDF reports, and save them in the `output_directory` folder.

## Customization

- You can modify the `columns_to_keep` variable to select specific columns from your Excel data.
- Adjust the coordinates in the `textbox_coordinates` variable to match your PDF template's text box positions.

## Example

To generate a PDF report for a specific row of data, modify the `selected_row_index` variable and run the script. The generated PDF will be saved in the `output_directory`.

```python
# Select the row you want to generate a PDF for (e.g., row index 1)
selected_row_index = 1
selected_row_data = row_datasets[selected_row_index]

# Generate a PDF for the selected row
pdf_path = save_image_as_pdf(image, selected_row_data, output_directory, selected_row_index)

# Print the path to the generated PDF
print("PDF created successfully:", pdf_path)


5. The script will process the data, create individual PDF reports, and save them in the `output_directory` folder.

## Customization

- You can modify the `columns_to_keep` variable to select specific columns from your Excel data.
- Adjust the coordinates in the `textbox_coordinates` variable to match your PDF template's text box positions.

## Example

To generate a PDF report for a specific row of data, modify the `selected_row_index` variable and run the script. The generated PDF will be saved in the `output_directory`.

```python
# Select the row you want to generate a PDF for (e.g., row index 1)
selected_row_index = 1
selected_row_data = row_datasets[selected_row_index]

# Generate a PDF for the selected row
pdf_path = save_image_as_pdf(image, selected_row_data, output_directory, selected_row_index)

# Print the path to the generated PDF
print("PDF created successfully:", pdf_path)

