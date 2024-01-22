import pandas as pd
from PyPDF2 import PdfReader

pdf_path = r"Data-Extraction/ICICI bank.pdf"

def read_pdf(pdf_path, page_number=0):
    with open(pdf_path, 'rb') as file:
        reader = PdfReader(file)
        # Check if the requested page number is within the valid range
        if 0 <= page_number < len(reader.pages):
            page = reader.pages[page_number]
            text = page.extract_text()
            return text
        else:
            return "Invalid page number"

# Read content from the first (0-indexed) page
pdf_content = read_pdf(pdf_path, page_number=0)

# Define the titles and their corresponding slicing indices
title_indices = {
    "Company Name": (0, 17),
    "Account No": (91, 104),
    "Customer Ref. No.": (124, 134),
    "Value Date": (147, 160),
    "UTR No.": (494, 508),
    "Beneficiary Name": (249, 284),
    "Beneficiary Code": (178, 189),
    "Document No.": (964, 974),
    "Invoice No.": (975, 988),
    "Invoice Date": (988, 997),
    "Gross Amount": (997, 1005),
    "Net Amount": (997, 1005),
    "Total Value": (997, 1005),
}

# Initialize a dictionary to store the extracted values
extracted_data = {}

# Iterate over the titles and extract corresponding values using the function
for title, indices in title_indices.items():
    start_index, end_index = indices
    value = pdf_content[start_index:end_index].strip()
    extracted_data[title.strip()] = value

# Convert the data to a DataFrame
df = pd.DataFrame([extracted_data])

# Save the DataFrame to an Excel file
excel_output_path = r"Data-Extraction/output_data.xlsx"
df.to_excel(excel_output_path, index=False)

print(f"Data saved to {excel_output_path}")
