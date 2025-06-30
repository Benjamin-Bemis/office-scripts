# Requisition Form Automation Script
#
# This script reads an Excel file containing a list of items to purchase,
# groups the items by vendor, and generates a filled-out PDF requisition
# form for each vendor.

# --- 1. SETUP ---
#
# Before running, you need to install the necessary Python libraries.
# Open your terminal or command prompt and run the following command:
#
# pip install pandas openpyxl fillpdf
#

import pandas as pd
from fillpdf import fillpdfs
import os
from datetime import datetime

# --- 2. CONFIGURATION ---
#
# Update these variables with your specific information and file paths.

# Input files
# Make sure these files are in the same directory as the script,
# or provide the full path to them.
EXCEL_FILE_PATH = 'excel_sheets/tool_sheet.xlsx'
PDF_TEMPLATE_PATH = 'form/UPDATED AME Purchase Request Form.pdf'

# Output directory
# This is where the filled-out PDF forms will be saved.
OUTPUT_DIR = 'filled_requisitions'

# Static information for the form
# This information will be the same on every generated form.
REQUESTED_BY = 'Benjamin Bemis'
PHONE = '740-915-1405'
DELIVER_TO = 'Benjamin Bemis/White Field Research Laboratory/142'
PI_NAME = 'Thomas Juliano'

# --- IMPORTANT: PDF Field Mapping ---
#
# You MUST map the column headers from your Excel file to the exact field
# names in your PDF form.
#
# HOW TO FIND PDF FIELD NAMES:
# 1. Run this script once with `RUN_FIELD_FINDER = True` (at the bottom).
# 2. The script will print a dictionary of all available field names.
# 3. Use those names to update the field mappings below.
#
# Based on your request, the logic now handles field names that end with a
# row number (e.g., 'QuantityRow1', 'DescriptionRow1').
#
# When you get the output from the field finder, find the pattern for the item
# rows. For example, if you see fields named 'QuantityRow1', 'QuantityRow2',
# etc., then the "base name" is 'QuantityRow'.
#
# Update the keys of this dictionary with the BASE NAMES of your PDF fields.
# I have provided realistic examples below.
PDF_FIELD_MAP = {
    # PDF Field Base Name   : Excel Column Name
    'QuantityRow':          'Quantity',
    'ItemCatalog Row':       'Link/Part #',
    'Item Description Complete SpecificationsRow':   'Description',
    'Unit PriceRow':         'Unit Price',
    'Total Price Row ':        'Total Price',
    'Chemical Bio HazardRow': 'Bio Hazard',
}

# --- 3. HELPER FUNCTION TO FIND PDF FIELD NAMES ---

def get_pdf_field_names(pdf_path):
    """
    Reads a PDF and prints its fillable field names.
    """
    print("--- Finding PDF Field Names ---")
    if not os.path.exists(pdf_path):
        print(f"Error: PDF template not found at '{pdf_path}'")
        return
    try:
        fields = fillpdfs.get_form_fields(pdf_path)
        print("Found the following fields in your PDF:")
        print(fields)
        print("\nUpdate the 'PDF_FIELD_MAP' and the 'data_dict' in the script with these names.")
    except Exception as e:
        print(f"An error occurred while trying to read the PDF fields: {e}")
        print("Please ensure the PDF path is correct and it is not password protected.")
    print("--- End of Field Finder ---")

# --- 4. MAIN SCRIPT LOGIC ---

def main():
    """
    Main function to run the automation process.
    """
    # --- File and Directory Validation ---
    if not os.path.exists(EXCEL_FILE_PATH):
        print(f"Error: Excel file not found at '{EXCEL_FILE_PATH}'")
        return

    if not os.path.exists(PDF_TEMPLATE_PATH):
        print(f"Error: PDF template not found at '{PDF_TEMPLATE_PATH}'")
        return

    if not os.path.exists(OUTPUT_DIR):
        print(f"Creating output directory: '{OUTPUT_DIR}'")
        os.makedirs(OUTPUT_DIR)

    # --- Read and Process Data ---
    try:
        df = pd.read_excel(EXCEL_FILE_PATH, engine='openpyxl')
        required_cols = ['Vendor', 'Quantity', 'Link/Part #', 'Description', 'Unit Price']
        if not all(col in df.columns for col in required_cols):
            print(f"Error: The Excel file must contain the following columns: {required_cols}")
            return

        df['Total Price'] = df['Quantity'] * df['Unit Price']
        vendors = df['Vendor'].unique()
        print(f"Found {len(vendors)} unique vendors: {', '.join(vendors)}")

    except Exception as e:
        print(f"An error occurred while reading the Excel file: {e}")
        return

    # --- Generate PDF for Each Vendor ---
    for vendor in vendors:
        print(f"\nProcessing vendor: {vendor}")
        vendor_df = df[df['Vendor'] == vendor].reset_index(drop=True)

        # --- Create the data dictionary to fill the PDF ---
        # IMPORTANT: The keys here ('SupplierName', 'RequestedBy', etc.) are PDF field
        # names. You MUST update them to match the output of the field finder.
        data_dict = {
            'Supplier Name': vendor,
            'Requested By': REQUESTED_BY,
            'Phone': PHONE,
            'DATE': datetime.now().strftime('%m/%d/%y'),
            'Deliver to NameBuildingRoom': DELIVER_TO,
            'PI Name': PI_NAME,
            'Fund Number.': '', # Left blank as requested
        }

        # --- Populate item rows ---
        # This section now constructs field names like 'QuantityRow1', 'QuantityRow2',
        # based on the base names you define in PDF_FIELD_MAP.
        for i, row in vendor_df.iterrows():
            # PDF form rows are usually 1-indexed, while pandas iterrows is 0-indexed.
            row_num = i + 1
            for pdf_field_base, excel_col in PDF_FIELD_MAP.items():
                # Constructs the full field name, e.g., 'QuantityRow' + '1' -> 'QuantityRow1'
                field_name = f'{pdf_field_base}{row_num}'
                
                # Format the price columns as currency strings, otherwise just use the value
                if excel_col in ['Unit Price', 'Total Price']:
                    data_dict[field_name] = f'$ {row[excel_col]:.2f}'
                else:
                    data_dict[field_name] = row[excel_col]

        # Calculate and add the order total
        order_total = vendor_df['Total Price'].sum()
        # IMPORTANT: You must find the exact field name for the total in your PDF.
        data_dict['Order Total'] = f'$ {order_total:.2f}'

        # --- Fill and Save the PDF ---
        date = data_dict['DATE'] 
        output_filename = f"requisition_form_{vendor.replace(' ', '_').lower()}_{date.replace('/','.')}.pdf"
        output_path = os.path.join(OUTPUT_DIR, output_filename)

        try:
            fillpdfs.write_fillable_pdf(
                input_pdf_path=PDF_TEMPLATE_PATH,
                output_pdf_path=output_path,
                data_dict=data_dict,
                flatten=False # Set to True if you want to make the final PDF non-editable
            )
            print(f"Successfully created PDF: {output_path}")
        except Exception as e:
            print(f"\n--- ERROR ---")
            print(f"Error writing PDF for {vendor}: {e}")
            print("This is likely due to incorrect field names in 'PDF_FIELD_MAP' or 'data_dict'.")
            print("Please run the field finder (set RUN_FIELD_FINDER = True) to get the correct names.")
            print("-------------\n")


if __name__ == '__main__':
    # --- SCRIPT EXECUTION ---

    # STEP 1: Set this to True and run the script once to discover your PDF's field names.
    # STEP 2: Update the PDF_FIELD_MAP and data_dict keys with the correct names.
    # STEP 3: Set this back to False to run the main form-filling logic.
    RUN_FIELD_FINDER = False

    if RUN_FIELD_FINDER:
        get_pdf_field_names(PDF_TEMPLATE_PATH)
    else:
        main()
