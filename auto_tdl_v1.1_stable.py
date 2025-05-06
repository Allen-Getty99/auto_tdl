import pdfplumber
import re
import os
import pandas as pd


filename = input("Enter invoice filename: ")
PDF_FILE_PATH = "FY25 P8 5057820314.pdf"
EXCEL_FILE_PATH = "TDL_DATABASE.xlsx"

# Python 3 compatibility note:
# This script is written for Python 3 and requires these packages:
# - pandas: for Excel processing
# - openpyxl: for Excel file reading (used by pandas)
# - pdfplumber: for PDF processing
#
# Setup instructions:
# 1. Navigate to the project directory:
#    cd /Users/allengettyliquigan/Downloads/Project_Auto_GFS
#
# 2. Create a virtual environment:
#    python3 -m venv tdl_env
#
# 3. Activate the virtual environment:
#    - On Mac/Linux: source tdl_env/bin/activate
#    - On Windows: tdl_env\Scripts\activate
#
# 4. Install required packages:
#    pip install pandas openpyxl pdfplumber
#
# 5. Run the script:
#    python3 auto_tdl_v1.0_stable.py

def process_tdl_invoice(pdf_file_path, excel_file_path):
    print(f"=== Processing Tim Hortons Invoice ===")
    print(f"Loading database: {excel_file_path}")
    
    try:
        db = pd.read_excel(excel_file_path, usecols=["Item Code", "GL Code", "GL Description"])
        db["Item Code"] = db["Item Code"].astype(str)
        print("✅ Database loaded successfully.")
    except Exception as e:
        print("❌ Failed to load database:", e)
        return

    # Placeholder for rest of invoice processing logic
    print("Processing PDF:", pdf_file_path)
    """
    Process a Tim Hortons invoice PDF and extract required information
    """
    print("=== Processing Tim Hortons Invoice ===")
    
    # Check if files exist
    if not os.path.exists(pdf_file_path):
        print(f"Error: PDF file not found: {pdf_file_path}")
        return
    
    if not os.path.exists(excel_file_path):
        print(f"Error: Excel database file not found: {excel_file_path}")
        return
    
    # Load database
    print(f"Loading database: {excel_file_path}")
    db = pd.read_excel(excel_file_path)
    
    # Ensure database has required columns
    required_cols = ["Item Code", "GL Code", "GL Description"]
    missing_cols = [col for col in required_cols if col not in db.columns]
    if missing_cols:
        print(f"Error: Missing required columns in database: {', '.join(missing_cols)}")
        return
    
    # Convert Item Code to string for matching
    db["Item Code"] = db["Item Code"].astype(str)
    
    # Extract data from PDF
    items = []
    invoice_number = None
    tariff_amount = 0.0
    fuel_surcharge = 0.0
    gst_hst_vat = 0.0
    
    try:
        with pdfplumber.open(pdf_file_path) as pdf:
            print(f"PDF opened successfully with {len(pdf.pages)} pages")
            
            # Process each page
            for page_num, page in enumerate(pdf.pages):
                print(f"Processing page {page_num + 1}/{len(pdf.pages)}")
                text = page.extract_text()
                
                # Extract invoice number (only from first page)
                if page_num == 0 and not invoice_number:
                    invoice_match = re.search(r'Invoice Number\s*:\s*(\d+)', text)
                    if invoice_match:
                        invoice_number = invoice_match.group(1)
                        print(f"Invoice Number: {invoice_number}")
                
                # Process line items on all pages
                lines = text.split('\n')
                for line in lines:
                    # Look for item codes (8 digits or occasionally 5 digits at start of line)
                    item_code_match = re.match(r'^(\d{5,8})\s', line)
                    if item_code_match:
                        item_code = item_code_match.group(1)
                        # Pad to 8 digits if shorter
                        if len(item_code) < 8:
                            item_code = item_code.zfill(8)
                        
                        # Extract quantities and price
                        parts = line.split()
                        
                        # Extract numerical values (find positions of numbers)
                        numbers = []
                        for part in parts:
                            if re.match(r'^\d+\.?\d*$', part) or re.match(r'^\d+$', part):
                                numbers.append(part)
                        
                        # Skip if we don't have enough numbers
                        if len(numbers) < 4:
                            continue
                        
                        # Tim Hortons invoice format has these numbers in specific positions
                        try:
                            qty = int(float(numbers[1]))  # Shipped qty is typically the 2nd number
                            unit_price = float(numbers[3])  # Unit price is typically the 4th number
                            line_total = round(qty * unit_price, 2)
                            
                            # Look up GL code and description
                            gl_code = "NOT_FOUND"
                            gl_desc = "NOT_FOUND"
                            
                            match = db[db["Item Code"] == item_code]
                            if not match.empty:
                                gl_code = match.iloc[0]["GL Code"]
                                gl_desc = match.iloc[0]["GL Description"]
                            
                            # Add to items list
                            items.append({
                                "Item Code": item_code,
                                "Quantity": qty,
                                "Unit Price": unit_price,
                                "Line Total": line_total,
                                "GL Code": gl_code,
                                "GL Description": gl_desc
                            })
                        except (IndexError, ValueError):
                            pass
                
                # Look for additional charges on last page
                if page_num == len(pdf.pages) - 1:
                    # Extract tariff amount
                    tariff_match = re.search(r'Tariff Allocation\s+(\d+\.\d+)', text)
                    if tariff_match:
                        tariff_amount = float(tariff_match.group(1))
                    
                    # Extract fuel surcharge
                    fuel_match = re.search(r'Fuel Surcharge\s+\d+\.\d+\s+0\.00\s+(\d+\.\d+)', text)
                    if fuel_match:
                        fuel_surcharge = float(fuel_match.group(1))
                    else:
                        # Try alternate format
                        fuel_match = re.search(r'Fuel Surcharge\s+(\d+\.\d+)', text)
                        if fuel_match:
                            fuel_surcharge = float(fuel_match.group(1))
                    
                    # Extract GST/HST/VAT
                    gst_match = re.search(r'GST/HST/VAT\s+(\d+\.\d+)', text)
                    if gst_match:
                        gst_hst_vat = float(gst_match.group(1))
    
    except Exception as e:
        print(f"Error processing PDF: {e}")
        return
    
    # Print extracted items in table format
    print("\n" + "=" * 100)
    print(f"{'ITEM CODE':<12} {'QTY':<6} {'UNIT PRICE':<12} {'LINE TOTAL':<12} {'GL CODE':<10} {'GL DESCRIPTION':<30}")
    print("=" * 100)
    
    for item in items:
        print(f"{item['Item Code']:<12} {item['Quantity']:<6} ${item['Unit Price']:<10.2f} ${item['Line Total']:<10.2f} {str(item['GL Code']):<10} {item['GL Description']:<30}")
    
    # Print additional charges
    print("\nAdditional Charges:")
    print(f"Tariff Amount: ${tariff_amount:.2f}")
    print(f"Fuel Surcharge: ${fuel_surcharge:.2f}")
    print(f"GST/HST/VAT: ${gst_hst_vat:.2f}")
    
    # Calculate summary by GL Description
    summary = {}
    for item in items:
        gl_desc = item['GL Description']
        summary[gl_desc] = summary.get(gl_desc, 0) + item['Line Total']
    
    # Print summary
    print("\nSummary by GL Description:")
    print("=" * 50)
    total_amount = 0
    for gl_desc, amount in sorted(summary.items(), key=lambda x: x[1], reverse=True):
        print(f"{gl_desc}: ${amount:.2f}")
        total_amount += amount
    
    print("=" * 50)
    print(f"Total Amount: ${total_amount:.2f}")
    print(f"Additional Charges: ${tariff_amount + fuel_surcharge + gst_hst_vat:.2f}")
    print(f"Grand Total: ${total_amount + tariff_amount + fuel_surcharge + gst_hst_vat:.2f}")

def main():
    process_tdl_invoice(PDF_FILE_PATH, EXCEL_FILE_PATH)

if __name__ == "__main__":
    main()