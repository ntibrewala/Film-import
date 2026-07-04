import pandas as pd
import requests
import json
import urllib3
import tkinter as tk
from tkinter import filedialog
import os

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Configuration from Environment Variables
SL_URL = os.environ.get('SAP_SL_URL', 'https://localhost:50000/b1s/v1')
SL_USER = os.environ.get('SAP_USER', 'manager')
SL_PASSWORD = os.environ.get('SAP_PASSWORD', 'password')
SL_DB = os.environ.get('SAP_COMPANY_DB', 'TEST_BHAVYA_23062026')
VENDOR_CODE = 'V00038'  # Hardcoded based on user request

def login():
    """Authenticate with SAP Service Layer"""
    print(f"Connecting to SAP Service Layer at {SL_URL} (DB: {SL_DB})...")
    session = requests.Session()
    login_url = f"{SL_URL}/Login"
    payload = {
        "CompanyDB": SL_DB,
        "UserName": SL_USER,
        "Password": SL_PASSWORD
    }
    response = session.post(login_url, json=payload, verify=False)
    if response.status_code == 200:
        print("Login successful.")
        return session
    else:
        raise Exception(f"Login failed: {response.text}")

def find_po(session, customer_po_no, order_no):
    """Find Purchase Order by NumAtCard or DocNum"""
    # Try by NumAtCard first
    if customer_po_no:
        query = f"{SL_URL}/PurchaseOrders?$filter=NumAtCard eq '{customer_po_no}'"
        res = session.get(query, verify=False).json()
        if res.get('value') and len(res['value']) > 0:
            return res['value'][0]
    
    # Try by DocNum if NumAtCard didn't yield results
    if order_no:
        try:
            order_no_int = int(order_no)
            query = f"{SL_URL}/PurchaseOrders?$filter=DocNum eq {order_no_int}"
            res = session.get(query, verify=False).json()
            if res.get('value') and len(res['value']) > 0:
                return res['value'][0]
        except ValueError:
            pass
            
    return None

def main(invoice_file, packing_file, dry_run=False):
    print(f"Reading invoice data from: {invoice_file}")
    inv_df = pd.read_excel(invoice_file)
    
    print(f"Reading packing slip data from: {packing_file}")
    pack_df = pd.read_excel(packing_file)
    
    # Clean up column names for safer access
    inv_df.columns = [str(c).strip() for c in inv_df.columns]
    pack_df.columns = [str(c).strip() for c in pack_df.columns]
    
    session = login()
    
    # Group by Invoice Number
    if 'Invoice' in inv_df.columns:
        invoices = inv_df.groupby('Invoice')
    else:
        print("Error: 'Invoice' column not found in invoice file.")
        return

    for inv_num, inv_group in invoices:
        print(f"\n--- Processing Invoice: {inv_num} ---")
        header = inv_group.iloc[0]
        
        customer_po = header.get('Customer PO No.', '')
        order_no = header.get('Sale Order', '')
        billing_date = header.get('Billing Date')
        
        # Format date for SAP (YYYY-MM-DD)
        try:
            doc_date = pd.to_datetime(billing_date).strftime('%Y-%m-%d')
        except:
            doc_date = None
            print(f"Warning: Could not parse Billing Date '{billing_date}'")
            
        po_data = find_po(session, customer_po, order_no)
        if not po_data:
            print(f"Error: PO not found for Invoice {inv_num} (Customer PO: {customer_po}, Order No: {order_no})")
            continue
            
        po_docentry = po_data['DocEntry']
        po_cardcode = po_data['CardCode']
        po_cardname = po_data['CardName']
        print(f"Found linked PO DocNum {po_data['DocNum']} (DocEntry {po_docentry}) from Vendor {po_cardcode}")
        
        grpo_payload = {
            "CardCode": VENDOR_CODE, 
            "NumAtCard": str(inv_num),
            "DocumentLines": []
        }
        if doc_date:
            grpo_payload["DocDate"] = doc_date
            
        # Add Lines
        for idx, inv_line in inv_group.iterrows():
            material = str(inv_line.get('Material', '')).strip()
            
            try:
                net_value = float(inv_line.get('Net Value', 0))
                total_qty = float(inv_line.get('Total Qty.', 0))
            except ValueError:
                net_value = 0.0
                total_qty = 0.0
                
            # Find matching line in PO
            base_line_num = None
            for pl in po_data.get('DocumentLines', []):
                if str(pl['ItemCode']).strip() == material:
                    base_line_num = pl['LineNum']
                    break
            
            if base_line_num is None:
                print(f"Warning: Item {material} not found in linked PO {po_data['DocNum']}. It will be added without BaseLine linking.")
            
            doc_line = {
                "ItemCode": material,
                "Quantity": total_qty,
                "Price": net_value / total_qty if total_qty > 0 else 0,
                "BatchNumbers": []
            }
            if base_line_num is not None:
                doc_line["BaseType"] = 22 # Purchase Order
                doc_line["BaseEntry"] = po_docentry
                doc_line["BaseLine"] = base_line_num
            
            # Find batches from packing slip
            # Linking packing slip using 'Billing Document' == Invoice Number
            batch_rows = pack_df[(pack_df['Billing Document'].astype(str) == str(inv_num)) & 
                                 (pack_df['Material'].astype(str) == material)]
            
            for _, batch_row in batch_rows.iterrows():
                price = doc_line["Price"]
                
                # Fetch weights safely
                net_wt = batch_row.get('Net_WT(KG)', batch_row.get('Net_WT(KGS)', 0))
                gross_wt = batch_row.get('Gross_WT(KGS)', batch_row.get('Gross_WT(KG)', 0))
                
                batch = {
                    "BatchNumber": str(batch_row.get('Roll No', '')),
                    "Quantity": float(net_wt) if pd.notnull(net_wt) else 0.0,
                    
                    # Batch UDFs mapping
                    "U_Length": str(batch_row.get('Length MTR', '')),
                    "U_width": str(batch_row.get('Width MM', '')),
                    "U_VechileNo": str(batch_row.get('Vehicle no', '')),
                    "U_Grade": str(batch_row.get('Grade', '')),
                    "U_OD": str(batch_row.get('OD MM', '')),
                    "U_Core": str(batch_row.get('Core INCH', '')),
                    "U_Price": price,
                    "U_CardCode": po_cardcode,
                    "U_CardName": po_cardname,
                    "U_Micron": str(batch_row.get('MIC', '')),
                    "U_NetWt": str(net_wt),
                    "U_GrossWt": str(gross_wt),
                    "U_SalesPrice": price
                }
                
                # Replace 'nan' string values with empty strings
                for k, v in batch.items():
                    if isinstance(v, str) and v.lower() == 'nan':
                        batch[k] = ''
                        
                doc_line["BatchNumbers"].append(batch)
                
            grpo_payload["DocumentLines"].append(doc_line)
            
        print(f"\nPayload prepared for Invoice {inv_num}:")
        if dry_run:
            print(json.dumps(grpo_payload, indent=2))
            print("[DRY RUN] Request not sent to SAP.")
        else:
            post_url = f"{SL_URL}/PurchaseDeliveryNotes"
            res = session.post(post_url, json=grpo_payload, verify=False)
            if res.status_code in (200, 201):
                print(f"Success! GRPO created with DocNum: {res.json().get('DocNum')}")
            else:
                print(f"Error creating GRPO: {res.text}")

if __name__ == '__main__':
    # Initialize tkinter and hide the root window
    root = tk.Tk()
    root.withdraw()
    
    print("Please select the Invoice Excel file...")
    invoice_file = filedialog.askopenfilename(
        title="Select Invoice Excel File",
        filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
    )
    if not invoice_file:
        print("No Invoice file selected. Exiting.")
        exit(0)
        
    print("Please select the Packing Slip Excel file...")
    packing_file = filedialog.askopenfilename(
        title="Select Packing Slip Excel File",
        filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
    )
    if not packing_file:
        print("No Packing Slip file selected. Exiting.")
        exit(0)
        
    print("Do you want to run in dry-run mode to preview payload? (y/n)")
    ans = input().strip().lower()
    is_dry_run = (ans == 'y')
    
    main(invoice_file, packing_file, dry_run=is_dry_run)
