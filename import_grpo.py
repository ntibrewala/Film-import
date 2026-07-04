import pandas as pd
import requests
import json
import urllib3
import tkinter as tk
from tkinter import filedialog
import os

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Configuration from Environment Variables
SL_URL = os.environ.get('SAP_SL_URL', 'https://10.10.0.113:50000/b1s/v1')
SL_USER = os.environ.get('SAP_USER', 'manager')
SL_PASSWORD = os.environ.get('SAP_PASSWORD', 'bppl@123')
SL_DB = os.environ.get('SAP_COMPANY_DB', 'TEST_BHAVYA_23062026')
def load_vendor_mappings():
    try:
        with open('vendor_mappings.json', 'r') as f:
            return json.load(f)
    except FileNotFoundError:
        print("Error: vendor_mappings.json not found.")
        return {}

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

item_code_cache = {}
def get_item_code_by_bp_catalog(session, bp_code, catalog_no):
    """Lookup SAP ItemCode by BP Catalog Number"""
    cache_key = f"{bp_code}_{catalog_no}"
    if cache_key in item_code_cache:
        return item_code_cache[cache_key]
    
    query = f"{SL_URL}/AlternateCatNum?$select=ItemCode&$filter=CardCode eq '{bp_code}' and Substitute eq '{catalog_no}'"
    res_obj = session.get(query, verify=False)
    res = res_obj.json()
    if res.get('value') and len(res['value']) > 0:
        item_code = res['value'][0]['ItemCode']
        item_code_cache[cache_key] = item_code
        return item_code
    else:
        print(f"Debug: BP Catalog query returned HTTP {res_obj.status_code}: {res}")
    return None

def main(vendor_code, invoice_file, packing_file, dry_run=False):
    mappings = load_vendor_mappings()
    if vendor_code not in mappings:
        print(f"Error: No mapping found for vendor {vendor_code} in vendor_mappings.json")
        return
    v_map = mappings[vendor_code]
    inv_map = v_map["invoice"]
    pack_map = v_map["packing"]
    print(f"Reading invoice data from: {invoice_file}")
    inv_df = pd.read_excel(invoice_file)
    
    print(f"Reading packing slip data from: {packing_file}")
    pack_df = pd.read_excel(packing_file)
    
    # Clean up column names for safer access
    inv_df.columns = [str(c).strip() for c in inv_df.columns]
    pack_df.columns = [str(c).strip() for c in pack_df.columns]
    
    session = login()
    
    # Group by Invoice Number
    invoice_col = inv_map.get("invoice_num", "Invoice")
    if invoice_col in inv_df.columns:
        invoices = inv_df.groupby(invoice_col)
    else:
        print(f"Error: '{invoice_col}' column not found in invoice file.")
        return

    for inv_num, inv_group in invoices:
        print(f"\n--- Processing Invoice: {inv_num} ---")
        header = inv_group.iloc[0]
        
        customer_po = header.get('Customer PO No.', '')
        order_no = header.get('Sale Order', '')
        billing_date = header.get(inv_map.get('date', 'Billing Date'))
        
        # Format date for SAP (YYYY-MM-DD)
        try:
            doc_date = pd.to_datetime(billing_date).strftime('%Y-%m-%d')
        except:
            doc_date = None
            print(f"Warning: Could not parse date '{billing_date}'")
            
        po_data = find_po(session, customer_po, order_no)
        if not po_data:
            print(f"Error: PO not found for Invoice {inv_num} (Customer PO: {customer_po}, Order No: {order_no})")
            continue
            
        po_docentry = po_data['DocEntry']
        po_cardcode = po_data['CardCode']
        po_cardname = po_data['CardName']
        print(f"Found linked PO DocNum {po_data['DocNum']} (DocEntry {po_docentry}) from Vendor {po_cardcode}")
        
        grpo_payload = {
            "CardCode": vendor_code, 
            "NumAtCard": str(inv_num),
            "DocumentLines": []
        }
        if doc_date:
            grpo_payload["DocDate"] = doc_date
            
        used_po_lines = set()
        
        # Add Lines
        for idx, inv_line in inv_group.iterrows():
            material = str(inv_line.get(inv_map.get('material', 'Material'), '')).strip()
            
            try:
                net_value = float(inv_line.get(inv_map.get('net_value', 'Net Value'), 0))
                total_qty = float(inv_line.get(inv_map.get('qty', 'Total Qty.'), 0))
            except ValueError:
                net_value = 0.0
                total_qty = 0.0
                
            # Get SAP ItemCode from BP Catalog
            sap_item_code = get_item_code_by_bp_catalog(session, vendor_code, material)
            
            if not sap_item_code:
                print(f"Warning: Could not find SAP ItemCode for BP Catalog No '{material}'. Skipping line.")
                continue
                
            # Find matching line in PO
            base_line_num = None
            for pl in po_data.get('DocumentLines', []):
                if str(pl['ItemCode']).strip() == str(sap_item_code).strip():
                    if pl['LineNum'] not in used_po_lines:
                        base_line_num = pl['LineNum']
                        used_po_lines.add(pl['LineNum'])
                        break
            
            if base_line_num is None:
                available_items = [f"{pl.get('ItemCode')} (Line {pl.get('LineNum')})" for pl in po_data.get('DocumentLines', [])]
                print(f"Warning: Item '{sap_item_code}' (FrgnName: '{material}') could not be linked to an unused PO line in PO {po_data['DocNum']}. Available: {available_items}. It will be added without BaseLine linking.")
            
            price = net_value / total_qty if total_qty > 0 else 0
            
            doc_line = {
                "ItemCode": sap_item_code,
                "Quantity": total_qty,
                "UnitPrice": price,
                "DiscountPercent": 0.0,
                "BatchNumbers": []
            }
            
            if base_line_num is not None:
                doc_line["BaseType"] = 22 # Purchase Order
                doc_line["BaseEntry"] = po_docentry
                doc_line["BaseLine"] = base_line_num
            
            # Invoice width handling
            try:
                inv_width_val = float(str(inv_line.get(inv_map.get('width', 'Width'), '')).lower().replace('mm', '').strip())
            except ValueError:
                inv_width_val = -1.0
                
            # Find batches from packing slip
            # Linking packing slip using 'Billing Document' == Invoice Number
            billing_col = pack_map.get('invoice_num', 'Billing Document')
            if billing_col not in pack_df.columns:
                # Try to find case-insensitive match
                matches = [c for c in pack_df.columns if str(c).lower().replace(' ', '') == 'billingdocument']
                if matches:
                    billing_col = matches[0]
                else:
                    print(f"Error: Column '{billing_col}' not found in packing slip. Available columns are: {list(pack_df.columns)}")
                    return
            
            material_col = pack_map.get('material', 'Material')
            if material_col not in pack_df.columns:
                matches = [c for c in pack_df.columns if str(c).lower().replace(' ', '') == 'material']
                if matches:
                    material_col = matches[0]
                else:
                    print(f"Error: Column '{material_col}' not found in packing slip. Available columns are: {list(pack_df.columns)}")
                    return
                    
            width_col = pack_map.get('width', 'Width MM')
            if width_col not in pack_df.columns:
                matches = [c for c in pack_df.columns if str(c).lower().replace(' ', '') == 'widthmm']
                if matches:
                    width_col = matches[0]
                else:
                    print(f"Warning: Column '{width_col}' not found in packing slip. Batches might be duplicated across widths!")
                    width_col = None

            if width_col and inv_width_val != -1.0:
                def match_width(w):
                    try:
                        return float(str(w).strip()) == inv_width_val
                    except:
                        return False
                width_mask = pack_df[width_col].apply(match_width)
            else:
                width_mask = True

            batch_rows = pack_df[(pack_df[billing_col].astype(str).str.strip() == str(inv_num).strip()) & 
                                 (pack_df[material_col].astype(str).str.strip() == str(material).strip()) &
                                 width_mask]
            
            for _, batch_row in batch_rows.iterrows():
                # Fetch weights safely
                net_wt_col = pack_map.get('net_wt', 'Net_WT(KG)')
                net_wt = batch_row.get(net_wt_col, batch_row.get('Net_WT(KGS)', 0))
                
                batch = {
                    "BatchNumber": str(batch_row.get(pack_map.get('batch_num', 'Roll No'), '')),
                    "Quantity": float(net_wt) if pd.notnull(net_wt) else 0.0,
                }
                
                udfs = pack_map.get('batch_udfs', {})
                for udf_name, col_name in udfs.items():
                    batch[udf_name] = str(batch_row.get(col_name, ''))
                
                base_price = doc_line.get("UnitPrice", 0)
                batch.update({
                    "U_Price": base_price,
                    "U_CardCode": po_cardcode,
                    "U_CardName": po_cardname,
                    "U_NetWt": str(net_wt),
                    "U_SalesPrice": base_price
                })
                
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
        
    print("Enter the Vendor Code (e.g. V00038):")
    vendor_code = input().strip()
    if not vendor_code:
        print("Vendor code cannot be empty. Exiting.")
        exit(0)

    print("Do you want to run in dry-run mode to preview payload? (y/n)")
    ans = input().strip().lower()
    is_dry_run = (ans == 'y')
    
    main(vendor_code, invoice_file, packing_file, dry_run=is_dry_run)
