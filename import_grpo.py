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
from pathlib import Path
import os

CONFIG_DIR = Path(r"D:\bhv\importgrpo")
try:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
except Exception:
    pass

CONFIG_PATH = CONFIG_DIR / "sap_vendor_mappings.json"

def load_vendor_mappings():
    # Migrate local file to global path if it exists
    if os.path.exists('vendor_mappings.json') and not CONFIG_PATH.exists():
        try:
            with open('vendor_mappings.json', 'r') as f:
                data = json.load(f)
            with open(CONFIG_PATH, 'w') as f:
                json.dump(data, f, indent=2)
        except Exception:
            pass
            
    try:
        with open(CONFIG_PATH, 'r') as f:
            return json.load(f)
    except FileNotFoundError:
        return {}

def save_vendor_mappings(mappings):
    with open(CONFIG_PATH, 'w') as f:
        json.dump(mappings, f, indent=2)

def prompt_col(message):
    if message.endswith(": "):
        message = message[:-2]
    try:
        val = input(f"{message} [press Ctrl+E to exit]: ").strip()
        # \x05 is the ASCII character for Ctrl+E
        if val == '\x05' or val.lower() in ['exit', 'quit', 'esc', 'e']:
            print("\nMapping aborted by user. No changes were saved.")
            import sys
            sys.exit(0)
        return val
    except KeyboardInterrupt:
        print("\nMapping aborted by user. No changes were saved.")
        import sys
        sys.exit(0)

def setup_new_vendor(vendor_code, mappings):
    print(f"\n--- Interactive Setup for New Vendor: {vendor_code} ---")
    print("Please type the exact column names as they appear in the Excel files.")
    print("If a field is not present in the Excel file, you can just press Enter to skip it.")
    print("Type 'exit' or 'esc' at any prompt to cancel the setup without saving.\n")
    
    invoice = {}
    packing = {}
    batch_udfs = {}
    
    print("[INVOICE EXCEL FILE]")
    invoice["invoice_num"] = prompt_col("Column name for 'Invoice Number' (e.g. Invoice): ")
    invoice["material"] = prompt_col("Column name for 'Material / Item Part Number': ")
    invoice["qty"] = prompt_col("Column name for 'Total Quantity': ")
    invoice["net_value"] = prompt_col("Column name for 'Line Net Value (Price)': ")
    invoice["width"] = prompt_col("Column name for 'Roll Width' (if any): ")
    invoice["date"] = prompt_col("Column name for 'Billing Date' (if any): ")
    
    print("\n[PACKING SLIP EXCEL FILE]")
    packing["invoice_num"] = prompt_col("Column name for 'Invoice / Billing Document': ")
    packing["material"] = prompt_col("Column name for 'Material / Item Part Number': ")
    packing["width"] = prompt_col("Column name for 'Width MM': ")
    packing["net_wt"] = prompt_col("Column name for 'Net Weight (KGS)': ")
    packing["batch_num"] = prompt_col("Column name for 'Batch Number / Roll No': ")
    
    print("\n[PACKING SLIP - BATCH DETAILS (UDFs)]")
    batch_udfs["U_Length"] = prompt_col("Column name for 'Length (MTR)': ")
    batch_udfs["U_width"] = prompt_col("Column name for 'Width MM': ")
    batch_udfs["U_EmName"] = prompt_col("Column name for 'EmName' (if any): ")
    batch_udfs["U_VechileNo"] = prompt_col("Column name for 'Vehicle No': ")
    batch_udfs["U_Grade"] = prompt_col("Column name for 'Grade': ")
    batch_udfs["U_OD"] = prompt_col("Column name for 'OD MM': ")
    batch_udfs["U_Core"] = prompt_col("Column name for 'Core INCH': ")
    batch_udfs["U_Micron"] = prompt_col("Column name for 'Micron (MIC)': ")
    batch_udfs["U_GrossWt"] = prompt_col("Column name for 'Gross Weight (KGS)': ")
    
    # Remove empty answers
    invoice = {k: v for k, v in invoice.items() if v}
    packing = {k: v for k, v in packing.items() if v}
    batch_udfs = {k: v for k, v in batch_udfs.items() if v}
    
    packing["batch_udfs"] = batch_udfs
    mappings[vendor_code] = {
        "invoice": invoice,
        "packing": packing
    }
    
    save_vendor_mappings(mappings)
    print(f"\nSuccess! Configuration for {vendor_code} has been saved to vendor_mappings.json.")
    print("Continuing with script execution...\n")
    return mappings

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
    """Lookup SAP ItemCode by BP Catalog Number, with fallback to Foreign Name"""
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
        
    # Fallback to ForeignName if BP Catalog mapping isn't found
    fallback_query = f"{SL_URL}/Items?$select=ItemCode&$filter=ForeignName eq '{catalog_no}'"
    fb_res_obj = session.get(fallback_query, verify=False)
    fb_res = fb_res_obj.json()
    if fb_res.get('value') and len(fb_res['value']) > 0:
        item_code = fb_res['value'][0]['ItemCode']
        item_code_cache[cache_key] = item_code
        return item_code

    print(f"Debug: Item '{catalog_no}' not found in BP Catalog or ForeignName.")
    return None

def main(vendor_code, invoice_file, packing_file, dry_run=False):
    mappings = load_vendor_mappings()
    if vendor_code not in mappings:
        print(f"\nVendor '{vendor_code}' is not mapped yet.")
        ans = input("Would you like to set up the Excel mapping for this vendor now? (y/n): ").strip().lower()
        if ans == 'y':
            mappings = setup_new_vendor(vendor_code, mappings)
        else:
            print("Exiting.")
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
    import sys
    
    # Pre-flight validation: Check if all items in invoice are mapped in SAP BP Catalog
    print("\nPerforming pre-flight item mapping validation...")
    missing_mappings = []
    material_col = inv_map.get('material', 'Material')
    if material_col in inv_df.columns:
        unique_materials = inv_df[material_col].dropna().unique()
        for mat in unique_materials:
            mat_str = str(mat).strip()
            if not mat_str or str(mat).lower() == 'nan':
                continue
            sap_code = get_item_code_by_bp_catalog(session, vendor_code, mat_str)
            if not sap_code:
                missing_mappings.append(mat_str)
    
    if missing_mappings:
        print("\n" + "="*80)
        print("ERROR: PRE-FLIGHT VALIDATION FAILED")
        print("The following vendor part numbers are missing in SAP Business Partner Catalog Numbers:")
        for mm in missing_mappings:
            print(f"  - {mm}")
        print(f"\nPlease map these items in SAP for Vendor {vendor_code} (Inventory > Item Management > Business Partner Catalog Numbers)")
        print("="*80 + "\n")
        sys.exit(1)
        
    print("Validation passed. All items are mapped correctly.\n")
    
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
    vendor_code = input().strip().upper()
    if not vendor_code:
        print("Vendor code cannot be empty. Exiting.")
        exit(0)

    print("Do you want to run in dry-run mode to preview payload? (y/n)")
    ans = input().strip().lower()
    is_dry_run = (ans == 'y')
    
    main(vendor_code, invoice_file, packing_file, dry_run=is_dry_run)
