import re
import os
from datetime import datetime
from typing import List, Dict, Optional
import pdfplumber
import pytesseract
from excel_handler_format2 import append_excel_format2

def pdf_to_text(pdf_path):
    '''Trích xuất text từ PDF (backup method)'''
    text = ""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
    except:
        pass
    
    if len(text.strip()) < 50:
        try:
            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    img = page.to_image(resolution=300).original
                    text += pytesseract.image_to_string(img, lang="eng") + "\n"
        except:
            pass
    
    return text

def clean_value(val):
    '''Làm sạch giá trị từ bảng'''
    if val is None:
        return ""
    val_str = str(val).strip()
    # Remove excessive whitespace
    val_str = re.sub(r'\s+', ' ', val_str)
    return val_str

def to_decimal(val, is_currency=False):
    '''
    Convert giá trị sang số decimal, xử lý thông minh:
    - is_currency=True: Luôn coi dấu chấm là phân cách hàng nghìn
      VD: 70.000 → 70000, 2.100.006 → 2100006
    - is_currency=False: Giữ số thập phân
      VD: 1.5 → 1.5, 10 → 10, 70.000 → 70000
    '''
    if val is None or val == "":
        return 0
    
    try:
        # Remove commas and spaces first
        cleaned = str(val).replace(",", "").replace(" ", "").strip()
        
        # Count dots in the string
        dot_count = cleaned.count(".")
        
        if dot_count == 0:
            # No dots, simple conversion
            return float(cleaned)
        
        elif dot_count == 1:
            # One dot - check context
            parts = cleaned.split(".")
            
            if is_currency:
                # For currency: always treat dot as thousand separator
                # 70.000 → 70000
                return float(cleaned.replace(".", ""))
            else:
                # For quantities: check if it's thousand separator or decimal
                # If after dot has exactly 3 digits → thousand separator
                # Otherwise → decimal point
                if len(parts[1]) == 3 and parts[1].isdigit() and int(parts[1]) == 0:
                    # 70.000 → 70000 (thousand separator with trailing zeros)
                    return float(cleaned.replace(".", ""))
                elif len(parts[1]) == 3 and parts[1].isdigit():
                    # Could be either 1.500 (1500) or rare case 1.234 (decimal)
                    # For safety, assume thousand separator if >= 3 digits
                    return float(cleaned.replace(".", ""))
                else:
                    # 1.5 → 1.5 (decimal point)
                    return float(cleaned)
        
        else:
            # Multiple dots → thousand separators
            # Example: 2.100.006 → remove all dots
            return float(cleaned.replace(".", ""))
    
    except:
        return 0

def extract_header_from_table(pdf_path) -> Dict:
    '''
    Extract header info từ 2 bảng đầu tiên
    Table 1: Order No, Order Date, Supplier Code, Com.Contract, Ad.Ch
    Table 2: Ordered By, Delivered To, For Store, By Supplier
    '''
    header = {
        "order_no": "",
        "order_date": "",
        "supplier_code": "",
        "com_contract": "",
        "ordered_by": "",
        "delivered_to": "",
        "for_store": ""
    }
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            page = pdf.pages[0]
            tables = page.extract_tables()
            
            if not tables or len(tables) < 2:
                return header
            
            # ===== TABLE 1: Header Info =====
            table1 = tables[0]
            if len(table1) >= 2:
                # Find column indices by header names
                headers_row = [clean_value(h).lower() for h in table1[0]]
                values_row = table1[1]
                
                for i, col_name in enumerate(headers_row):
                    if i >= len(values_row):
                        continue
                    
                    value = clean_value(values_row[i])
                    
                    if "order no" in col_name or "orderno" in col_name:
                        header["order_no"] = value
                    elif "order date" in col_name or "orderdate" in col_name:
                        header["order_date"] = value
                    elif "supplier" in col_name and "code" in col_name:
                        header["supplier_code"] = value
                    elif "contract" in col_name:
                        header["com_contract"] = value
            
            # ===== TABLE 2: Address Info =====
            table2 = tables[1]
            if len(table2) >= 2:
                headers_row = [clean_value(h).lower() for h in table2[0]]
                
                # Find column indices
                ordered_by_idx = -1
                delivered_to_idx = -1
                for_store_idx = -1
                
                for i, col_name in enumerate(headers_row):
                    if "ordered by" in col_name or "orderedby" in col_name:
                        ordered_by_idx = i
                    elif "delivered to" in col_name or "deliveredto" in col_name:
                        delivered_to_idx = i
                    elif "for store" in col_name or "forstore" in col_name:
                        for_store_idx = i
                
                # Collect all values from subsequent rows
                ordered_by_parts = []
                delivered_to_parts = []
                for_store_parts = []
                
                for row in table2[1:]:
                    if not row:
                        continue
                    
                    if ordered_by_idx >= 0 and ordered_by_idx < len(row):
                        val = clean_value(row[ordered_by_idx])
                        # Filter out generic location names
                        if val and val not in ["South", "North", "Vietnam", "East", "West", ""]:
                            ordered_by_parts.append(val)
                    
                    if delivered_to_idx >= 0 and delivered_to_idx < len(row):
                        val = clean_value(row[delivered_to_idx])
                        if val and val not in ["South", "North", "Vietnam", "East", "West", ""]:
                            delivered_to_parts.append(val)
                    
                    if for_store_idx >= 0 and for_store_idx < len(row):
                        val = clean_value(row[for_store_idx])
                        if val and val not in ["South", "North", "Vietnam", "East", "West", ""]:
                            for_store_parts.append(val)
                
                # Join with " | " separator for multi-line addresses
                header["ordered_by"] = " | ".join(ordered_by_parts)
                header["delivered_to"] = " | ".join(delivered_to_parts)
                header["for_store"] = " | ".join(for_store_parts)
    
    except Exception as e:
        print(f"Error extracting header: {e}")
    
    return header

def extract_items_from_table(pdf_path) -> List[Dict]:
    '''
    Extract items từ bảng chính (Table 3)
    Required columns: Article, Article Desc, OU Type, LV, SKU/OU, OU Qty, 
                      Free Qty, Net Purchase Price, Unit, Total Net Purchase Price
    '''
    items = []
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages):
                tables = page.extract_tables()
                
                # The items table is typically table index 2 (third table)
                # But we'll search for it by checking for "Article" header
                for table_idx, table in enumerate(tables):
                    if not table or len(table) < 2:
                        continue
                    
                    # Check if this is the items table
                    header_row = [clean_value(h).lower() for h in table[0]]
                    header_text = " ".join(header_row)
                    
                    # Must contain both "article" and at least one of the key columns
                    if "article" not in header_text:
                        continue
                    
                    # Map column indices
                    col_map = {}
                    for i, col_name in enumerate(header_row):
                        if "article desc" in col_name or (col_name == "article desc"):
                            col_map["description"] = i
                        elif col_name == "article":
                            col_map["article"] = i
                        elif "ou type" in col_name or col_name == "ou type":
                            col_map["ou_type"] = i
                        elif col_name == "lv":
                            col_map["lv"] = i
                        elif "sku/ou" in col_name or "sku ou" in col_name:
                            col_map["sku_ou"] = i
                        elif "ou qty" in col_name or col_name == "ou qty":
                            col_map["ou_qty"] = i
                        elif "free qty" in col_name or col_name == "free qty":
                            col_map["free_qty"] = i
                        elif "net purchase price" in col_name and "total" not in col_name:
                            col_map["net_price"] = i
                        elif col_name == "unit":
                            col_map["unit"] = i
                        elif "total" in col_name and "net" in col_name:
                            col_map["total_net_price"] = i
                    
                    # Extract items from data rows
                    for row_idx, row in enumerate(table[1:], start=1):
                        if not row:
                            continue
                        
                        # Get article code (must be 13 digits)
                        article_val = clean_value(row[col_map.get("article", 0)] if "article" in col_map and col_map["article"] < len(row) else "")
                        
                        # Validate: Article must be exactly 13 digits
                        if not re.match(r'^\d{13}$', article_val):
                            continue
                        
                        # Build item dictionary with proper type conversion
                        item = {
                            "article": article_val,
                            "description": clean_value(row[col_map.get("description", 1)] if "description" in col_map and col_map["description"] < len(row) else ""),
                            "ou_type": clean_value(row[col_map.get("ou_type", 2)] if "ou_type" in col_map and col_map["ou_type"] < len(row) else ""),
                            # Convert numeric fields to decimal
                            "lv": to_decimal(row[col_map.get("lv", 3)] if "lv" in col_map and col_map["lv"] < len(row) else 0),
                            "sku_ou": to_decimal(row[col_map.get("sku_ou", 4)] if "sku_ou" in col_map and col_map["sku_ou"] < len(row) else 0),
                            "ou_qty": to_decimal(row[col_map.get("ou_qty", 5)] if "ou_qty" in col_map and col_map["ou_qty"] < len(row) else 0),
                            "free_qty": to_decimal(row[col_map.get("free_qty", 6)] if "free_qty" in col_map and col_map["free_qty"] < len(row) else 0),
                            "net_price": to_decimal(row[col_map.get("net_price", 7)] if "net_price" in col_map and col_map["net_price"] < len(row) else 0),
                            "unit": clean_value(row[col_map.get("unit", 8)] if "unit" in col_map and col_map["unit"] < len(row) else ""),
                            "total_net_price": to_decimal(row[col_map.get("total_net_price", 9)] if "total_net_price" in col_map and col_map["total_net_price"] < len(row) else 0)
                        }
                        
                        items.append(item)
    
    except Exception as e:
        print(f"Error extracting items: {e}")
    
    return items

def process_pdf_format2(pdf_path, log_callback=None, debug=False):
    '''Process PDF format 2 - Complete table-based extraction with validation'''
    filename = os.path.basename(pdf_path)
    
    def log(msg):
        if log_callback:
            log_callback(msg)
    
    log(f"[Format 2] Processing: {filename}")
    
    # Step 1: Extract header from first 2 tables
    header = extract_header_from_table(pdf_path)
    
    if debug:
        log(f"  🔍 Debug - Header extracted:")
        log(f"    Order No: {header['order_no']}")
        log(f"    Order Date: {header['order_date']}")
        log(f"    Supplier: {header['supplier_code']}")
        log(f"    Contract: {header['com_contract']}")
        log(f"    Ordered By: {header['ordered_by'][:50]}..." if len(header['ordered_by']) > 50 else f"    Ordered By: {header['ordered_by']}")
        log(f"    Delivered To: {header['delivered_to'][:50]}..." if len(header['delivered_to']) > 50 else f"    Delivered To: {header['delivered_to']}")
        log(f"    For Store: {header['for_store'][:50]}..." if len(header['for_store']) > 50 else f"    For Store: {header['for_store']}")
    
    if not header["order_no"]:
        log(f"  ⚠️ Order No not found in tables")
        raise Exception("Order No not found")
    
    # Step 2: Extract items from main table
    items = extract_items_from_table(pdf_path)
    
    if not items:
        log(f"  ⚠️ No items found in tables")
        raise Exception("No items found")
    
    log(f"  ✓ Order: {header['order_no']} | Date: {header['order_date']} | Items: {len(items)}")
    
    if debug and items:
        log(f"  🔍 Debug - First item:")
        first = items[0]
        log(f"    Article: {first['article']}")
        log(f"    Desc: {first['description'][:40]}..." if len(first['description']) > 40 else f"    Desc: {first['description']}")
        log(f"    OU Type: {first['ou_type']}")
        log(f"    LV: {first['lv']}")
        log(f"    SKU/OU: {first['sku_ou']}")
        log(f"    OU Qty: {first['ou_qty']}")
        log(f"    Free Qty: {first['free_qty']}")
        log(f"    Net Price: {first['net_price']}")
        log(f"    Unit: {first['unit']}")
        log(f"    Total Net: {first['total_net_price']}")
    
    # Step 3: Prepare data for Excel export
    now = datetime.now().strftime("%H:%M:%S %d/%m/%Y")
    rows = []
    
    for item in items:
        rows.append([
            now,                        # ThoiGianThucThi
            filename,                   # FileName
            header["order_no"],         # OrderNo
            header["order_date"],       # OrderDate
            header["supplier_code"],    # SupplierCode
            header["com_contract"],     # ComContract
            header["ordered_by"],       # OrderedBy
            header["delivered_to"],     # DeliveredTo
            header["for_store"],        # ForStore
            item["article"],            # Article
            item["description"],        # ArticleDesc
            item["ou_type"],            # OUType
            item["lv"],                 # LV
            item["sku_ou"],             # SKU_OU
            item["ou_qty"],             # OUQty
            item["free_qty"],           # FreeQty
            item["net_price"],          # NetPurchasePrice
            item["unit"],               # Unit
            item["total_net_price"]     # TotalNetPurchasePrice
        ])
    
    # Step 4: Save to Excel
    if append_excel_format2(rows):
        return len(rows)
    else:
        raise Exception("Failed to save to Excel")