import re
import os
from datetime import datetime
from typing import List, Dict, Optional

import pdfplumber
import pytesseract
from excel_handler import append_excel

def pdf_to_text(pdf_path):
    '''Trích xuất text từ PDF'''
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

def is_number(s):
    '''Kiểm tra có phải số không'''
    if not s:
        return False
    s = str(s).replace(",", "").strip()
    return bool(re.match(r'^-?\d*\.?\d+$', s))

def clean_number(s):
    '''Làm sạch số'''
    if not s:
        return ""
    s = str(s).replace(",", "").strip()
    cleaned = re.sub(r'[^\d.-]', '', s)
    return cleaned

def extract_po_number(text: str) -> Optional[str]:
    '''Tìm PO Number'''
    patterns = [
        r"P[/\\]O\s+Number[:\s]+(\d+[-/]\d+(?:[-/]\d+)?)",
        r"PO\s+Number[:\s]+(\d+[-/]\d+(?:[-/]\d+)?)",
        r"Purchase\s+Order[:\s]+(\d+[-/]\d+(?:[-/]\d+)?)",
        r"P[/\\]O[:\s]+(\d+[-/]\d+(?:[-/]\d+)?)",
        r"PO[:\s]+(\d+[-/]\d+(?:[-/]\d+)?)",
        r"P[/\\]O\s*#\s*(\d+[-/]\d+(?:[-/]\d+)?)",
        r"PO\s*#\s*(\d+[-/]\d+(?:[-/]\d+)?)",
        r"(?:P[/\\]O|PO|Purchase\s+Order).*?(\d{2,6}[-/]\d{2,6}(?:[-/]\d{1,6})?)",
        r"\b(\d{3,6}[-/]\d{3,6}(?:[-/]\d{1,6})?)\b",
    ]
    
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
        if match:
            po = match.group(1).strip()
            if '-' in po or '/' in po:
                return po
    
    return None

def parse_data_line(line: str) -> Optional[Dict]:
    '''
    Parse dòng data - HỖ TRỢ 2 FORMAT:
    
    FORMAT 1 (có VendorPartNo + 2 cột U/M):
    SKU | Description | VendorPartNo | Sell U/M | Buy U/M | BuyCost | NetBuyCost | QtyOrdCS | QtyOrdPcs | QtyRecPcs | ExtendedCost
    Ví dụ: 3571807-4 KDR dIieu NGOC CHAU 170g EA C72 3727631.00 3727631.00 .25 18.00 .00 931907.75
    
    FORMAT 2 (không có VendorPartNo + 1 cột U/M):
    SKU | Description | U/M | BuyCost | NetBuyCost | QtyOrdCS | QtyOrdPcs | QtyRecPcs | ExtendedCost
    Ví dụ: 3539973-4 NSM duoc lieu NGOC CHAU 350ml EA 2077766.00 2077766.00 .10 3.00 .00 207776.60
    '''
    # Tìm SKU
    sku_match = re.search(r'\b(\d{6,8}[-/]\d)\b', line)
    if not sku_match:
        return None
    
    sku = sku_match.group(1)
    sku_end = sku_match.end()
    
    # Lấy phần sau SKU
    after_sku = line[sku_end:].strip()
    tokens = after_sku.split()
    
    if len(tokens) < 4:
        return None
    
    # CHIẾN LƯỢC: Tìm đơn vị đo
    # Pattern 1: EA C72 (2 đơn vị - format 1)
    # Pattern 2: EA (1 đơn vị - format 2)
    
    # Thử tìm 2 đơn vị trước
    unit_pattern_2 = r'\b([A-Z]{2,3})\s+([A-Z]\d{2,3}|[A-Z]{3})\b'
    unit_match_2 = re.search(unit_pattern_2, after_sku)
    
    # Nếu không tìm thấy 2 đơn vị, thử tìm 1 đơn vị
    unit_pattern_1 = r'\b([A-Z]{2,3})\b(?!\s+[A-Z]\d{2,3})'
    unit_match_1 = re.search(unit_pattern_1, after_sku)
    
    # Ưu tiên format 1 (2 đơn vị)
    if unit_match_2:
        # FORMAT 1: Có 2 đơn vị
        sell_um = unit_match_2.group(1)
        buy_um = unit_match_2.group(2)
        unit_start = unit_match_2.start()
        unit_end = unit_match_2.end()
        
        desc_vendor_part = after_sku[:unit_start].strip()
        numbers_part = after_sku[unit_end:].strip()
        
        # Phân tích Description + VendorPartNo
        desc_vendor_tokens = desc_vendor_part.split()
        description_tokens = []
        vendor_part_no = ""
        
        for token in desc_vendor_tokens:
            if token.isdigit() and len(token) >= 10:
                vendor_part_no = token
            else:
                description_tokens.append(token)
        
        description = " ".join(description_tokens).strip()
        
    elif unit_match_1:
        # FORMAT 2: Chỉ có 1 đơn vị (không có VendorPartNo)
        sell_um = unit_match_1.group(1)
        buy_um = ""
        unit_start = unit_match_1.start()
        unit_end = unit_match_1.end()
        
        description = after_sku[:unit_start].strip()
        numbers_part = after_sku[unit_end:].strip()
        vendor_part_no = ""
        
    else:
        return None
    
    # Parse các số
    number_tokens = []
    for token in numbers_part.split():
        if is_number(token):
            number_tokens.append(clean_number(token))
    
    # Cần ít nhất 4 số
    if len(number_tokens) < 4:
        return None
    
    buy_cost = number_tokens[0]
    net_buy_cost = number_tokens[1]
    qty_ord_cs = number_tokens[2]
    
    qty_ord_pcs = ""
    qty_rec_pcs = ""
    extended_cost = number_tokens[-1]
    
    if len(number_tokens) >= 5:
        qty_ord_pcs = number_tokens[3]
    if len(number_tokens) >= 6:
        qty_rec_pcs = number_tokens[4]
        extended_cost = number_tokens[5]
    elif len(number_tokens) >= 5:
        extended_cost = number_tokens[4]
    
    # Validate extended cost
    try:
        ext_val = float(extended_cost)
        if ext_val < 10 and len(number_tokens) >= 5:
            qty_rec_pcs = extended_cost
            extended_cost = number_tokens[-2]
    except:
        pass
    
    return {
        "sku": sku,
        "description": description,
        "vendor_part_no": vendor_part_no,
        "sell_um": sell_um,
        "buy_um": buy_um,
        "buy_cost": buy_cost,
        "net_buy_cost": net_buy_cost,
        "qty_ord_cs": qty_ord_cs,
        "qty_ord_pcs": qty_ord_pcs,
        "qty_rec_pcs": qty_rec_pcs,
        "extended_cost": extended_cost
    }

def extract_items_smart(text: str) -> List[Dict]:
    '''Trích xuất items'''
    items = []
    lines = text.split('\n')
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        if re.search(r'\b(Sub\s*Total|Total|Notes|FOB)', line, re.IGNORECASE):
            break
        
        item = parse_data_line(line)
        if item:
            items.append(item)
    
    return items

def process_pdf(pdf_path, log_callback=None, debug=False):
    '''Xử lý một file PDF'''
    filename = os.path.basename(pdf_path)
    
    def log(msg):
        if log_callback:
            log_callback(msg)
    
    log(f"Đang xử lý: {filename}")
    
    text = pdf_to_text(pdf_path)
    
    if debug and len(text.strip()) >= 20:
        log(f"  📝 Debug - Text preview (first 800 chars):")
        preview = text[:800].replace('\n', ' | ')
        log(f"  {preview}")
    
    if len(text.strip()) < 20:
        raise Exception("Không đọc được nội dung")
    
    po_number = extract_po_number(text)
    if not po_number:
        log(f"  ⚠️ Không tìm thấy PO Number")
        log(f"  📄 Text preview (first 500 chars):")
        preview = text[:500].replace('\n', ' | ')
        log(f"  {preview}")
        raise Exception("Không tìm thấy PO Number")
    
    items = extract_items_smart(text)
    if not items:
        raise Exception("Không tìm thấy items")
    
    log(f"  ✓ PO: {po_number} | Items: {len(items)}")
    
    if debug and items:
        log(f"  🔍 Debug - Item đầu tiên:")
        first = items[0]
        log(f"    SKU: {first['sku']}")
        log(f"    Desc: {first['description']}")
        log(f"    VendorPartNo: {first['vendor_part_no']}")
        log(f"    SellUM: {first['sell_um']}")
        log(f"    BuyUM: {first['buy_um']}")
        log(f"    Buy: {first['buy_cost']}")
        log(f"    Net: {first['net_buy_cost']}")
        log(f"    QtyCS: {first['qty_ord_cs']}")
        log(f"    QtyOrdPcs: {first['qty_ord_pcs']}")
        log(f"    QtyRecPcs: {first['qty_rec_pcs']}")
        log(f"    Extended: {first['extended_cost']}")
    
    now = datetime.now().strftime("%H:%M:%S %d/%m/%Y")
    rows = []
    for item in items:
        rows.append([
            now, filename, po_number, item["sku"], item["description"],
            item["vendor_part_no"], item["sell_um"], item["buy_um"],
            item["buy_cost"], item["net_buy_cost"], 
            item["qty_ord_cs"], item["qty_ord_pcs"], item["qty_rec_pcs"],
            item["extended_cost"]
        ])
    
    if append_excel(rows):
        return len(rows)
    else:
        raise Exception("Không thể lưu vào Excel")