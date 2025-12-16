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
    
    # OCR fallback
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
    s = s.replace(",", "").strip()
    return bool(re.match(r'^-?\d+\.?\d*$', s))

def clean_number(s):
    '''Làm sạch số'''
    if not s:
        return ""
    s = str(s).replace(",", "").strip()
    return re.sub(r'[^\d.-]', '', s)

def extract_po_number(text: str) -> Optional[str]:
    '''Tìm PO Number - CẢI THIỆN'''
    # Nhiều pattern khác nhau
    patterns = [
        # Format chuẩn: P/O Number: 123-456
        r"P[/\\]O\s+Number[:\s]+(\d+[-/]\d+(?:[-/]\d+)?)",
        r"PO\s+Number[:\s]+(\d+[-/]\d+(?:[-/]\d+)?)",
        r"Purchase\s+Order[:\s]+(\d+[-/]\d+(?:[-/]\d+)?)",
        
        # Format không có "Number": P/O: 123-456
        r"P[/\\]O[:\s]+(\d+[-/]\d+(?:[-/]\d+)?)",
        r"PO[:\s]+(\d+[-/]\d+(?:[-/]\d+)?)",
        
        # Format có # : P/O #123-456
        r"P[/\\]O\s*#\s*(\d+[-/]\d+(?:[-/]\d+)?)",
        r"PO\s*#\s*(\d+[-/]\d+(?:[-/]\d+)?)",
        
        # Format loose: tìm số có dạng XXX-XXX hoặc XXX/XXX
        r"(?:P[/\\]O|PO|Purchase\s+Order).*?(\d{2,6}[-/]\d{2,6}(?:[-/]\d{1,6})?)",
        
        # Tìm bất kỳ số nào có format XXX-XXX (fallback)
        r"\b(\d{3,6}[-/]\d{3,6}(?:[-/]\d{1,6})?)\b",
    ]
    
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
        if match:
            po = match.group(1).strip()
            # Validate: phải có ít nhất 1 dấu - hoặc /
            if '-' in po or '/' in po:
                return po
    
    return None

def parse_data_line(line: str) -> Optional[Dict]:
    '''Parse một dòng data'''
    # Tìm SKU: 6-8 số + dấu - hoặc / + 1 số
    sku_match = re.search(r'\b(\d{6,8}[-/]\d)\b', line)
    if not sku_match:
        return None
    
    sku = sku_match.group(1)
    tokens = line.split()
    
    sku_idx = -1
    for i, token in enumerate(tokens):
        if sku in token:
            sku_idx = i
            break
    
    if sku_idx == -1:
        return None
    
    desc_tokens = []
    numbers = []
    
    # Lấy phần sau SKU
    for i in range(sku_idx + 1, len(tokens)):
        token = tokens[i]
        if is_number(token):
            num_val = clean_number(token)
            if num_val:
                try:
                    val = float(num_val)
                    if val < 10000000000:
                        numbers.append(num_val)
                except:
                    pass
        elif not re.match(r'^[A-Z]{1,3}\d*$', token):
            desc_tokens.append(token)
    
    description = " ".join(desc_tokens)
    
    if len(numbers) < 4:
        return None
    
    extended_cost = numbers[-1]
    
    # Tìm Qty (thường là số nhỏ 0-1000)
    qty_candidates = []
    for i in range(len(numbers) - 1, -1, -1):
        try:
            val = float(numbers[i])
            if 0 < val <= 1000:
                qty_candidates.append((i, numbers[i]))
        except:
            pass
    
    qty_ord_cs = qty_candidates[1][1] if len(qty_candidates) > 1 else (
        qty_candidates[0][1] if qty_candidates else numbers[-2]
    )
    
    buy_cost = numbers[0] if len(numbers) > 0 else ""
    net_buy_cost = numbers[1] if len(numbers) > 1 else buy_cost
    
    return {
        "sku": sku,
        "description": description.strip(),
        "buy_cost": buy_cost,
        "net_buy_cost": net_buy_cost,
        "qty_ord_cs": qty_ord_cs,
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
    
    # DEBUG MODE: CHỈ HIỂN thị preview trong log, KHÔNG LƯU FILE
    if debug and len(text.strip()) >= 20:
        log(f"  📝 Debug - Text preview (first 500 chars):")
        preview = text[:500].replace('\n', ' | ')
        log(f"  {preview}")
    
    if len(text.strip()) < 20:
        raise Exception("Không đọc được nội dung")
    
    po_number = extract_po_number(text)
    if not po_number:
        # Hiển thị preview để debug
        log(f"  ⚠️ Không tìm thấy PO Number")
        log(f"  📄 Text preview (first 500 chars):")
        preview = text[:500].replace('\n', ' | ')
        log(f"  {preview}")
        raise Exception("Không tìm thấy PO Number - Xem log preview để biết format")
    
    items = extract_items_smart(text)
    if not items:
        raise Exception("Không tìm thấy items")
    
    log(f"  ✓ PO: {po_number} | Items: {len(items)}")
    
    now = datetime.now().strftime("%H:%M:%S %d/%m/%Y")
    rows = []
    for item in items:
        rows.append([
            now, filename, po_number, item["sku"], item["description"],
            item["buy_cost"], item["net_buy_cost"], 
            item["qty_ord_cs"], item["extended_cost"]
        ])
    
    # Append với error handling
    if append_excel(rows):
        return len(rows)
    else:
        raise Exception("Không thể lưu vào Excel")