import openpyxl
import xlrd
import os
import re
import sys
import glob

sys.stdout.reconfigure(encoding='utf-8')

def classify_excel(fpath):
    filename = os.path.basename(fpath).upper()
    acc_type = None
    entity_type = None
    
    # 1. Content check
    try:
        if fpath.endswith('.xlsx'):
            wb = openpyxl.load_workbook(fpath, data_only=True, read_only=True)
            ws = wb.active
            text_lines = []
            for r in range(1, 20):
                for c in range(1, 10):
                    val = ws.cell(r, c).value
                    if val is not None:
                        text_lines.append(str(val).strip())
            wb.close()
        elif fpath.endswith('.xls'):
            wb = xlrd.open_workbook(fpath)
            ws = wb.sheet_by_index(0)
            text_lines = []
            for r in range(min(20, ws.nrows)):
                for c in range(min(10, ws.ncols)):
                    val = ws.cell_value(r, c)
                    if val != "":
                        text_lines.append(str(val).strip())
        else:
            text_lines = []
            
        full_text = " ".join(text_lines).upper()
        
        # Check Account
        if "136 - PHẢI THU" in full_text or "TÀI KHOẢN: 136" in full_text or "TK: 136" in full_text:
            acc_type = "136"
        elif "336 - PHẢI TRẢ" in full_text or "TÀI KHOẢN: 336" in full_text or "TK: 336" in full_text:
            acc_type = "336"
        else:
            # Look for sub-accounts like 1363..., 3363...
            if re.search(r'\b136\d{4}\b', full_text):
                acc_type = "136"
            elif re.search(r'\b336\d{4}\b', full_text):
                acc_type = "336"
                
        # Check Entity
        if "CHI NHÁNH" in full_text or "CHI NHANH" in full_text or "CÔNG TY ĐIỆN LỰC" in full_text or "CONG TY DIEN LUC" in full_text or "VŨNG TÀU" in full_text or "PCVT" in full_text:
            entity_type = "DIEN_LUC"
        elif "KẾ TOÁN NGÀNH" in full_text or "KE TOAN NGANH" in full_text or "KTN" in full_text or "(KẾ TOÁN NGÀNH)" in full_text:
            entity_type = "HCM"
        elif "TẬP ĐOÀN ĐIỆN LỰC" in full_text and "CHI NHÁNH" not in full_text:
            entity_type = "HCM"
    except Exception as e:
        print(f"Error reading {fpath}: {e}")

    # Fallback to filename if still unknown
    if not acc_type:
        if "136" in filename:
            acc_type = "136"
        elif "336" in filename:
            acc_type = "336"
            
    if not entity_type:
        if any(k in filename for k in ["HCM", "TCT", "KTN", "HO CHI MINH"]):
            entity_type = "HCM"
        elif any(k in filename for k in ["ĐIỆN LỰC", "DIEN LUC", "PCVT", "VŨNG TÀU", "VUNG TAU", "CHI NHÁNH", "CHI NHANH", "DL", "PC"]):
            entity_type = "DIEN_LUC"
            
    return acc_type, entity_type

print("Testing classification on all files in 'đầu vào':")
for p in sorted(glob.glob("đầu vào/*/*.*")):
    acc, ent = classify_excel(p)
    print(f"File: {os.path.basename(p)} -> TK: {acc} | Đơn vị: {ent}")
