import io
import os
import re
import sys
import openpyxl
import xlrd

sys.stdout.reconfigure(encoding='utf-8')

def read_header_cells(fpath_or_stream, filename=""):
    """Reads first 20 rows of cell values into a list of strings and text blocks."""
    text_blocks = []
    r1_text = ""
    r2_text = ""
    r3_text = ""
    acc_found = None
    
    is_stream = not isinstance(fpath_or_stream, str)
    stream_bytes = None
    if is_stream:
        stream_bytes = fpath_or_stream.read()
        fpath_or_stream.seek(0)
    
    # Determine format
    is_xlsx = True
    if is_stream:
        if stream_bytes and stream_bytes[:4] == b'\xd0\xcf\x11\xe0':
            is_xlsx = False
    else:
        if fpath_or_stream.lower().endswith('.xls'):
            is_xlsx = False
            
    if is_xlsx:
        try:
            wb = openpyxl.load_workbook(io.BytesIO(stream_bytes) if is_stream else fpath_or_stream, data_only=True, read_only=True)
            ws = wb.active
            for r in range(1, 20):
                row_vals = []
                for c in range(1, 10):
                    val = ws.cell(r, c).value
                    if val is not None and str(val).strip():
                        row_vals.append(str(val).strip())
                if row_vals:
                    text_blocks.append(" ".join(row_vals))
                if r == 1 and row_vals: r1_text = " ".join(row_vals)
                if r == 2 and row_vals: r2_text = " ".join(row_vals)
                if r == 3 and row_vals: r3_text = " ".join(row_vals)
            wb.close()
        except Exception:
            is_xlsx = False
            
    if not is_xlsx:
        try:
            if is_stream:
                wb = xlrd.open_workbook(file_contents=stream_bytes)
            else:
                wb = xlrd.open_workbook(fpath_or_stream)
            ws = wb.sheet_by_index(0)
            for r in range(min(20, ws.nrows)):
                row_vals = []
                for c in range(min(10, ws.ncols)):
                    val = ws.cell_value(r, c)
                    if val != "" and str(val).strip():
                        row_vals.append(str(val).strip())
                if row_vals:
                    text_blocks.append(" ".join(row_vals))
                if r == 0 and row_vals: r1_text = " ".join(row_vals)
                if r == 1 and row_vals: r2_text = " ".join(row_vals)
                if r == 2 and row_vals: r3_text = " ".join(row_vals)
        except Exception as e:
            pass
            
    return r1_text, r2_text, r3_text, " \n ".join(text_blocks)

def smart_classify(fpath_or_stream, orig_filename=""):
    fname = (orig_filename or (fpath_or_stream if isinstance(fpath_or_stream, str) else "")).upper()
    r1, r2, r3, full_text = read_header_cells(fpath_or_stream, fname)
    full_upper = full_text.upper()
    header_upper = f"{r1} {r2} {r3}".upper()
    
    # 1. Determine Account (136 vs 336)
    acc = None
    if "136 - PHẢI THU" in full_upper or "TÀI KHOẢN: 136" in full_upper or "TK: 136" in full_upper:
        acc = "136"
    elif "336 - PHẢI TRẢ" in full_upper or "TÀI KHOẢN: 336" in full_upper or "TK: 336" in full_upper:
        acc = "336"
    elif re.search(r'\b136\d{4}\b', full_upper):
        acc = "136"
    elif re.search(r'\b336\d{4}\b', full_upper):
        acc = "336"
    elif "136" in fname and "336" not in fname:
        acc = "136"
    elif "336" in fname and "136" not in fname:
        acc = "336"
        
    # 2. Determine Entity (HCM vs Điện lực)
    entity = None
    # Check for HCM markers
    is_ktn = ("KẾ TOÁN NGÀNH" in header_upper or "KE TOAN NGANH" in header_upper or "KTN" in header_upper)
    has_tap_doan = ("TẬP ĐOÀN ĐIỆN LỰC" in header_upper or "TAP DOAN DIEN LUC" in header_upper)
    has_chi_nhanh = ("CHI NHÁNH" in header_upper or "CHI NHANH" in header_upper)
    has_cong_ty_dl = ("CÔNG TY ĐIỆN LỰC" in r2.upper() or "CONG TY DIEN LUC" in r2.upper() or "PCVT" in full_upper or "VŨNG TÀU" in header_upper or "VUNG TAU" in header_upper)

    if (is_ktn or has_tap_doan) and not has_chi_nhanh:
        entity = "hcm"
    elif has_chi_nhanh or has_cong_ty_dl:
        entity = "dl"
    else:
        # Fallback to filename
        if any(k in fname for k in ["HCM", "TCT", "KTN", "HO CHI MINH", "HCMC"]):
            entity = "hcm"
        elif any(k in fname for k in ["ĐIỆN LỰC", "DIEN LUC", "PCVT", "VŨNG TÀU", "VUNG TAU", "CHI NHÁNH", "CHI NHANH", "DL", "PC"]):
            entity = "dl"

    role = None
    if entity and acc:
        role = f"{entity}_{acc}"
        
    return role, acc, entity

# Test with actual files
import glob
print("Testing smart_classify on all sample files:")
for f in sorted(glob.glob("đầu vào/*/*.*")):
    role, acc, ent = smart_classify(f)
    print(f"[{f}] -> Role: {role} (Acc: {acc}, Entity: {ent})")
