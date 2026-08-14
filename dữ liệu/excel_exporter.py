import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import os

def export_reconciliation_excel(df_tong_hop, output_filepath, period_text="Tháng 7/2026", cross_matches=None, unmatched_by_pair=None):
    """
    Exports df_tong_hop and discrepancy analysis into a formatted Excel file.
    """
    wb = openpyxl.Workbook()
    
    # ==========================================
    # SHEET 1: TONG_HOP
    # ==========================================
    ws = wb.active
    ws.title = "TONG_HOP"
    ws.views.sheetView[0].showGridLines = True
    
    # Styles definition
    font_title = Font(name="Calibri", size=14, bold=True, color="1F497D")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_data = Font(name="Calibri", size=11)
    font_bold = Font(name="Calibri", size=11, bold=True)
    font_diff = Font(name="Calibri", size=11, bold=True, color="9C0006")
    
    fill_header = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    fill_diff_header = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
    fill_diff_row = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    fill_total = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    
    thick_bottom = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='double', color='000000')
    )
    
    align_center = Alignment(horizontal='center', vertical='center')
    align_right = Alignment(horizontal='right', vertical='center')
    align_left = Alignment(horizontal='left', vertical='center')
    
    # Title Rows
    ws.merge_cells("A1:D1")
    ws["A1"] = f"Bảng TH Chênh Lệch Từ {period_text}"
    ws["A1"].font = font_title
    ws["A1"].alignment = align_left
    
    ws["E1"] = "CHÊNH LỆCH"
    ws["E1"].font = Font(name="Calibri", size=12, bold=True, color="C00000")
    ws["E1"].alignment = align_center
    
    # Table Header Row 2
    headers = ["HCM", "PCVT", "Nợ", "Có", "CHÊNH LỆCH"]
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=2, column=col_idx, value=h)
        cell.font = font_header
        cell.alignment = align_center
        if h == "CHÊNH LỆCH":
            cell.fill = fill_diff_header
        else:
            cell.fill = fill_header
            
    # Data Rows (Row 3 onwards)
    start_row = 3
    for idx, row in df_tong_hop.iterrows():
        current_row = start_row + idx
        hcm_val = row['HCM']
        pcvt_val = row['PCVT']
        no_val = row['Nợ']
        co_val = row['Có']
        diff_val = row['CHÊNH LỆCH']
        
        c_hcm = ws.cell(row=current_row, column=1, value=str(hcm_val))
        c_pcvt = ws.cell(row=current_row, column=2, value=str(pcvt_val))
        c_no = ws.cell(row=current_row, column=3, value=no_val)
        c_co = ws.cell(row=current_row, column=4, value=co_val)
        c_diff = ws.cell(row=current_row, column=5, value=f"=C{current_row}-D{current_row}")
        
        c_hcm.alignment = align_center
        c_pcvt.alignment = align_center
        c_no.alignment = align_center
        c_co.alignment = align_center
        c_diff.alignment = align_center
        
        c_hcm.font = font_data
        c_pcvt.font = font_data
        c_no.font = font_data
        c_co.font = font_data
        c_diff.font = font_data
        
        c_no.number_format = '#,##0;(#,##0);"-"'
        c_co.number_format = '#,##0;(#,##0);"-"'
        c_diff.number_format = '#,##0;(#,##0);"-"'
        
        for col_i in range(1, 6):
            ws.cell(row=current_row, column=col_i).border = thin_border
            
        if abs(diff_val) > 0.01:
            c_diff.font = font_diff
            c_diff.fill = fill_diff_row

    # Total Summary Row
    end_data_row = start_row + len(df_tong_hop) - 1
    total_row = end_data_row + 1
    
    ws.cell(row=total_row, column=1, value="TỔNG CỘNG").font = font_bold
    ws.cell(row=total_row, column=1).alignment = align_center
    ws.cell(row=total_row, column=2, value="").font = font_bold
    
    c_tot_no = ws.cell(row=total_row, column=3, value=f"=SUM(C3:C{end_data_row})")
    c_tot_co = ws.cell(row=total_row, column=4, value=f"=SUM(D3:D{end_data_row})")
    c_tot_diff = ws.cell(row=total_row, column=5, value=f"=SUM(E3:E{end_data_row})")
    
    for c, val_cell in [(3, c_tot_no), (4, c_tot_co), (5, c_tot_diff)]:
        val_cell.font = font_bold
        val_cell.alignment = align_center
        val_cell.number_format = '#,##0;(#,##0);"-"'
        
    for col_i in range(1, 6):
        cell = ws.cell(row=total_row, column=col_i)
        cell.fill = fill_total
        cell.border = thick_bottom
        
    column_widths = {'A': 15, 'B': 15, 'C': 25, 'D': 25, 'E': 25}
    for col, width in column_widths.items():
        ws.column_dimensions[col].width = width

    # ==========================================
    # SHEET 2: PHAN_TICH_CHENH_LECH (Discrepancy Analysis)
    # ==========================================
    if cross_matches is not None or unmatched_by_pair is not None:
        ws_an = wb.create_sheet(title="PHAN_TICH_CHENH_LECH")
        ws_an.views.sheetView[0].showGridLines = True
        
        ws_an.cell(row=1, column=1, value="BÁO CÁO PHÂN TÍCH CHI TIẾT NGUYÊN NHÂN GÂY CHÊNH LỆCH").font = font_title
        ws_an.cell(row=2, column=1, value="(Tất cả các số liệu đúng đã được ẩn đi)").font = Font(name="Calibri", size=11, italic=True)
        
        # Section 1: Misclassifications
        ws_an.cell(row=4, column=1, value="1. DANH SÁCH BÚT TOÁN HẠCH TOÁN LỆCH TÀI KHOẢN ĐỐI ỨNG").font = font_bold
        cm_headers = ["STT", "Ngày CT", "TK HCM", "TK PCVT", "Số Tiền Lệch TK", "Diễn Giải (HCM)", "Diễn Giải (PCVT)"]
        for c_i, h in enumerate(cm_headers, 1):
            cell = ws_an.cell(row=5, column=c_i, value=h)
            cell.font = font_header
            cell.fill = PatternFill(start_color="2A5298", end_color="2A5298", fill_type="solid")
            cell.alignment = align_center
            
        row_cur = 6
        if cross_matches:
            for idx_cm, cm in enumerate(cross_matches, 1):
                ws_an.cell(row=row_cur, column=1, value=idx_cm).alignment = align_center
                ws_an.cell(row=row_cur, column=2, value=str(cm['date'])[:10]).alignment = align_center
                ws_an.cell(row=row_cur, column=3, value=str(cm['hcm_acc'])).alignment = align_center
                ws_an.cell(row=row_cur, column=4, value=str(cm['pcvt_acc'])).alignment = align_center
                
                c_val = ws_an.cell(row=row_cur, column=5, value=cm['net'])
                c_val.number_format = '#,##0;(#,##0);"-"'
                c_val.alignment = align_center
                c_val.font = font_diff
                
                ws_an.cell(row=row_cur, column=6, value=str(cm['hcm_desc'])).alignment = align_left
                ws_an.cell(row=row_cur, column=7, value=str(cm['pcvt_desc'])).alignment = align_left
                
                for c_i in range(1, 8):
                    ws_an.cell(row=row_cur, column=c_i).border = thin_border
                row_cur += 1
        else:
            ws_an.cell(row=6, column=1, value="Không có bút toán hạch toán lệch tài khoản.")
            row_cur = 7
            
        row_cur += 2
        # Section 2: Unmatched per pair
        ws_an.cell(row=row_cur, column=1, value="2. DANH SÁCH BÚT TOÁN KHUYẾT BỊ LỆCH THEO TỪNG CẶP TÀI KHOẢN").font = font_bold
        row_cur += 1
        
        un_headers = ["TK HCM", "TK PCVT", "Số Tiền Lệch Cặp", "Đơn Vị Khuyết", "Ngày CT", "Số CT GL", "Số Tiền Bút Toán", "Diễn Giải Bút Toán", "Người Lập"]
        for c_i, h in enumerate(un_headers, 1):
            cell = ws_an.cell(row=row_cur, column=c_i, value=h)
            cell.font = font_header
            cell.fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
            cell.alignment = align_center
            
        row_cur += 1
        if unmatched_by_pair:
            for pair_key, p_data in unmatched_by_pair.items():
                h_un = p_data['hcm_unmatched']
                p_un = p_data['pcvt_unmatched']
                
                for u in h_un:
                    ws_an.cell(row=row_cur, column=1, value=p_data['h_code']).alignment = align_center
                    ws_an.cell(row=row_cur, column=2, value=p_data['p_code']).alignment = align_center
                    
                    c_d = ws_an.cell(row=row_cur, column=3, value=p_data['diff'])
                    c_d.number_format = '#,##0;(#,##0);"-"'
                    c_d.alignment = align_center
                    c_d.font = font_diff
                    
                    ws_an.cell(row=row_cur, column=4, value="Chỉ có tại HCM (Khuyết PCVT)").alignment = align_center
                    ws_an.cell(row=row_cur, column=5, value=str(u['date'])[:10]).alignment = align_center
                    ws_an.cell(row=row_cur, column=6, value=str(u['gl_doc'])).alignment = align_center
                    
                    c_v = ws_an.cell(row=row_cur, column=7, value=u['net'])
                    c_v.number_format = '#,##0;(#,##0);"-"'
                    c_v.alignment = align_center
                    
                    ws_an.cell(row=row_cur, column=8, value=str(u['desc'])).alignment = align_left
                    ws_an.cell(row=row_cur, column=9, value=str(u['creator'])).alignment = align_center
                    
                    for c_i in range(1, 10):
                        ws_an.cell(row=row_cur, column=c_i).border = thin_border
                    row_cur += 1

                for u in p_un:
                    ws_an.cell(row=row_cur, column=1, value=p_data['h_code']).alignment = align_center
                    ws_an.cell(row=row_cur, column=2, value=p_data['p_code']).alignment = align_center
                    
                    c_d = ws_an.cell(row=row_cur, column=3, value=p_data['diff'])
                    c_d.number_format = '#,##0;(#,##0);"-"'
                    c_d.alignment = align_center
                    c_d.font = font_diff
                    
                    ws_an.cell(row=row_cur, column=4, value="Chỉ có tại PCVT (Khuyết HCM)").alignment = align_center
                    ws_an.cell(row=row_cur, column=5, value=str(u['date'])[:10]).alignment = align_center
                    ws_an.cell(row=row_cur, column=6, value=str(u['gl_doc'])).alignment = align_center
                    
                    c_v = ws_an.cell(row=row_cur, column=7, value=u['net'])
                    c_v.number_format = '#,##0;(#,##0);"-"'
                    c_v.alignment = align_center
                    
                    ws_an.cell(row=row_cur, column=8, value=str(u['desc'])).alignment = align_left
                    ws_an.cell(row=row_cur, column=9, value=str(u['creator'])).alignment = align_center
                    
                    for c_i in range(1, 10):
                        ws_an.cell(row=row_cur, column=c_i).border = thin_border
                    row_cur += 1

        for c_i in range(1, 10):
            col_ltr = get_column_letter(c_i)
            ws_an.column_dimensions[col_ltr].width = 20 if c_i not in [6, 8] else 40

    os.makedirs(os.path.dirname(output_filepath), exist_ok=True)
    wb.save(output_filepath)
    try:
        print(f"Exported Excel successfully to: {output_filepath}")
    except Exception:
        pass
    return output_filepath

if __name__ == "__main__":
    from data_processor import process_month_folder, analyze_discrepancy_causes
    folder_t7 = r"d:\DATA\DATA KTTH\KIEMDO_136-336\đầu vào\tháng 7"
    df_th, h_accs, p_accs, _ = process_month_folder(folder_t7)
    cm, un = analyze_discrepancy_causes(df_th, h_accs, p_accs)
    out_file = r"d:\DATA\DATA KTTH\KIEMDO_136-336\đầu ra\Bao_Cao_Doi_Soat_136_336_Thang_7_2026.xlsx"
    export_reconciliation_excel(df_th, out_file, "Tháng 7/2026", cross_matches=cm, unmatched_by_pair=un)
