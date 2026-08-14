import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import os
import re

def export_reconciliation_excel(df_tong_hop, output_filepath, period_text="Tháng 7/2026", cross_matches=None, unmatched_by_pair=None):
    """
    Exports df_tong_hop and discrepancy analysis into a formatted Excel file.
    Sheet 1: TONG_HOP
    Sheet 2: PHAN_TICH_CHENH_LECH (Structured pair-by-pair layout matching Web UI)
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
    align_left = Alignment(horizontal='left', vertical='center', wrap_text=True)
    
    # Title Rows
    ws.merge_cells("A1:D1")
    ws["A1"] = f"Bảng TH Chênh Lệch Từ {period_text}"
    ws["A1"].font = font_title
    ws["A1"].alignment = align_left
    
    ws["E1"] = "CHÊNH LỆCH"
    ws["E1"].font = Font(name="Calibri", size=12, bold=True, color="C00000")
    ws["E1"].alignment = align_center
    
    # Table Header Row 2
    headers = ["HCM", "Điện lực", "Nợ", "Có", "CHÊNH LỆCH"]
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
    # SHEET 2: PHAN_TICH_CHENH_LECH (Structured Pair-by-Pair Layout)
    # ==========================================
    if unmatched_by_pair is not None:
        ws_an = wb.create_sheet(title="PHAN_TICH_CHENH_LECH")
        ws_an.views.sheetView[0].showGridLines = True
        
        # Styles for Sheet 2
        font_pair_header = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
        font_diff_label = Font(name="Calibri", size=11, bold=True, color="9C0006")
        font_remark = Font(name="Calibri", size=11, color="1E3C72")
        font_tbl_header = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        font_tbl_data = Font(name="Calibri", size=10)
        font_success = Font(name="Calibri", size=10, italic=True, color="006100")

        fill_pair_header = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
        fill_diff_bg = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        fill_remark_bg = PatternFill(start_color="EBF3FE", end_color="EBF3FE", fill_type="solid")
        fill_tbl_header = PatternFill(start_color="2A5298", end_color="2A5298", fill_type="solid")
        fill_success_bg = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        
        ws_an.cell(row=1, column=1, value="BÁO CÁO PHÂN TÍCH CHI TIẾT NGUYÊN NHÂN GÂY CHÊNH LỆCH KẾ TOÁN").font = font_title
        ws_an.cell(row=2, column=1, value="(Trình bày theo từng cặp tài khoản chênh lệch - trực quan như giao diện phần mềm)").font = Font(name="Calibri", size=11, italic=True)
        
        row_cur = 4
        
        # Helper to get remark text without markdown tags
        try:
            from data_processor import build_human_remark
        except ImportError:
            def build_human_remark(h, p, d, hm, pm, cm):
                return f"Chênh lệch {d:,.0f} VNĐ do chứng từ chưa hạch toán đồng bộ."

                
        for pair_key, p_data in unmatched_by_pair.items():
            diff = p_data['diff']
            if abs(diff) <= 0.01:
                continue
                
            h_code = p_data['h_code']
            p_code = p_data['p_code']
            h_missing = p_data['hcm_missing'] # HCM missing
            p_missing = p_data['pcvt_missing'] # Điện lực missing
            
            remark_raw = build_human_remark(h_code, p_code, diff, h_missing, p_missing, cross_matches or [])
            remark_clean = remark_raw.replace("**", "").replace("💡 ", "")
            
            # 1. Pair Header Box
            ws_an.merge_cells(start_row=row_cur, start_column=1, end_row=row_cur, end_column=6)
            cell_hdr = ws_an.cell(row=row_cur, column=1, value=f"🎯 CẶP TÀI KHOẢN: {h_code} (HCM) ⇄ {p_code} (Điện lực)")
            cell_hdr.font = font_pair_header
            cell_hdr.fill = fill_pair_header
            cell_hdr.alignment = align_left
            row_cur += 1
            
            # 2. Discrepancy Amount
            ws_an.merge_cells(start_row=row_cur, start_column=1, end_row=row_cur, end_column=6)
            diff_str = f"Chênh lệch cuối kỳ: {diff:,.0f} VNĐ".replace(",", ".")
            cell_diff = ws_an.cell(row=row_cur, column=1, value=diff_str)
            cell_diff.font = font_diff_label
            cell_diff.fill = fill_diff_bg
            cell_diff.alignment = align_left
            row_cur += 1
            
            # 3. Remark Box
            ws_an.merge_cells(start_row=row_cur, start_column=1, end_row=row_cur+1, end_column=6)
            cell_rmk = ws_an.cell(row=row_cur, column=1, value=remark_clean)
            cell_rmk.font = font_remark
            cell_rmk.fill = fill_remark_bg
            cell_rmk.alignment = align_left
            ws_an.row_dimensions[row_cur].height = 24
            ws_an.row_dimensions[row_cur+1].height = 24
            row_cur += 2
            
            # 4. Sub-table 1: HCM Missing
            ws_an.cell(row=row_cur, column=1, value=f"🏛️ BÚT TOÁN HCM BỊ THIẾU ({len(h_missing)} chứng từ)").font = Font(name="Calibri", size=11, bold=True, color="1F497D")
            row_cur += 1
            
            tbl_headers = ["STT", "Ngày CT", "Số CT GL", "Số Tiền (VNĐ)", "Diễn Giải Bút Toán", "Người Lập"]
            for c_i, h in enumerate(tbl_headers, 1):
                cell = ws_an.cell(row=row_cur, column=c_i, value=h)
                cell.font = font_tbl_header
                cell.fill = fill_tbl_header
                cell.alignment = align_center
            row_cur += 1
            
            if h_missing:
                for idx_m, u in enumerate(h_missing, 1):
                    ws_an.cell(row=row_cur, column=1, value=idx_m).alignment = align_center
                    ws_an.cell(row=row_cur, column=2, value=str(u['date'])[:10]).alignment = align_center
                    ws_an.cell(row=row_cur, column=3, value=str(u['gl_doc'])).alignment = align_center
                    
                    c_val = ws_an.cell(row=row_cur, column=4, value=u['net'])
                    c_val.number_format = '#,##0;(#,##0);"-"'
                    c_val.alignment = align_right
                    
                    ws_an.cell(row=row_cur, column=5, value=str(u['desc'])).alignment = align_left
                    ws_an.cell(row=row_cur, column=6, value=str(u['creator'])).alignment = align_center
                    
                    for c_i in range(1, 7):
                        ws_an.cell(row=row_cur, column=c_i).font = font_tbl_data
                        ws_an.cell(row=row_cur, column=c_i).border = thin_border
                    row_cur += 1
            else:
                ws_an.merge_cells(start_row=row_cur, start_column=1, end_row=row_cur, end_column=6)
                c_ok = ws_an.cell(row=row_cur, column=1, value="✅ Phía HCM đã hạch toán đầy đủ.")
                c_ok.font = font_success
                c_ok.fill = fill_success_bg
                c_ok.alignment = align_left
                row_cur += 1
                
            row_cur += 1
            
            # 5. Sub-table 2: Điện lực Missing
            ws_an.cell(row=row_cur, column=1, value=f"🏢 BÚT TOÁN ĐIỆN LỰC BỊ THIẾU ({len(p_missing)} chứng từ)").font = Font(name="Calibri", size=11, bold=True, color="1F497D")
            row_cur += 1
            
            for c_i, h in enumerate(tbl_headers, 1):
                cell = ws_an.cell(row=row_cur, column=c_i, value=h)
                cell.font = font_tbl_header
                cell.fill = fill_tbl_header
                cell.alignment = align_center
            row_cur += 1
            
            if p_missing:
                for idx_m, u in enumerate(p_missing, 1):
                    ws_an.cell(row=row_cur, column=1, value=idx_m).alignment = align_center
                    ws_an.cell(row=row_cur, column=2, value=str(u['date'])[:10]).alignment = align_center
                    ws_an.cell(row=row_cur, column=3, value=str(u['gl_doc'])).alignment = align_center
                    
                    c_val = ws_an.cell(row=row_cur, column=4, value=u['net'])
                    c_val.number_format = '#,##0;(#,##0);"-"'
                    c_val.alignment = align_right
                    
                    ws_an.cell(row=row_cur, column=5, value=str(u['desc'])).alignment = align_left
                    ws_an.cell(row=row_cur, column=6, value=str(u['creator'])).alignment = align_center
                    
                    for c_i in range(1, 7):
                        ws_an.cell(row=row_cur, column=c_i).font = font_tbl_data
                        ws_an.cell(row=row_cur, column=c_i).border = thin_border
                    row_cur += 1
            else:
                ws_an.merge_cells(start_row=row_cur, start_column=1, end_row=row_cur, end_column=6)
                c_ok = ws_an.cell(row=row_cur, column=1, value="✅ Phía Điện lực đã hạch toán đầy đủ.")
                c_ok.font = font_success
                c_ok.fill = fill_success_bg
                c_ok.alignment = align_left
                row_cur += 1

            # Section Separator
            row_cur += 2

        # Auto Column Widths
        ws_an.column_dimensions['A'].width = 8
        ws_an.column_dimensions['B'].width = 14
        ws_an.column_dimensions['C'].width = 14
        ws_an.column_dimensions['D'].width = 24
        ws_an.column_dimensions['E'].width = 65
        ws_an.column_dimensions['F'].width = 22

    os.makedirs(os.path.dirname(output_filepath), exist_ok=True)
    wb.save(output_filepath)
    try:
        print(f"Exported Excel successfully to: {output_filepath}")
    except Exception:
        pass
    return output_filepath

if __name__ == "__main__":
    from data_processor import process_month_folder, analyze_discrepancy_causes
    base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    folder_t7 = os.path.join(base_dir, "đầu vào", "tháng 7")
    df_th, h_accs, p_accs, _ = process_month_folder(folder_t7)
    cm, un = analyze_discrepancy_causes(df_th, h_accs, p_accs)
    out_file = os.path.join(base_dir, "đầu ra", "Bao_Cao_Doi_Soat_136_336_Thang_7_2026.xlsx")
    export_reconciliation_excel(df_th, out_file, "Tháng 7/2026", cross_matches=cm, unmatched_by_pair=un)
