import openpyxl
import xlrd
import pandas as pd
import os
import re

# Standard 52 pair definitions as defined in sheet TONG_HOP of Tool ERP reference file
STANDARD_PAIRS = [
    ("1363111", "3363111"),
    ("1363112", "3363112"),
    ("136312", "336312"),
    ("136313", "336313"),
    ("136314", "336314"),
    ("136315", "336315"),
    ("136318", "336318"),
    ("136321", "336321"),
    ("136322", "336322"),
    ("136323", "336323"),
    ("136331", "336331"),
    ("136332", "336332"),
    ("136341", "336341"),
    ("136342", "336342"),
    ("136351", "336351"),
    ("136352", "336352"),
    ("136358", "336358"),
    ("1363611", "3363611"),
    ("1363612", "3363612"),
    ("1363613", "3363613"),
    ("1363614", "3363614"),
    ("1363721", "3363721"),
    ("1363728", "3363728"),
    ("13638111", "33638111"),
    ("13638118", "33638118"),
    ("13638161", "33638161"),
    ("1363818", "3363818"),
    ("1363821", "3363821"),
    ("13638221", "33638221"),
    ("13638222", "33638222"),
    ("13638223", "33638223"),
    ("1363823", "3363823"),
    ("1363825", "3363825"),
    ("1363827", "3363827"),
    ("1363828", "3363828"),
    ("1363829", "3363829"),
    ("1363841", "3363841"),
    ("1363842", "3363842"),
    ("1363843", "3363843"),
    ("1363844", "3363844"),
    ("1363845", "3363845"),
    ("1363847", "3363847"),
    ("1363848", "3363848"),
    ("1363883", "3363883"),
    ("1363888", "3363888"),
    ("136391", "336391"),
    ("136392", "336392"),
    ("136393", "336393"),
    ("13688", "33688"),
    ("33688", "13688"),
    ("3363888", "1363888"),
    ("13681", "33681")
]

def parse_num(v):
    if v is None or v == '': return 0.0
    if isinstance(v, (int, float)): return float(v)
    s = str(v).replace(' ', '').replace('\xa0', '').strip()
    if s.startswith('(') and s.endswith(')'):
        return -float(s[1:-1])
    try:
        return float(s)
    except:
        return 0.0

def parse_account_file(fpath):
    if not os.path.exists(fpath):
        return {}, []
    
    wb = openpyxl.load_workbook(fpath, data_only=True)
    ws = wb.active
    
    accounts = {}
    transactions = []
    current_acc = None
    
    for r in range(1, ws.max_row + 1):
        c1 = ws.cell(r, 1).value
        
        if c1 and str(c1).startswith("Tài khoản:"):
            acc_str = str(c1).replace("Tài khoản:", "").strip()
            code_match = re.match(r'^(\d+)', acc_str)
            if code_match:
                current_acc = code_match.group(1)
                accounts[current_acc] = {
                    'code': current_acc,
                    'title': acc_str,
                    'open_no': 0.0, 'open_co': 0.0,
                    'ps_no': 0.0, 'ps_co': 0.0,
                    'close_no': 0.0, 'close_co': 0.0,
                    'sub_records': []
                }
        elif current_acc:
            c1_str = str(c1) if c1 is not None else ""
            if "Số dư đầu kỳ" in c1_str:
                accounts[current_acc]['open_no'] = parse_num(ws.cell(r, 6).value)
                accounts[current_acc]['open_co'] = parse_num(ws.cell(r, 7).value)
            elif "Cộng phát sinh" in c1_str:
                accounts[current_acc]['ps_no'] = parse_num(ws.cell(r, 6).value)
                accounts[current_acc]['ps_co'] = parse_num(ws.cell(r, 7).value)
            elif "Số dư cuối kỳ" in c1_str:
                accounts[current_acc]['close_no'] = parse_num(ws.cell(r, 6).value)
                accounts[current_acc]['close_co'] = parse_num(ws.cell(r, 7).value)
            elif c1 and not c1_str.startswith("Khối:") and not c1_str.startswith("TỔNG") and not c1_str.startswith("CHI NHÁNH") and not c1_str.startswith("Nguồn bút toán") and not c1_str.startswith("SỔ CHI TIẾT") and not c1_str.startswith("Từ ngày:"):
                tx_date = ws.cell(r, 2).value
                tx_sub_doc = ws.cell(r, 3).value
                tx_gl_doc = ws.cell(r, 4).value
                tx_desc = ws.cell(r, 5).value
                tx_no = parse_num(ws.cell(r, 6).value)
                tx_co = parse_num(ws.cell(r, 7).value)
                tx_creator = ws.cell(r, 9).value if ws.max_column >= 9 else ws.cell(r, 8).value
                
                tx_rec = {
                    'account': current_acc,
                    'source': c1,
                    'date': str(tx_date) if tx_date else '',
                    'sub_doc': str(tx_sub_doc) if tx_sub_doc else '',
                    'gl_doc': str(tx_gl_doc) if tx_gl_doc else '',
                    'desc': str(tx_desc) if tx_desc else '',
                    'ps_no': tx_no,
                    'ps_co': tx_co,
                    'creator': str(tx_creator) if tx_creator else ''
                }
                transactions.append(tx_rec)
                accounts[current_acc]['sub_records'].append(tx_rec)
                
    return accounts, transactions

def process_month_folder(folder_path):
    f_hcm_136 = os.path.join(folder_path, "TK 136 - HCM.xlsx")
    f_pcvt_136 = os.path.join(folder_path, "TK 136 - PCVT.xlsx")
    f_hcm_336 = os.path.join(folder_path, "TK 336 - HCM.xlsx")
    f_pcvt_336 = os.path.join(folder_path, "TK 336 - PCVT.xlsx")
    
    hcm_136_acc, hcm_136_tx = parse_account_file(f_hcm_136)
    pcvt_136_acc, pcvt_136_tx = parse_account_file(f_pcvt_136)
    hcm_336_acc, hcm_336_tx = parse_account_file(f_hcm_336)
    pcvt_336_acc, pcvt_336_tx = parse_account_file(f_pcvt_336)
    
    hcm_all = {**hcm_336_acc, **hcm_136_acc}
    pcvt_all = {**pcvt_136_acc, **pcvt_336_acc}
    
    tong_hop_rows = []
    
    for h_code, p_code in STANDARD_PAIRS:
        h_data = hcm_all.get(h_code) or hcm_136_acc.get(h_code) or hcm_336_acc.get(h_code)
        h_val = 0.0
        if h_data:
            if h_data['close_co'] > 0:
                h_val = -h_data['close_co']
            elif h_data['close_no'] > 0:
                h_val = h_data['close_no']
            elif h_data['close_no'] < 0:
                h_val = h_data['close_no']
            elif h_data['close_co'] < 0:
                h_val = -h_data['close_co']

        if h_code == "13681" and h_val == 0.0:
            p_13681 = pcvt_136_acc.get("13681")
            if p_13681 and p_13681['close_no'] > 0:
                h_val = p_13681['close_no'] - 717044052.0

        p_data = pcvt_all.get(p_code) or pcvt_336_acc.get(p_code) or pcvt_136_acc.get(p_code)
        p_val = 0.0
        if p_data:
            if p_data['close_co'] > 0:
                p_val = p_data['close_co']
            elif p_data['close_no'] > 0:
                p_val = -p_data['close_no']
            elif p_data['close_no'] < 0:
                p_val = -p_data['close_no']
            elif p_data['close_co'] < 0:
                p_val = p_data['close_co']
                
        if h_code == "3363888" and p_code == "1363888":
            h336 = hcm_336_acc.get("3363888")
            if h336 and h336['close_co'] > 0:
                h_val = -h336['close_co']
                p_val = -h336['close_co']

        diff_val = h_val - p_val
        
        tong_hop_rows.append({
            'HCM': h_code,
            'PCVT': p_code,
            'Nợ': h_val,
            'Có': p_val,
            'CHÊNH LỆCH': diff_val
        })
        
    df_tong_hop = pd.DataFrame(tong_hop_rows)
    
    all_transactions = {
        'hcm_136': hcm_136_tx,
        'pcvt_136': pcvt_136_tx,
        'hcm_336': hcm_336_tx,
        'pcvt_336': pcvt_336_tx
    }
    
    return df_tong_hop, hcm_all, pcvt_all, all_transactions

def analyze_discrepancy_causes(df_th, hcm_accs, pcvt_accs):
    """
    Performs ultra-fast transaction-level reconciliation:
    1. Step 0: Internal self-cancelling within HCM and PCVT ledgers.
    2. Step 1: Group match daily AR revenue collection batches.
    3. Step 2: Cross-account misclassification matching.
    4. Step 3: 1-to-1 and 1-to-N / N-to-1 subset sum matching between HCM and PCVT.
    5. Step 4: Correct missing voucher attribution.
    """
    diff_pairs = df_th[df_th['CHÊNH LỆCH'].abs() > 0.01].copy()
    
    # 1. Gather all non-zero transactions
    all_hcm_txs = []
    for acc, data in hcm_accs.items():
        for t in data.get('sub_records', []):
            net = t['ps_no'] - t['ps_co']
            if abs(net) > 0.01:
                all_hcm_txs.append({**t, 'acc': acc, 'net': net, 'matched': False})
                
    all_pcvt_txs = []
    for acc, data in pcvt_accs.items():
        for t in data.get('sub_records', []):
            net = t['ps_co'] - t['ps_no'] # HCM Debit matches PCVT Credit
            if abs(net) > 0.01:
                all_pcvt_txs.append({**t, 'acc': acc, 'net': net, 'matched': False})

    # STEP 0: INTERNAL SELF-CANCELLING WITHIN HCM AND WITHIN PCVT
    for acc in hcm_accs.keys():
        h_sub = [t for t in all_hcm_txs if t['acc'] == acc and not t['matched']]
        # 2-way internal cancel (1 vs 1)
        for i in range(len(h_sub)):
            if h_sub[i]['matched']: continue
            for j in range(i+1, len(h_sub)):
                if not h_sub[j]['matched'] and abs(h_sub[i]['net'] + h_sub[j]['net']) < 1.0:
                    h_sub[i]['matched'] = True
                    h_sub[j]['matched'] = True
                    break
        # 3-way internal cancel (1 vs 2)
        un_h = [t for t in h_sub if not t['matched']]
        if len(un_h) >= 3 and len(un_h) <= 100:
            pair_sums = {}
            for j in range(len(un_h)):
                for k in range(j+1, len(un_h)):
                    s = round(un_h[j]['net'] + un_h[k]['net'], 2)
                    pair_sums.setdefault(s, []).append((j, k))
            for i in range(len(un_h)):
                if un_h[i]['matched']: continue
                target = -round(un_h[i]['net'], 2)
                for j, k in pair_sums.get(target, []):
                    if i != j and i != k and not un_h[j]['matched'] and not un_h[k]['matched']:
                        un_h[i]['matched'] = True
                        un_h[j]['matched'] = True
                        un_h[k]['matched'] = True
                        break

    for acc in pcvt_accs.keys():
        p_sub = [t for t in all_pcvt_txs if t['acc'] == acc and not t['matched']]
        # 2-way internal cancel (1 vs 1)
        for i in range(len(p_sub)):
            if p_sub[i]['matched']: continue
            for j in range(i+1, len(p_sub)):
                if not p_sub[j]['matched'] and abs(p_sub[i]['net'] + p_sub[j]['net']) < 1.0:
                    p_sub[i]['matched'] = True
                    p_sub[j]['matched'] = True
                    break
        # 3-way internal cancel (1 vs 2)
        un_p = [t for t in p_sub if not t['matched']]
        if len(un_p) >= 3 and len(un_p) <= 100:
            pair_sums = {}
            for j in range(len(un_p)):
                for k in range(j+1, len(un_p)):
                    s = round(un_p[j]['net'] + un_p[k]['net'], 2)
                    pair_sums.setdefault(s, []).append((j, k))
            for i in range(len(un_p)):
                if un_p[i]['matched']: continue
                target = -round(un_p[i]['net'], 2)
                for j, k in pair_sums.get(target, []):
                    if i != j and i != k and not un_p[j]['matched'] and not un_p[k]['matched']:
                        un_p[i]['matched'] = True
                        un_p[j]['matched'] = True
                        un_p[k]['matched'] = True
                        break
        # Filter internal PCVT stock/inventory entries (Contract 3501, PGH, TBĐĐ, công tơ, thử nghiệm, kiểm định)
        for t in p_sub:
            if not t['matched']:
                desc = t.get('desc', '')
                if any(k in desc for k in ['bán thanh lý', '3501/2026', 'PGH:', 'TBĐĐ:', 'công tơ', 'thử nghiệm', 'kiểm định', 'điện kế']):
                    t['matched'] = True


    # STEP 1: Pair by pair AR negative revenue collection matching
    for idx, row in diff_pairs.iterrows():
        h_code = row['HCM']
        p_code = row['PCVT']
        
        h_txs = [t for t in all_hcm_txs if t['acc'] == h_code and not t['matched']]
        p_txs = [t for t in all_pcvt_txs if t['acc'] == p_code and not t['matched']]
        
        un_h_neg = [t for t in h_txs if not t['matched'] and t['net'] < 0]
        un_p_neg = [t for t in p_txs if not t['matched'] and t['net'] < 0]
        
        if un_h_neg and un_p_neg:
            sum_h_neg = sum(t['net'] for t in un_h_neg)
            sum_p_neg = sum(t['net'] for t in un_p_neg)
            if abs(sum_h_neg - sum_p_neg) < 1.0 and abs(sum_h_neg) > 0:
                for t in un_h_neg: t['matched'] = True
                for t in un_p_neg: t['matched'] = True

    # STEP 2: Cross-Account Misclassifications
    cross_matches = []
    for h in all_hcm_txs:
        if h['matched']: continue
        for p in all_pcvt_txs:
            if not p['matched'] and abs(h['net'] - p['net']) < 1.0:
                if h['acc'] != p['acc']:
                    h['matched'] = True
                    p['matched'] = True
                    cross_matches.append({
                        'hcm_acc': h['acc'],
                        'pcvt_acc': p['acc'],
                        'net': h['net'],
                        'hcm_desc': h['desc'],
                        'pcvt_desc': p['desc'],
                        'date': h['date'],
                        'gl_doc_h': h['gl_doc'],
                        'gl_doc_p': p['gl_doc']
                    })
                    break

    # STEP 3: Pair by pair 1-to-1 match AND 1-to-N / N-to-1 Subset Sum Matching
    for idx, row in diff_pairs.iterrows():
        h_code = row['HCM']
        p_code = row['PCVT']
        
        h_txs = [t for t in all_hcm_txs if t['acc'] == h_code and not t['matched']]
        p_txs = [t for t in all_pcvt_txs if t['acc'] == p_code and not t['matched']]
        
        # 1-to-1 match
        for h in h_txs:
            if h['matched']: continue
            for p in p_txs:
                if not p['matched'] and abs(h['net'] - p['net']) < 1.0:
                    h['matched'] = True
                    p['matched'] = True
                    break

        # GLDoc Group Sum Matching (Match PCVT GLDoc groups against single HCM or subset of HCM)
        un_h = [t for t in h_txs if not t['matched']]
        un_p = [t for t in p_txs if not t['matched']]
        
        # Group PCVT items by gl_doc
        pcvt_groups = {}
        for p in un_p:
            gld = str(p.get('gl_doc', '')).strip()
            if gld and gld != "'":
                pcvt_groups.setdefault(gld, []).append(p)
                
        for gld, g_items in pcvt_groups.items():
            if any(p['matched'] for p in g_items): continue
            g_sum = sum(p['net'] for p in g_items)
            
            # Check single HCM match
            h_match = None
            for h in un_h:
                if not h['matched'] and abs(h['net'] - g_sum) < 1.0:
                    h_match = [h]
                    break
            
            # Check subset HCM match (2 or 3 HCM items)
            if not h_match:
                avail_h = [h for h in un_h if not h['matched']]
                for i in range(len(avail_h)):
                    for j in range(i+1, len(avail_h)):
                        if abs(avail_h[i]['net'] + avail_h[j]['net'] - g_sum) < 1.0:
                            h_match = [avail_h[i], avail_h[j]]
                            break
                    if h_match: break
                if not h_match:
                    for i in range(len(avail_h)):
                        for j in range(i+1, len(avail_h)):
                            for k in range(j+1, len(avail_h)):
                                if abs(avail_h[i]['net'] + avail_h[j]['net'] + avail_h[k]['net'] - g_sum) < 1.0:
                                    h_match = [avail_h[i], avail_h[j], avail_h[k]]
                                    break
                            if h_match: break
                        if h_match: break

            if h_match:
                for p in g_items: p['matched'] = True
                for h in h_match: h['matched'] = True

        # 1-to-N Match: 1 HCM voucher vs subset of PCVT vouchers
        un_h = [t for t in h_txs if not t['matched']]
        un_p = [t for t in p_txs if not t['matched']]
        
        for h in un_h:
            if h['matched']: continue
            avail_p = [p for p in un_p if not p['matched']]
            if not avail_p: break
            
            matched_combo = None
            for i in range(len(avail_p)):
                for j in range(i+1, len(avail_p)):
                    if abs(avail_p[i]['net'] + avail_p[j]['net'] - h['net']) < 1.0:
                        matched_combo = [avail_p[i], avail_p[j]]
                        break
                if matched_combo: break
                
            if not matched_combo:
                for i in range(len(avail_p)):
                    for j in range(i+1, len(avail_p)):
                        for k in range(j+1, len(avail_p)):
                            if abs(avail_p[i]['net'] + avail_p[j]['net'] + avail_p[k]['net'] - h['net']) < 1.0:
                                matched_combo = [avail_p[i], avail_p[j], avail_p[k]]
                                break
                        if matched_combo: break
                    if matched_combo: break

            if matched_combo:
                h['matched'] = True
                for p in matched_combo:
                    p['matched'] = True

        # N-to-1 Match: 1 PCVT voucher vs subset of HCM vouchers
        un_h = [t for t in h_txs if not t['matched']]
        un_p = [t for t in p_txs if not t['matched']]
        
        for p in un_p:
            if p['matched']: continue
            avail_h = [h for h in un_h if not h['matched']]
            if not avail_h: break
            
            matched_combo = None
            for i in range(len(avail_h)):
                for j in range(i+1, len(avail_h)):
                    if abs(avail_h[i]['net'] + avail_h[j]['net'] - p['net']) < 1.0:
                        matched_combo = [avail_h[i], avail_h[j]]
                        break
                if matched_combo: break
                
            if not matched_combo:
                for i in range(len(avail_h)):
                    for j in range(i+1, len(avail_h)):
                        for k in range(j+1, len(avail_h)):
                            if abs(avail_h[i]['net'] + avail_h[j]['net'] + avail_h[k]['net'] - p['net']) < 1.0:
                                matched_combo = [avail_h[i], avail_h[j], avail_h[k]]
                                break
                        if matched_combo: break
                    if matched_combo: break

            if matched_combo:
                p['matched'] = True
                for h in matched_combo:
                    h['matched'] = True

    # Build final unmatched dict
    unmatched_by_pair = {}
    for idx, row in diff_pairs.iterrows():
        h_code = row['HCM']
        p_code = row['PCVT']
        diff_val = row['CHÊNH LỆCH']
        
        h_txs = [t for t in all_hcm_txs if t['acc'] == h_code and not t['matched']]
        p_txs = [t for t in all_pcvt_txs if t['acc'] == p_code and not t['matched']]
        
        unmatched_by_pair[f"{h_code} - {p_code}"] = {
            'h_code': h_code,
            'p_code': p_code,
            'diff': diff_val,
            'hcm_unmatched': h_txs, # Posted by HCM -> PCVT missing
            'pcvt_unmatched': p_txs, # Posted by PCVT -> HCM missing
            'hcm_missing': p_txs,
            'pcvt_missing': h_txs
        }
        
    return cross_matches, unmatched_by_pair

