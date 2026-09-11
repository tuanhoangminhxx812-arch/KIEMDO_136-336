import os
import sys
import io

sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, r'd:\DATA\DATA KTTH\KIEMDO_136-336\dữ liệu')
from data_processor import process_month_folder, smart_classify_file

print("=== 1. Test smart_classify_file with in-memory ByteStreams ===")
test_files = [
    r'd:\DATA\DATA KTTH\KIEMDO_136-336\đầu vào\Tháng 8\TK 136 - HCM.xlsx',
    r'd:\DATA\DATA KTTH\KIEMDO_136-336\đầu vào\Tháng 8\TK 136 - Điện lực.xlsx',
    r'd:\DATA\DATA KTTH\KIEMDO_136-336\đầu vào\Tháng 8\TK 336 - HCM.xlsx',
    r'd:\DATA\DATA KTTH\KIEMDO_136-336\đầu vào\Tháng 8\TK 336 - Điện lực.xlsx'
]

stream_dict = {}
for p in test_files:
    with open(p, 'rb') as f:
        b = f.read()
    bio = io.BytesIO(b)
    # Give random filename
    fake_name = "random_name_" + os.path.basename(p)[-8:]
    role, acc, ent, label = smart_classify_file(bio, fake_name)
    print(f"File [{fake_name}] -> Role: {role} ({label})")
    stream_dict[role] = bio

print(f"\n=== 2. Test in-memory process_month_folder ===")
df_th, h_accs, p_accs, all_txs = process_month_folder(None, file_dict=stream_dict)
print(f"Success! Rows: {len(df_th)}, HCM Accs: {len(h_accs)}, DL Accs: {len(p_accs)}")

print("\n=== All Tests Passed! ===")
