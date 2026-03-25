import pandas as pd
import json

def export_to_excel(data, output_path="scan_results.xlsx"):
    """Xuất danh sách kết quả ra file Excel"""
    try:
        df = pd.DataFrame(data)
        # Sắp xếp lại các cột cho đẹp
        cols = ['name', 'size_display', 'modified', 'hash', 'path']
        df = df[cols]
        df.to_excel(output_path, index=False)
        print(f"--- Đã xuất báo cáo Excel: {output_path} ---")
    except Exception as e:
        print(f"Lỗi xuất Excel: {e}")

def export_to_json(data, output_path="scan_results.json"):
    """Xuất danh sách kết quả ra file JSON"""
    try:
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
        print(f"--- Đã xuất báo cáo JSON: {output_path} ---")
    except Exception as e:
        print(f"Lỗi xuất JSON: {e}")