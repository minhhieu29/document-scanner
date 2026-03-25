import os
import hashlib
from pathlib import Path
from datetime import datetime

class DocumentScanner:
    def __init__(self):
        self.extensions = {'.pdf', '.docx', '.doc', '.xlsx', '.xls', '.pptx', '.txt', '.md'}
        self.exclude_dirs = {'Windows', 'AppData', 'Program Files', 'node_modules', '$Recycle.Bin'}

    def get_file_hash(self, file_path):
        """Tính mã MD5 để nhận diện nội dung file (phát hiện trùng lặp)"""
        hash_md5 = hashlib.md5()
        try:
            with open(file_path, "rb") as f:
                # Chỉ đọc 4096 bytes đầu để lấy hash nhanh, không cần đọc cả file nặng
                for chunk in iter(lambda: f.read(4096), b""):
                    hash_md5.update(chunk)
            return hash_md5.hexdigest()
        except:
            return None

    def scan_directory(self, root_path):
        results = []
        path_obj = Path(root_path)
        
        print(f"--- Đang quét: {root_path} ---")
        
        # rglob('*') vẫn ổn với thư mục nhỏ, nhưng tao thêm lọc ngay từ đầu
        for entry in path_obj.rglob('*'):
            try:
                if entry.is_file() and entry.suffix.lower() in self.extensions:
                    # Kiểm tra xem có nằm trong vùng cấm không
                    if not any(ex in entry.parts for ex in self.exclude_dirs):
                        
                        file_hash = self.get_file_hash(entry)
                        
                        file_info = {
                            "name": entry.name,
                            "path": str(entry.absolute()),
                            "size_bytes": entry.stat().st_size,
                            "size_display": f"{entry.stat().st_size / 1024:.2f} KB",
                            "modified": datetime.fromtimestamp(entry.stat().st_mtime).strftime('%Y-%m-%d %H:%M:%S'),
                            "hash": file_hash
                        }
                        results.append(file_info)
                        print(f"[FOUND] {entry.name} - Hash: {file_hash[:8]}...")
            except (PermissionError, OSError):
                continue
            
        return results

if __name__ == "__main__":
    from utils import export_to_excel, export_to_json # Import hàm từ utils.py
    
    scanner = DocumentScanner()
    test_path = os.path.expanduser("~/Documents") 
    data = scanner.scan_directory(test_path)
    
    if data:
        export_to_excel(data)
        export_to_json(data)
    
    print(f"\n--- TỔNG KẾT: Tìm thấy {len(data)} file ---")