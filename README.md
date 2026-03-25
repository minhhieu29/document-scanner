#  Document Scanner Pro - Hiếu Nguyễn

Một ứng dụng Python hiện đại hỗ trợ quét, tìm kiếm nội dung chuyên sâu bên trong các định dạng file phổ biến (.docx, .pdf, .txt) với tốc độ cao nhờ Multithreading.

##  Tính năng chính
- Quét Đa Luồng: Tối ưu tốc độ tìm kiếm trên hàng nghìn file.
- Deep Scan: Tìm kiếm từ khóa nằm bên trong nội dung file Word, PDF.
- Smart Hash: Phát hiện file trùng lặp bằng thuật toán MD5.
- Giao diện hiện đại: Dark Mode cực nghệ với CustomTkinter.
- Xuất báo cáo: Hỗ trợ xuất kết quả ra file Excel chuyên nghiệp.

##  Cài đặt & Sử dụng
1. Cài đặt thư viện:
   ```bash
   pip install customtkinter python-docx pymupdf Pillow pandas openpyxl
2. Chạy ứng dụng:
   python src/gui.py
3. AI:
Để dùng tính năng AI, hãy lấy API Key tại Google AI Studio và dán vào biến api_key trong file gui.py
