# Dahua Warranty Checker

Ứng dụng GUI đơn giản được viết bằng Python (Tkinter) để tra cứu thông tin bảo hành của các thiết bị Dahua thông qua API của DSS (dahua.vn).

## Tính năng
- Đọc danh sách số serial từ file `.txt`.
- Tra cứu thông tin bảo hành cho từng serial thông qua API.
- Hiển thị kết quả trong bảng giao diện người dùng (GUI).
- Tự động cuộn bảng để hiển thị kết quả mới nhất.
- Xuất dữ liệu kết quả ra file Excel (`.xlsx`) với định dạng đẹp (tiêu đề, tự động điều chỉnh độ rộng cột, đường viền, đóng băng hàng tiêu đề).
- Tự động đề xuất tên file Excel khi xuất.
- Giao diện người dùng trực quan, dễ sử dụng.

## Yêu cầu
- Python 3.x
- Các thư viện Python được liệt kê trong `requirements.txt`.

Bạn có thể cài đặt tất cả các thư viện này bằng `pip`:
```bash
pip install -r requirements.txt
```

## Cách sử dụng

### 1. Chuẩn bị file Serial
- Ứng dụng đi kèm với một file mẫu `serials.txt` trên Desktop của bạn.
- File này chứa các số serial mẫu để bạn có thể kiểm tra ngay.
- Bạn có thể chỉnh sửa file `serials.txt` này hoặc tạo một file `.txt` mới của riêng bạn.
- Mỗi dòng trong file `.txt` phải chứa một số serial cần tra cứu.
  Ví dụ nội dung file `serials.txt`:
  ```
  SERIAL12345
  SERIAL67890
  ANOTHERSERIAL
  ```

### 2. Chạy ứng dụng
- **Cách 1: Chạy trực tiếp từ mã nguồn Python**
  ```bash
  python check_warranty_gui.py
  ```
- **Cách 2: Chạy từ file .exe (sau khi đã đóng gói)**
  Nhấp đúp vào file `DahuaWarrantyChecker.exe` trong thư mục `dist`.

### 3. Thao tác trên giao diện
1.  Nhấn nút **"Chọn file"** và chọn file `.txt` chứa danh sách serial của bạn.
2.  Nhấn nút **"KIỂM TRA BẢO HÀNH"** để bắt đầu quá trình tra cứu.
3.  Kết quả sẽ hiển thị trong bảng bên dưới. Bảng sẽ tự động cuộn xuống khi có kết quả mới.
4.  Để xuất dữ liệu ra Excel, nhấn nút **"Xuất Excel"**. Một hộp thoại sẽ hiện ra để bạn chọn nơi lưu file. Tên file sẽ được tự động đề xuất theo định dạng `Dahua_Warranty_Checker_YYYYMMDD_HHMMSS.xlsx`.

## Đóng gói ứng dụng thành file .exe

Bạn có thể sử dụng file `build_exe.bat` để tự động hóa quá trình này.

1.  **Đảm bảo các file cần thiết:**
    - `check_warranty_gui.py` (mã nguồn chính)
    - `serials.txt` (file serial mẫu)
    - `icon.ico` (biểu tượng ứng dụng)
    - `build_exe.bat` (file script đóng gói)
    - `requirements.txt` (danh sách thư viện)
    
    Tất cả các file này nên nằm cùng một thư mục.

2.  **Chạy `build_exe.bat`:**
    - Nhấp đúp vào file `build_exe.bat`.
    - Script sẽ tự động kiểm tra và cài đặt/cập nhật tất cả các thư viện cần thiết từ `requirements.txt`.
    - Sau đó, nó sẽ chạy `PyInstaller` để tạo file `.exe`.

3.  **Kết quả:**
    - Sau khi quá trình hoàn tất, bạn sẽ tìm thấy file thực thi `DahuaWarrantyChecker.exe` trong thư mục `dist` (ví dụ: `dist\DahuaWarrantyChecker.exe`).
    - File `.exe` này đã được tích hợp `icon.ico` làm biểu tượng và không cần cài đặt Python trên máy tính mục tiêu để chạy.

## Ghi chú
- Dữ liệu bảo hành được cung cấp bởi nhà phân phối DSS (dahua.vn).
- Nếu `icon.ico` không hiển thị, hãy kiểm tra xem file có nằm đúng vị trí và có phải là định dạng `.ico` hợp lệ không.
