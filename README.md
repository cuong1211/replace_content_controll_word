# Excel to Word Automation 📄✨

Công cụ tự động tạo nhiều file Word từ dữ liệu Excel bằng cách thay thế Content Control trong template Word.

## 🎯 Tính năng chính

- **Đọc danh sách từ Excel**: Tự động đọc dữ liệu từ file Excel
- **Thay thế Content Control**: Sử dụng template Word với Content Control để tạo file mới
- **Tạo hàng loạt**: Tự động tạo một file Word riêng cho mỗi tên trong danh sách
- **Xử lý lỗi thông minh**: Có nhiều phương pháp backup khi một phương pháp thất bại
- **Báo cáo chi tiết**: Hiển thị quá trình và kết quả xử lý

## 📋 Yêu cầu hệ thống

### Python Libraries
```bash
pip install pandas python-docx openpyxl
```

### File cần chuẩn bị
1. **Excel file** (`tenhs.xlsx`): Chứa cột `name` với danh sách tên
2. **Template Word** (`template.docx`): File mẫu có Content Control với tag `name`

## 🚀 Cách sử dụng

### Bước 1: Chuẩn bị file Excel
Tạo file `tenhs.xlsx` với cấu trúc:

| name |
|------|
| Nguyễn Văn An |
| Trần Thị Bình |
| Lê Văn Cường |

### Bước 2: Tạo template Word
1. Mở Microsoft Word
2. Tạo document mới
3. Thêm Content Control:
   - Ribbon **Developer** → **Controls** → **Rich Text Content Control**
   - Click chuột phải vào Content Control → **Properties**
   - Đặt **Tag**: `name`
   - Đặt **Title**: `Tên học sinh` (tùy chọn)

### Bước 3: Chạy script
```bash
python main.py
```

### Bước 4: Kiểm tra kết quả
Các file Word được tạo sẽ nằm trong thư mục `generated_files/`

## 📁 Cấu trúc project

```
project/
├── main.py              # File chính
├── tenhs.xlsx           # Dữ liệu Excel đầu vào
├── template.docx        # Template Word
└── generated_files/     # Thư mục chứa file output
    ├── Nguyễn Văn An.docx
    ├── Trần Thị Bình.docx
    └── ...
```

## ⚙️ Cấu hình

Có thể thay đổi các thông số trong hàm `main()`:

```python
excel_file = "tenhs.xlsx"        # File Excel đầu vào
template_file = "template.docx"   # File template
output_folder = "generated_files" # Thư mục output
tag_name = "name"                # Tag của Content Control
```

## 🔧 Các phương pháp xử lý

Script sử dụng 2 phương pháp thay thế Content Control:

### Method 1: Python-docx
- Sử dụng thư viện `python-docx`
- Tìm và thay thế Content Control trực tiếp
- Phương pháp chính, độ tin cậy cao

### Method 4: XML Replacement  
- Xử lý trực tiếp file XML trong docx
- Phương pháp backup khi Method 1 thất bại
- Sử dụng regex để tìm và thay thế

## 📊 Báo cáo kết quả

Script sẽ hiển thị:
- Số lượng tên được xử lý
- Trạng thái từng file (thành công/thất bại)
- Danh sách file đã tạo với kích thước
- Thống kê tổng kết

Ví dụ output:
```
🚀 BẮT ĐẦU XỬ LÝ EXCEL → WORD FILES
============================================================
📁 Thư mục output: generated_files
📊 Đọc được 7 dòng từ tenhs.xlsx
✅ Tìm thấy 7 tên hợp lệ

📝 [1/7] Xử lý: Nguyễn Văn An
   🔄 Đang tạo file cho: Nguyễn Văn An
      ✅ Thay thế: '' → 'Nguyễn Văn An'
   ✅ Method 1 thành công: generated_files/Nguyễn Văn An.docx
   ✅ Xác nhận: File chứa 'Nguyễn Văn An'

...

📊 KẾT QUẢ CUỐI CÙNG:
   ✅ Thành công: 7/7 file
   ❌ Thất bại: 0/7 file
   📁 Thư mục output: generated_files
```

## 🛠️ Troubleshooting

### Lỗi thường gặp

1. **Không tìm thấy cột 'name'**
   - Kiểm tra tên cột trong Excel phải là `name`
   - Đảm bảo không có dấu cách thừa

2. **Content Control không được thay thế**
   - Kiểm tra tag của Content Control phải là `name`
   - Thử tạo lại Content Control trong Word

3. **File template không tồn tại**
   - Đảm bảo file `template.docx` ở cùng thư mục với script
   - Kiểm tra tên file chính xác

4. **Ký tự đặc biệt trong tên**
   - Script tự động chuyển đổi ký tự không hợp lệ thành `_`
   - Ví dụ: `Name/Test` → `Name_Test.docx`

### Debug

Để xem thêm thông tin debug, kiểm tra:
- Các cột có sẵn trong Excel
- Nội dung Content Control trong template
- Quyền ghi file trong thư mục output

## 📝 Tùy chỉnh

### Thay đổi cột dữ liệu
Nếu muốn sử dụng cột khác trong Excel:

```python
# Trong hàm process_excel_to_word_files()
if "ten_hoc_sinh" not in df.columns:  # Thay đổi tên cột
    print(f"❌ Không tìm thấy cột 'ten_hoc_sinh' trong Excel!")
    return

# Và thay đổi cách lấy dữ liệu
name = row["ten_hoc_sinh"]  # Thay đổi tên cột
```

### Thêm nhiều Content Control
Có thể mở rộng để thay thế nhiều tag khác nhau:

```python
replacements = {
    "name": row["name"],
    "class": row["class"], 
    "school": row["school"]
}

for tag, value in replacements.items():
    method1_replace_content_control(doc, tag, value)
```

## 📄 License

MIT License - Tự do sử dụng và chỉnh sửa.

## 🤝 Đóng góp

Mọi đóng góp và cải thiện đều được hoan nghênh!

## 📞 Hỗ trợ

Nếu gặp vấn đề, vui lòng:
1. Kiểm tra lại các bước chuẩn bị
2. Xem phần Troubleshooting
3. Tạo issue mới với thông tin chi tiết
