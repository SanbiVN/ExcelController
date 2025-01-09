# ExcelController
 Các hàm UDF Excel thực thi hành động với Trang tính

[![Lượt tải](https://img.shields.io/github/downloads/SanbiVN/ExcelController/total.svg)](https://github.com/SanbiVN/ExcelController/releases/download/excel_controller/ExcelController_v1.43.xlam) 

[Nhấn tải ExcelController](https://github.com/SanbiVN/ExcelController/releases/download/excel_controller/ExcelController_v1.43.xlam)



### Các Hàm Thực thi thao tác với Excel

#### Thao tác với Trang tính ​
- SheetNew() - Thêm mới 1 trang tính
- SheetCopy() - Sao chép các trang tính được chọn
- SheetMove() - Di chuyển các trang tính được chọn
- SheetHide() - Ẩn các trang tính được chọn
- SheetDelete() - Xóa các trang tính được chọn
(Chức năng Di chuyển và Sao chép sẽ tự động xóa Link dự án trong công thức, biểu đồ series, Data Validation, Liên kết Macro từ đối tượng, Named, ...)

#### Hành động thực thi khác ​
- BookNewXLSX() - Thêm 1 dự án mới Xlsx
- BookNewXLSM() - Thêm 1 dự án mới Xlsm
- BookNewXLSB() - Thêm 1 dự án dạng mã hóa
- BookNewCSV() - Thêm 1 dự án CSV
- BookNewCSV_UTF8() - Thêm 1 dự án CSV-UTF8
- BookSaveAddin() - Tạo Add-in Xlam cho dự án hiện tại
- BookSaveAs() - Hiện hộp thoại Lưu như
- BookFolder() - Mở thư mục chứa dự án
​
#### Hàm kích hoạt tự động tìm bản cập nhật mới ​
(Chế độ tự động tìm bản cập nhật mặc định là tắt tìm kiếm)​
- UpdateEnableXLC() - Kích hoạt
- UpdateDisableXLC() - Hủy

## Cập nhật mới

### Hàm REPX - Tìm kiếm và thay thế nhanh sử dụng biểu thức chính quy Regular Expression cho Excel

#### Ưu điểm của chức năng tìm định dạng:​

1. Chỉ cần gõ hàm, thêm tùy chọn để tìm kiếm.​
2. Tìm trong vùng ô hoặc cả trang tính​
3. Tìm được cấu trúc văn bản phức tạp với Biểu thức chính quy.​
​
#### Hướng dẫn sử dụng:
Hàm: =REPX(Finds, Replace (Mặc định là rỗng), Arguments,...)​
Cách viết hàm nhanh, gõ vào ô chuỗi =REPX và ấn tổ hợp phím Ctrl+Shift+A​

#### Vị trí	Tham số
- Finds	Chuỗi tìm kiếm
- Replace	Chuỗi cần thay thế
- bGlobal	Tìm toàn bộ hoặc chỉ 1
- matchCase	Có phân biệt ký tự Hoa thường
