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
4. Thay thế không làm hỏng định dạng phông chữ.​
5. Tìm và thay thế cả công thức và chuỗi trong ô Excel.​
6. Thay thế không làm ảnh hưởng chế độ Undo và Redo của Excel.​

   
#### Hướng dẫn sử dụng:
Hàm: =REPX(Finds, Replace (Mặc định là rỗng), Arguments,...)​
Cách viết hàm nhanh, gõ vào ô chuỗi =REPX và ấn tổ hợp phím Ctrl+Shift+A​

#### Vị trí	Tham số
- Finds	Chuỗi tìm kiếm
- Replace	Chuỗi cần thay thế
- bGlobal	Tìm toàn bộ hoặc chỉ 1
- matchCase	Có phân biệt ký tự Hoa thường

  Gõ hàm vào một ô bất kì sẽ tìm cả trang tính, hoặc chọn một vùng ô và nhập hàm tìm kiếm.
Hàm không làm ảnh hưởng đến chế độ Undo của Excel. Nếu sau khi thay thế không đúng thì có thể Undo trở lại.

#### Ví dụ sử dụng hàm REPX:
1. Tìm ký tự không phải số và thay thế thành rỗng: =REPX("\D")​
2. Tìm xóa các ký tự số: =REPX("\d")​
3. Tìm chuỗi ABC thay thế thành 123: =REPX("ABC", "123")​
4. Xóa khoảng trắng trước và sau chuỗi: =REPX(" +$|^ +")​
5. Xóa ký tự xuống dòng và thay thế thành cách: =REPX("\n", " ")​

--------------------------------------------------------------------------------------------------------------------------
# Hướng dẫn cú pháp cơ bản để đặt biểu thức chính quy
​
#### Ký hiệu định nghĩa cơ bản​
​
- \ là một ký tự bắt đầu cú pháp khi theo sau là: r, n, t, s, S, w, W, b, B, f, x## (định nghĩa Ascii 01-FF), uXXXX (định nghĩa ký tự unicode 0001-FFFF)​
- \ xóa bỏ định nghĩa của cú pháp khi theo sau thành chuỗi, gồm: \, $, ^, |, ., ?, +, *, ... và các ký tự bất kỳ nhưng không phải là ký tự như ở định nghĩa 1.​
- Ví dụ: \\ là xóa cú pháp chính nó, \|, \., \?, ...​
- | Thanh dọc hiểu là "hoặc". Ví dụ: hello|hi lấy chuỗi kí tự hello hoặc chuỗi hi​
​
#### Cú pháp định nghĩa xác định ký tự​
​
- . Dấu chấm hiểu là chụp 1 ký tự bất kỳ không bao gồm ký tự xuống dòng​
- [ ] Đối sánh bất kỳ ký tự đơn nào giữa các dấu ngoặc [ ]. Ví dụ: [AaSs] chụp 1 ký tự là A hoặc a, S, s​
- Nếu muốn nhập chính ký hiệu này vào trong nó thì nhập là [[] tìm [ , nhập là [][] thì tìm dấu [ hoặc ], hoặc [\[]​
- [^ ] Chụp 1 ký tự không chứa các ký tự. Ví dụ: [^AaSs] chụp 1 ký tự khác A và a, S, s​
- [-] Chụp từ ký tự cho đến ký tự. Ví dụ: [A-Za-z0-9]+ chụp các ký tự A đến Z, a đến z, 0 đến 9 với một hoặc nhiều lần.​
- Nếu muốn nhập chính ký hiệu gạch ngang (-) thì đặt sau cùng, ví dụ [A-] thì tìm dấu A hoặc -​
- \s Chụp 1 ký tự phân cách bao gồm các ký tự:​
- \S Chụp 1 ký tự không phải ký tự phân cách​
- \w Chụp 1 ký tự và số và dấu gạch dưới​
- \W Chụp 1 Ký tự không phải Ký tự chữ và số và dấu gạch dưới​
- \d Chụp 1 ký tự là số​
- \D Chụp 1 ký tự không thuộc số​
- \t Chụp 1 ký tự Tab​
- \r Chụp 1 ký tự trở lại dòng trên (Charcode: 13)​
- \n Chụp 1 ký tự xuống dòng (Charcode: 10)​
- \f feed​
- \uXXXX Chụp 1 ký tự Unicode (định nghĩa ký tự unicode 0001-FFFF). Ví dụ: \u1EA5 lấy ký tự "ấ", lấy ký tự từ Z đến ấ thì biểu thức là [Z-\u1EA5]​
- \xXX Chụp 1 ký tự ASCII (định nghĩa Ascii 01-FF). Ví dụ: \x41 ký tự A​
​
#### Cú pháp định nghĩa xác định không gian liền kề giữa các kiểu ký tự​
Ký tự bao gồm có chữ số, ký tự chữ, ký tự dấu, ký tự phân tách, \b và \B hiểu là định nghĩa ràng buộc liền kề của một ký tự.​
​
- \b Đến ký tự phân tách. Ví dụ: a\b hiểu là chụp ký tự a nếu theo sau a ký tự phân tách.​
- \B Không đến ký tự phân tách, hiểu là ngược lại ở trên.​
​
#### Ký hiệu chỉ định số lượng​
​
- ? không hoặc lấy một lần của cú pháp​
Nếu dấu ? nằm sau một ký hiệu xác định nhiều số lượng, thì hiểu là chỉ tìm đến trước khớp mẫu phí sau.​
Ví dụ: .+?b , tìm ký các tự bất kỳ cho đến khi gặp khớp mẫu là b, nếu không có dấu ? thì tìm bỏ qua các vị trí khớp mẫu b, cho đến khi tìm thấy khớp mẫu b cuối cùng.​
- + Một hoặc nhiều lần của cú pháp​
- * Không hoặc nhiều lần của cú pháp​
- {9} Giới hạn số lượng khớp mẫu là 9 lần. Ví dụ: a{9} lấy 9 ký tự a liên tục​
- {2,9} Lấy 2 đến 9 lần của cú pháp. Ví dụ: a{2,4} lấy 2 đến 4 ký tự a liên tục​
- {3,} Hiểu là tìm từ 3 lần khớp mẫu trở lên. Ví dụ: a{4,} lấy a từ 4 ký tự trở lên​
- {,12} Hiểu là tìm từ 12 lần khớp mẫu trở xuống. Ví dụ: a{,12} lấy a từ 12 ký tự trở xuống​
​
#### Ký hiệu buộc phải tìm khớp từ đầu hoặc ở cuối văn bản​
​
- ^ Bắt đầu phải là chuỗi khớp với biểu thức. Ví dụ: ^hello.* bắt đầu bằng hello và chuỗi bất kỳ​
- $ Kết thúc phải là chuỗi trước nó. Ví dụ: .+a$ chụp bất kỳ chuỗi nào cuối phải là a​
​
#### Nhóm trong biểu thức chính quy:​
Nhóm được định nghĩa là biểu thức nằm trong cặp ngoặc tròn ( và ).​
Trong lớp Scipting.RegExp chỉ hỗ trợ 4 dạng nhóm sau đây:​
1. ( ) Nhóm: chụp khớp biểu thức có chỉ định vị trí thứ tự nhóm.​
Ví dụ: chụp ký tự a và b thì biểu thức nhóm là (a)(b) thì hiểu a nằm trong là nhóm 1, b nằm trong nhóm 2​
Vị trí của nhóm có 2 chức năng:​
+ Dùng để thay thế nhưng giữ lại nhóm đó, sử dụng ký tự đô-la là $ và 1 số để chỉ định vị trí nhóm trong chuỗi thay thế:​
Ví dụ: trong chuỗi "abce" tìm "(a)bc" thay thế thành rỗng, nhưng giữ lại nhóm 1, thì chuỗi thay thế là "$1"​
+ Dùng để kế thừa, sử dụng dấu \ và 1 số trong biểu thức:​
Ví dụ: trong chuỗi "abbce" tìm "(a)(?=b\1c)" hiểu là tìm ký tự a và theo sau phải là b và \1 (a là nhóm 1) và ký tự c​
2. (?: ) Nhóm: Chụp nhưng không chỉ định vị trí của nhóm.​
Ví dụ: (a)(?:b)(c) chụp ký tự a và b và c, hiểu a là nhóm 1, c là nhóm 2​
3. (?= ) Nhóm: Nhóm liền kề sau nhưng không chụp.​
Ví dụ: a(?=b) chụp ký tự a, nếu theo sau a là ký tự b.​
4. (?! ) Nhóm: Nhóm tìm không khớp không chỉ định vị trí nhóm.​
Ví dụ: (?!hello) hiểu là chụp 5 ký tự bất kỳ nhưng phải khác chuỗi ký tự hello​


