# BÀI TẬP: HOÀN THIỆN BỘ HỒ SƠ TRIỂN KHAI NỀN TẢNG SMART-HOME (SESSION 10 - EX06)

## PHẦN 1: PHÂN TÍCH — ĐẶC ĐIỂM CỦA BẢNG EXCEL "XẤU" SO VỚI BẢNG "ĐÚNG" (DATABASE MINDSET)

Trong quá trình quản lý dữ liệu bất động sản và nhà thông minh (Smart-Home), việc thiết kế bảng tính tuân thủ tư duy cơ sở dữ liệu (Database) quyết định trực tiếp đến khả năng tự động hóa và lọc dữ liệu. Dưới đây là bảng so sánh chi tiết:

| Tiêu chí | Bảng Excel "Xấu" (Vi phạm Quy tắc Vàng) | Bảng Excel "Đúng" (Chuẩn Database Flat-File) |
| --- | --- | --- |
| **1. Tiêu đề hàng đầu (Header)** | Gộp ô (Merge Cell) tiêu đề lớn ở ngay dòng 1 (`A1:E1`), dùng nhiều dòng trống làm tiêu đề chung. | Hàng đầu tiên (Header Row) là duy nhất, không gộp ô, mỗi cột là tên trường (Field) rõ ràng. |
| **2. Tính toàn vẹn của Record** | Có các hàng trống, cột trống ngắt quãng giữa các bản ghi nhằm mục đích phân chia trực quan bằng mắt. | Dữ liệu liên tục, không có hàng/cột trống ngắt quãng từ dòng đầu đến dòng cuối. |
| **3. Vị trí dòng Tổng kết (Summary)** | Đặt dòng "Tổng cộng / Trung bình" chèn ngang ngay giữa các bản ghi dữ liệu. | Dòng tổng cộng được đặt tách biệt hoàn toàn ở bên ngoài vùng dữ liệu chính (cách ít nhất 1 dòng trống). |
| **4. Kiểu dữ liệu & Định dạng** | Trộn lẫn nhiều kiểu dữ liệu trong một cột, định dạng ô lộn xộn, tự ý xuống dòng trong ô (Alt+Enter). | Mỗi cột chỉ chứa một loại thông tin đồng nhất (Text, Number, Date), dễ dàng thực hiện các hàm tính toán và AutoFilter. |
| **5. Hiệu quả vận hành** | **AutoFilter bị hỏng:** Lệnh lọc chỉ quét được phần dữ liệu phía trên hàng trống đầu tiên, bỏ sót toàn bộ phần sau. Hàm tra cứu `VLOOKUP` dễ trả về lỗi `#N/A`. | **Tối ưu hóa hoàn toàn:** AutoFilter hoạt động mượt mà trên 100% bản ghi; các hàm tính toán, PivotTable và `VLOOKUP` chạy chính xác tuyệt đối. |

---

## PHẦN 2: TRIỂN KHAI — HỒ SƠ DỰ ÁN SMART-HOME

### 1. Tài liệu Word: Đặc tả yêu cầu phần mềm SRS (Nền tảng Smart-Home & Đấu giá Nhà đất trực tuyến)

Tài liệu SRS được phân cấp cấu trúc 3 cấp (`Heading 1`, `Heading 2`, `Heading 3`) phục vụ tự động hóa mục lục (`Auto TOC`) và điều hướng qua *Navigation Pane*:

* **1. Giới thiệu Dự án Smart-Home** (`Heading 1`)
* **1.1. Tầm nhìn và sứ mệnh nền tảng** (`Heading 2`)
* **1.1.1. Giải pháp kết nối thiết bị IoT và giao dịch bất động sản tích hợp** (`Heading 3`)
* **1.1.2. Mục tiêu chiến lược giai đoạn 2026 - 2028** (`Heading 3`)


* **1.2. Đối tượng khách hàng mục tiêu** (`Heading 2`)


* **2. Yêu cầu Chức năng hệ thống** (`Heading 1`)
* **2.1. Module quản lý thiết bị Smart-Home và giám sát thời gian thực** (`Heading 2`)
* **2.2. Module đấu giá nhà đất và giao dịch trực tuyến** (`Heading 2`)
* **2.2.1. Quy trình xác thực danh tính định danh eKYC và đặt cọc đấu giá** (`Heading 3`)
* **2.2.2. Cơ chế trả giá tự động và bảo mật thông tin phiên đấu giá** (`Heading 3`)


* **2.3. Module tính toán chiết khấu, giá thuê và hoa hồng môi giới** (`Heading 2`)


* **3. Yêu cầu Phi Chức năng và Bảo mật** (`Heading 1`)
* **3.1. An toàn thông tin tài chính và mã hóa dữ liệu giao dịch** (`Heading 2`)
* **3.2. Khả năng chịu tải đồng thời của hệ thống đấu giá trực tuyến** (`Heading 2`)



---

### 2. Bảng tính Excel: Quản lý danh sách nhà đất & Tra cứu hoa hồng môi giới (VLOOKUP)

* **Cấu trúc Sheet `DanhMucNhaDat` (Bảng tra cứu hoa hồng gốc):**
* Cột A: `Mã BĐS` (VD: `BDS01`, `BDS02`...)
* Cột B: `Loại Hình`
* Cột C: `Tỷ Lệ Hoa Hồng` (VD: `1.5%`, `2.0%`...)


* **Cột tính toán tại Sheet quản lý chính (`DanhSachGiaoDich`):**
Để tự động tra cứu tỷ lệ hoa hồng môi giới từ sheet danh mục dựa theo `Mã BĐS` (giả sử đặt ở cột C) với dữ liệu đã được làm sạch khoảng trắng bằng `TRIM()`, ta sử dụng công thức:
```excel
=IFERROR(VLOOKUP(TRIM(C2), DanhMucNhaDat!$A$2:$C$100, 3, FALSE), 0)

```


*Giải thích:*
* `TRIM(C2)` giúp loại bỏ khoảng trắng thừa vô hình ở mã bất động sản.
* `DanhMucNhaDat!$A$2:$C$100` là vùng bảng tra cứu được cố định tuyệt đối bằng dấu `$`.
* `3` là số thứ tự cột lấy dữ liệu (cột Tỷ lệ hoa hồng).
* `FALSE` yêu cầu tìm kiếm chính xác tuyệt đối.
* `IFERROR` bẫy lỗi trả về `0` nếu mã BĐS không tồn tại trong danh mục.



---

### 3. Slide Pitch Deck gọi vốn (Tuân thủ Rule 10/20/30 & White Space)

Bộ slide gọi vốn dự án Smart-Home được thiết kế tinh gọn, đảm bảo không quá 10 slide, mỗi slide có nhiều khoảng trắng (**White Space**) để làm nổi bật thông điệp cốt lõi:

* **Slide 1 (Tiêu đề):**
* **Assertion Title:** `Smart-Home: Hệ Sinh Thái Quản Trị Nhà Thông Minh Và Đấu Giá Bất Động Sản Số 1 Việt Nam`
* **Nội dung (White Space rộng):** Gọi vốn vòng Seed: 500,000 USD.


* **Slide 2 (Vấn đề thị trường):**
* **Assertion Title:** `Thị Trường Giao Dịch Nhà Đất Truyền Thống Gặp Khó Khăn Về Tính Minh Bạch Và Quản Lý Vận Hành`
* **Nội dung:** 80% người mua gặp rắc rối về thông tin quy hoạch và chi phí môi giới mập mờ.


* **Slide 3 (Giải pháp cốt lõi):**
* **Assertion Title:** `Nền Tảng Smart-Home Tích Hợp Đấu Giá Trực Tuyến Giúp Tối Ưu Hóa 100% Trải Nghiệm`
* **Nội dung:** Tự động hóa từ điều khiển thiết bị IoT đến quy trình đặt cọc và đấu giá minh bạch.


* **Slide 4 (Mô hình kinh doanh):**
* **Assertion Title:** `Doanh Thu Đa Dòng Từ Phí Môi Giới, Đấu Giá Và Dịch Vụ Giá Trị Gia Tăng IoT`


* **Slide 5 (Tăng trưởng & Dữ liệu):**
* **Assertion Title:** `Hơn 10,000 Căn Hộ Đã Được Tích Hợp Hệ Thống Smart-Home Trong Quý Vừa Qua`


* **Slide 6 (Kế hoạch tài chính & Gọi vốn):**
* **Assertion Title:** `Nguồn Vốn Đầu Tư Sẽ Được Phân Bổ 60% Cho Công Nghệ R&D Và 40% Mở Rộng Thị Trường`


* **Slide 7 (Đội ngũ sáng lập):**
* **Assertion Title:** `Đội Ngũ Chuyên Gia Công Nghệ Và Bất Động Sản Giàu Kinh Nghiệm Thực Chiến`


* **Slide 8 (Call-to-Action):**
* **Assertion Title:** `Đồng Hành Cùng Smart-Home Kiến Tạo Tương Lai Ngành PropTech Việt Nam`



---

## PHẦN 3: KIỂM THỬ — SỬA LỖI TRÔI VÙNG DỮ LIỆU KHI KÉO CÔNG THỨC VLOOKUP BẰNG DẤU `$` (DOLLAR SIGN)

### 1. Hiện tượng lỗi

* Khi thực hiện viết hàm `VLOOKUP` ở dòng đầu tiên (ví dụ ô `D2`) với bảng tra cứu dạng tương đối không có dấu đô la:
`=VLOOKUP(C2, DanhMucNhaDat!A2:C100, 3, FALSE)`
* Khi người dùng thực hiện thao tác kéo (fill down) công thức từ ô `D2` xuống các dòng phía dưới (ví dụ đến ô `D50`), địa chỉ của bảng tra cứu sẽ tự động dịch chuyển tỷ lệ theo số dòng kéo xuống, biến thành `DanhMucNhaDat!A50:C148`.

### 2. Hậu quả

* Các bất động sản nằm ở những dòng đầu tiên của danh mục gốc (từ `A2` đến `A4`) sẽ bị nằm ngoài phạm vi quét mới (`A50:C148`), khiến hàm `VLOOKUP` không thể tìm thấy dữ liệu và đột ngột trả về hàng loạt lỗi `#N/A` dù thực tế mã BĐS đó có tồn tại trong hệ thống.

### 3. Giải pháp khắc phục bằng dấu `$` (Absolute Cell Reference)

Để khóa chặt cố định vùng bảng tra cứu không bao giờ bị dịch chuyển dù công thức được kéo xuống hàng nghìn dòng, ta phải sử dụng ký hiệu đô la `$` cho cả tọa độ cột và hàng trong `table_array`:

```excel
=VLOOKUP(C2, DanhMucNhaDat!$A$2:$C$100, 3, FALSE)

```

* **Cơ chế hoạt động:** Ký hiệu `$A$2:$C$100` ra lệnh cho bảng tính cố định cứng cáp ranh giới của vùng dữ liệu tra cứu tại ô góc trên bên trái ($A$2) và góc dưới bên phải ($C$100). Nhờ đó, tính năng tra cứu hoạt động ổn định và chính xác 100% trên toàn bộ tập dữ liệu.