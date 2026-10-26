# CONTEXT & PROJECT BRIEF: HỆ THỐNG TỰ ĐỘNG HÓA CHỨNG TỪ XNK & ĐIỀN TỜ KHAI HẢI QUAN ECUS5-VNACCS

> **Hướng dẫn cho người dùng**: Sao chép toàn bộ nội dung từ phần `--- BẮT ĐẦU NỘI DUNG GỬI CHO AI ---` đến hết để dán vào bất kỳ AI nào (ChatGPT, Claude, Gemini...). AI sẽ hiểu toàn bộ dự án ngay lập tức.

--- BẮT ĐẦU NỘI DUNG GỬI CHO AI ---

Chào AI, tôi đang phát triển một dự án tự động hóa nghiệp vụ Xuất Nhập Khẩu (RPA + Data ETL). Dưới đây là toàn bộ bối cảnh, kiến trúc, logic nghiệp vụ, công nghệ và cấu trúc mã nguồn của dự án để bạn nắm rõ và hỗ trợ tôi tiếp tục phát triển:

## 1. TỔNG QUAN DỰ ÁN & BỐI CẢNH DOANH NGHIỆP
* **Mục tiêu dự án**: Tự động hóa 100% hai khâu tốn nhiều thời gian và dễ sai sót nhất của nhân viên khai báo hải quan:
  1. Đọc và bóc tách dữ liệu từ các bộ chứng từ hỗn hợp Hóa đơn & Phiếu đóng gói (`IV + PL G-*.xlsx`), dán vào file mẫu chuẩn (`Template.xlsx`) và tính toán dữ liệu.
  2. Đọc file kết quả và tự động điều khiển phần mềm khai báo hải quan điện tử Việt Nam (**ECUS5-VNACCS** của công ty Thái Sơn) trên Windows để điền toàn bộ thông tin tờ khai xuất khẩu mà không cần gõ tay.
* **Đối tượng doanh nghiệp**: Doanh nghiệp chế xuất chuyên gia công linh kiện composite / cán gậy golf (**Công Ty TNHH ALDILA COMPOSITE PRODUCTS**, MST: `3700704014`, Chi cục Hải quan KCN Việt Nam - Singapore `43NF`, Loại hình tờ khai: `E42` - Xuất khẩu sản phẩm của DNCX).

---

## 2. KIẾN TRÚC HỆ THỐNG & CÔNG NGHỆ (TECH STACK)
Hệ thống được chia làm 2 giai đoạn độc lập:

### Giai đoạn 1: ETL & Chuẩn hóa Dữ liệu Excel (`main.py`)
* **Công nghệ**: Python 3.8+, thư viện `openpyxl`, `pathlib`, `re`, `win32com.client` (tùy chọn trên Windows).
* **Nền tảng hỗ trợ**: Chạy được trên cả **macOS** và **Windows**.
* **Đầu vào**:
  * Thư mục `input/` chứa các file Excel chứng từ gốc dạng `IV + PL G-*.xlsx` (gộp cả sheet Invoice và Packing List).
  * File mẫu chuẩn `Template.xlsx` ở thư mục gốc (đã cài đặt sẵn công thức, bảng hàng và định dạng tờ khai).
* **Đầu ra**: Thư mục `output/` chứa các file thành phẩm dạng `Template G-xx-xxxx.xlsx`.

### Giai đoạn 2: Tự động hóa Desktop GUI ECUS5-VNACCS (`ecus_fill_demo.py`)
* **Công nghệ**: Python thuần + `ctypes` (Win32 API trên Windows) + chế độ Simulation/Validation trên macOS. **Không phụ thuộc vào thư viện bên thứ 3 cồng kềnh như pywinauto hay pywin32**.
* **Nền tảng hỗ trợ**: **Chạy được trên cả 2 hệ điều hành**:
  * **Trên macOS**: Chạy ở chế độ **Kiểm thử & Mô phỏng (Simulation & Pre-validation)**: Đọc file template, hiển thị bảng xem trước 12 trường dữ liệu, tự động copy chuỗi "Ký hiệu & số hiệu" vào Clipboard, và mô phỏng toàn bộ tiến trình gõ phím & chuyển tab.
  * **Trên Windows**: Tự động kích hoạt **Win32 Automation thật**: Kết nối cửa sổ ECUS5 đang mở, tự chuyển tab 0 ("Thông tin chung"), định vị các ô nhập liệu và gõ phím trực tiếp vào tờ khai.
* **Cơ chế an toàn**: Hiển thị bảng đối chiếu giá trị (Preview), bắt buộc người dùng gõ `DIEN` mới nhập; tuyệt đối không tự động bấm nút "Ghi" hoặc "Khai báo/Gửi".

---

## 3. CẤU TRÚC THƯ MỤC VÀ TỪNG FILE
```text
project_root/
├── main.py                         # [Giai đoạn 1] ETL trích xuất Excel sang Template
├── ecus_fill_demo.py               # [Giai đoạn 2] RPA tự động điền ECUS5 (Tab Thông tin chung)
├── Template.xlsx                   # File Excel mẫu chuẩn (Sheet1: Tờ khai, Sheet2/3: Hàng hóa)
├── input/                          # Thư mục chứa các file hóa đơn gốc (IV + PL G-*.xlsx)
├── output/                         # Thư mục lưu các file kết quả đã xử lý (Template G-*.xlsx)
├── README.md                       # Tài liệu kỹ thuật chi tiết của dự án
├── ecus_win32_danhsachhang.txt     # Dump cấu trúc control của ECUS5 tại Tab 'Danh sách hàng'
├── ecus_win32_themdanhsachhang.txt # Dump cấu trúc control của Popup thêm hàng trong ECUS5
└── Old but gold/                   # Các công cụ R&D bổ trợ và phiên bản trước:
    ├── ecus_inspect.py             # Script trích xuất cây control Win32/UIA của mọi cửa sổ
    ├── ecus_goto_tab.py            # Script chuyên biệt chuyển tab SysTabControl32 bằng Win32 API
    ├── main.py & ecus_fill_demo.py # Các phiên bản đời đầu
    └── ecus_win32_*.txt & IV*.xlsx # Dữ liệu mẫu kiểm thử
```

---

## 4. CHI TIẾT LOGIC NGHIỆP VỤ & KỸ THUẬT

### 4.1. Logic của `main.py`
1. **Quét file**: Tự động nhận diện `BASE_DIR`, ưu tiên quét các file `IV + PL G-*.xlsx` trong thư mục `input/` (fallback về thư mục gốc nếu không có `input/`). Trích xuất số invoice bằng Regex `G-\d{2}-\d+`.
2. **Nhận diện Sheet**: Tìm sheet hóa đơn qua marker `invoice no` và sheet đóng gói qua marker `packing list no`.
3. **Xử lý Merged Cells**: Hàm `real_cell()` tự quy chiếu tọa độ về ô góc trên bên trái (`min_row, min_col`) để không bị mất giá trị khi ô bị gộp.
4. **Bóc tách Labels thông tin chung**: Quét các nhãn: `Mode of Delivery:`, `Forwarder/Consolidator:`, `Terms of sale:`, `Ship via:`, `Container Stuffing Location:`. Cấu hình linh hoạt cho nhãn `Terms of sale:` (quét từ cột K đến Q).
5. **Bóc tách Đóng gói & Gross Weight**:
   * Nhận dạng dòng đóng gói chứa từ khóa: `box`, `bundle`, `package`, `plastic pallet`. Đọc số lượng kiện ở ô số liền kề bên trái.
   * Tìm dòng `TOTAL WEIGHT` lấy Gross KGS. **Sanity check**: Tự tính tỷ lệ $LBS / KGS$, cảnh báo nếu lệch chuẩn $\approx 2.2046$.
6. **Bóc tách Tiền tệ & Tổng giá trị**:
   * Đọc `processing_total` (Phí gia công) và `custom_total` (Tiền hàng).
   * Tự cộng dồn các dòng chi tiết để đối chiếu chéo (reconcile) xem có khớp với ô `TOTAL` không.
7. **Sao chép bảng chi tiết hàng hóa**:
   * Quét từ dòng tiêu đề `Aldila part` (cột I) đến cột `Extension` (cột Q), dừng trước dòng `TOTAL`.
   * Lọc bỏ dòng trống, dòng ghi chú (như `GOLF SHAFTS ****`).
   * Dán vào `Sheet2` của template từ dòng `A2`.
8. **Dọn dẹp công thức thừa (`trim_extra_rows`)**: Tự động ẩn (`hidden = True`) các dòng công thức dư thừa ở `Sheet2` và `Sheet3` để triệt tiêu lỗi `#N/A`.
9. **Lưu file & Tính toán lại**: Bật cờ `target_wb.calculation.fullCalcOnLoad = True`. Nếu chạy trên Windows có `win32com`, gọi ngầm Excel `CalculateFull()` để lưu kết quả tính sẵn.

---

### 4.2. Logic của `ecus_fill_demo.py`
1. **Đọc dữ liệu từ `output/Template G-*.xlsx`**: Đọc các ô màu vàng từ `Sheet1`.
2. **Tự động sinh chuỗi 'Ký hiệu và số hiệu'**:
   `PTTT: TT, TGTT = TGGC =   {tggc} {cur},TG NVL = {tgnvl} {cur}, NNH: {nnh}`
   (với `tggc`: Tổng gia công ô B25, `tgnvl`: Tổng NVL ô B26, `cur`: Đồng tiền C12, `nnh`: Người nhận C13).
3. **Định vị & Chuyển Tab Win32 (`go_to_tab`)**:
   * Tìm control `SysTabControl32` lớn nhất trong cửa sổ `ECUS5`.
   * Cấp phát bộ nhớ trong process ECUS bằng `VirtualAllocEx`, đọc vị trí tab bằng `ReadProcessMemory` qua `TCM_GETITEMRECT`.
   * Click ngầm chuyển tab sang tab 0 ("Thông tin chung") mà không chiếm chuột thật của người dùng.
4. **Định vị ô nhập liệu động (Dynamic Relative Locating - `locate`)**:
   * Tìm nhãn `STATIC` tương ứng, sau đó tìm ô `EDIT` hiển thị cùng dòng bên phải nhãn đó (tránh lỗi sai lệch do kích thước cửa sổ hay độ phân giải màn hình).
5. **Điền 12 trường thông tin vào tab 'Thông tin chung'**:
   * Số hóa đơn (B21)
   * Ngày phát hành (B22)
   * Điều kiện giá (C11 - ví dụ `FOB`)
   * Mã đồng tiền (C12 - ví dụ `USD`)
   * Tổng trị giá hóa đơn (B27)
   * Trị giá tính thuế (B27)
   * Số lượng kiện (D16)
   * Đơn vị kiện (D17 - ví dụ `BX`, `PP`)
   * Gross KGS (D18)
   * Ký hiệu và số hiệu (chuỗi tự tạo ở mục 2)
   * Ngày khởi hành VC (B22)
   * Ngày đến ĐĐ bảo thuế (B22 - ô thứ 3)
6. **Nhập liệu & Xác thực**:
   * Bắt buộc người dùng gõ xác nhận `DIEN`. Đếm ngược 5s.
   * Gửi phím qua `SendInput` dạng Unicode (`KEYEVENTF_UNICODE`), bấm `TAB` để trigger validation của ECUS.
   * Đọc lại nội dung ô bằng `WM_GETTEXT` để báo trạng thái `OK` hoặc `KHAC`.
   * Tự động chuyển tiếp sang tab 2 ("Danh sách hàng") để nhân viên tiếp tục thao tác.

---

## 5. BẢNG ÁNH XẠ DỮ LIỆU (DATA MAPPING MATRIX)

| Trường thông tin | Nguồn gốc (`IV + PL`) | Vị trí trên `Template.xlsx` | Vị trí điền trên `ECUS5` | Định dạng xử lý |
| :--- | :--- | :--- | :--- | :--- |
| **Số hóa đơn** | Sheet IV, dưới `INVOICE NO.` | `Sheet1!B21` | `Số hóa đơn:` | Text nguyên bản |
| **Ngày phát hành** | Sheet IV, cạnh `Ship Date` | `Sheet1!B22` | `Ngày phát hành:` | `dd/mm/yyyy` |
| **Điều kiện giá** | Nhãn `Terms of sale:` | `Sheet1!C11` | `Điều kiện giá hóa đơn:` | Lấy từ đầu viết hoa (VD: `FOB`) |
| **Mã đồng tiền** | Sheet IV, cạnh `Currency` | `Sheet1!C12` | `Mã đồng tiền của hóa đơn:`| Text (VD: `USD`) |
| **Tổng trị giá hóa đơn**| Sheet IV, dòng `TOTAL` | `Sheet1!B27` | `Tổng trị giá hóa đơn:` & `Trị giá tính thuế:` | `float: 272.20` |
| **Số lượng kiện** | Sheet PL, số bên trái loại bao bì| `Sheet1!B16` $\rightarrow$ `D16` | `Số lượng kiện:` (ô số) | Số nguyên (VD: `9`) |
| **Đơn vị kiện** | Sheet PL, từ khóa `box`/`pallet` | `Sheet1!B17` $\rightarrow$ `D17` | `Số lượng kiện:` (ô đơn vị) | Mã chuẩn HQ (`BX`, `PP`) |
| **Gross KGS** | Sheet PL, dòng `TOTAL WEIGHT` | `Sheet1!B18` $\rightarrow$ `D18` | `Tổng trọng lượng hàng (Gross):` | Số thập phân (VD: `31.5`) |
| **Ký hiệu & số hiệu**| Ghép từ B25, B26, C12, C13 | Tự động ghép chuỗi | `Ký hiệu và số hiệu:` | Chuỗi định dạng PTTT... |
| **Ngày khởi hành VC** | Ngày hóa đơn / Ship Date | `Sheet1!B22` | `Ngày khởi hành vận chuyển:`| `dd/mm/yyyy` |
| **Ngày đến ĐĐ bảo thuế**| Ngày hóa đơn / Ship Date | `Sheet1!B22` | `Địa điểm đích...` (ô 3) | `dd/mm/yyyy` |
| **Chi tiết hàng hóa** | Sheet IV, dòng `Aldila part` $\rightarrow$ `TOTAL` | `Sheet2!A2` đến hết | Tab `Danh sách hàng` | Copy nguyên trạng các dòng hàng |

---

## 6. MỤC TIÊU VÀ ĐỊNH HƯỚNG BẠN CẦN HỖ TRỢ TÔI TIẾP THEO
Tôi gửi bản tóm tắt này để bạn hiểu trọn vẹn dự án. Khi tôi yêu cầu, bạn có thể hỗ trợ tôi các việc như:
1. Giải quyết các trường hợp ngoại lệ trong file Excel nguồn (ví dụ: dòng đóng gói không đúng chuẩn, định dạng ngày tháng dị biệt).
2. Phát triển tiếp tính năng tự động nhập dữ liệu vào **Tab "Danh sách hàng"** của ECUS5 (dựa trên các file phân tích control `ecus_win32_danhsachhang.txt` và `ecus_win32_themdanhsachhang.txt` đã có trong project: dùng lưới `C1TrueDBGrid1` hoặc kích hoạt phím tắt `F6: Nhập từ file Excel` của ECUS).
3. Tối ưu mã nguồn, xử lý lỗi hoặc xây dựng giao diện người dùng (GUI Desktop) đơn giản nếu cần.

Bạn đã hiểu toàn bộ dự án chưa? Vui lòng xác nhận ngắn gọn và cho biết bạn đã sẵn sàng hỗ trợ!

--- HẾT NỘI DUNG GỬI CHO AI ---
