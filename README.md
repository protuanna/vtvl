# Công cụ nhập liệu & xuất Bộ biểu VTVL

Web app chạy local, giúp nhập dữ liệu theo 9 biểu nghiệp vụ (A1, A2, B1, B2, C1, C2, C3, D1, D2)
và xuất ra đúng file mẫu chính thức `03. Bo bieu VTVL 19-10-2026_30000_dong.xlsx` (bản 30.000
dòng/biểu), giữ nguyên toàn bộ công thức/validation của file gốc (chỉ ghi vào các ô nhập liệu).

Engine xuất Excel thao tác trực tiếp trên XML bên trong file (không dùng openpyxl để ghi) vì
file mẫu vật lý hóa sẵn công thức "shared" cho 30.000 dòng/sheet — mở bằng openpyxl rồi lưu lại
sẽ giải nén công thức, mất 2-3 phút và phình file gấp ~3 lần. Cách làm hiện tại chỉ sửa đúng các
ô cần ghi, giữ nguyên mọi ô công thức khác → xuất xong trong vài giây, file không phình.

## Chạy app

```bash
cd "/Users/vickypham/Downloads/vtvl-tool"
source venv/bin/activate
python app.py
```

Mở trình duyệt vào `http://127.0.0.1:8000` (nếu `localhost` không vào được, dùng `127.0.0.1`).
Dừng server bằng `Ctrl+C` trong terminal đang chạy.

Lưu ý macOS: không dùng cổng 5000/7000 vì trùng với AirPlay Receiver của hệ thống
(ControlCenter) — khi trùng cổng, trình duyệt sẽ nhận lỗi 403 ngẫu nhiên thay vì lỗi
"port in use" rõ ràng. App này đã cấu hình sẵn cổng 8000 để tránh xung đột.

Lần đầu chạy, nếu chưa có file `instance/vtvl.db`, cần tạo bảng:

```bash
python -c "from app import app, db; app.app_context().push(); db.create_all()"
```

## Cách dùng

1. **Trang chủ** – tạo "đơn vị" (xã/phòng/ban). Mỗi đơn vị có dữ liệu riêng cho 9 biểu.
2. **Mở đơn vị** → vào từng biểu để thêm/sửa/xóa dòng dữ liệu (chỉ các trường tương ứng
   cột trắng/nhập liệu trong file mẫu; cột công thức do Excel tự tính khi mở file).
3. **Danh mục sản phẩm** (DM_SAN_PHAM) – danh mục dùng chung toàn hệ thống, do cơ quan
   đầu mối quản lý, không gán theo đơn vị.
4. **Xuất Excel đơn vị này** (trong trang mỗi đơn vị) – xuất file chỉ chứa dữ liệu đơn vị đó.
5. **Tổng hợp & Xuất** – chọn nhiều đơn vị, xuất một file gộp dữ liệu các đơn vị đã chọn
   (nối tiếp theo thứ tự chọn) để cơ quan đầu mối (SNV) tổng hợp.
6. **Nhập từ tài liệu (AI)** (nút trong trang từng biểu) – tải lên một file PDF/DOCX
   (quyết định, danh mục...), AI đọc nội dung (kể cả bản scan, bằng nhận diện hình ảnh)
   và gợi ý các dòng dữ liệu đúng theo cột của biểu đó. Luôn hiện ra trang xem lại để
   sửa/bỏ chọn trước khi lưu — không có gì được lưu tự động vào dữ liệu chính thức.
   Cần biến môi trường `ANTHROPIC_API_KEY` (xem mục bên dưới).

### Cấu hình AI import (tùy chọn)

Tính năng "Nhập từ tài liệu (AI)" gọi API Claude để đọc file. Cần có API key riêng:

```bash
export ANTHROPIC_API_KEY="sk-ant-..."
# tùy chọn: đổi model nếu cần
export ANTHROPIC_MODEL="claude-sonnet-4-5"
cd "/Users/vickypham/Downloads/vtvl-tool" && source venv/bin/activate && python app.py
```

Không có key, các phần còn lại của app vẫn hoạt động bình thường — chỉ riêng nút
"Nhập từ tài liệu (AI)" sẽ báo lỗi rõ ràng thay vì nhập được dữ liệu.

## Dữ liệu đã có sẵn

Đơn vị "Hệ thống chính trị xã Thịnh Minh" (mã `XTM`) đã được nhập sẵn từ tài liệu thật
trong thư mục Data:
- **A1 (14 dòng)**: nhiệm vụ của Cơ quan MTTQ, Văn phòng Đảng ủy, Ban Xây dựng Đảng,
  Cơ quan UBKT Đảng ủy — trích từ `25a.QĐ...MTTQ.pdf` và `26.QĐ...TMGV ĐU.pdf` (đọc bằng AI-vision
  vì là bản scan).
- **C1 (47 dòng)**: toàn bộ danh mục vị trí việc làm theo `475-QĐ/TU` — trích từ
  `Xã Thịnh Minh.Danh mục vị trí việc làm cấp xã kèm theo QĐ 475.pdf` (đã đối chiếu khớp
  đúng tổng số 47 vị trí ghi trong văn bản gốc).

File nguồn trên (từ trang ~4 trở đi) còn có **bản mô tả chi tiết từng vị trí việc làm**
(mảng công việc, tiêu chí đánh giá, yêu cầu năng lực...) — có thể trích tiếp vào A2 và C2
nếu cần, hiện chưa nhập vì khối lượng rất lớn (300 trang mô tả cho 47 vị trí).

Mã nhiệm vụ/VTVL dùng mẫu `NV.XTM.<phòng>.<số>` / `VTVL.XTM.<số thứ tự 3 chữ số>` — bạn
có thể đổi lại cho khớp quy ước mã chính thức nếu cơ quan đầu mối đã cấp mã khác.

## Giới hạn kỹ thuật cần biết

- File mẫu có 30.000 dòng dữ liệu/sheet (dòng 3–30002). Nếu tổng số dòng của một biểu
  (khi xuất riêng hoặc tổng hợp) vượt 30.000, app báo lỗi rõ ràng thay vì tự chèn dòng
  (đúng nguyên tắc trong cẩm nang: phải báo đầu mối mở rộng file có kiểm thử).
- Sheet `5. B2_ThayDoi` không có trong app vì đây là trang tổng hợp tự động bằng công thức
  từ `4. B2_CT_TacDong` — chỉ cần nhập đúng B2_CT, mở file trong Excel sẽ tự tính.
- Sheet `12. CHIA_VUNG_CAP_XA` và các sheet kiểm tra/tiến độ/nhật ký/tham chiếu nội bộ
  (K_KiemTra, T_TienDo, S_NhatKy, NGUON_*, DM_06_NHOM, QUY_TAC_MA...) chưa đưa vào app
  (theo cẩm nang, đây là bảng quản lý của đầu mối, không phải biểu kê khai lặp lại).
- Bản mẫu hiện dùng (19-10-2026) **không có sheet `DU_BAO_5_NAM`** — sheet này chỉ xuất
  hiện ở bản sửa đổi sau (vd. v2 ngày 23-10-2026 nhắc trong video hướng dẫn của phường Tân
  Hưng). Vì vậy cột "Số người cần bố trí trong năm phương án" ở D2 trong app này là ô
  **nhập tay**, không phải ô tự tính. Nếu đổi sang bản có DU_BAO_5_NAM sau này, cần thêm
  lại biểu đó và đổi cột D2 về "tự tính" (xem ghi chú trong `schema.py`).
- Danh sách giá trị chuẩn cho các cột trạng thái/phương pháp/loại tác động đã lấy đúng
  theo tài liệu "Hướng dẫn chuẩn nhập dữ liệu Bộ biểu VTVL" (Ban Tổ chức Trung ương); cột
  "Chế độ làm việc" (C1), "Phương pháp" (C3), "Trạng thái" (C3, D2) bắt buộc chọn đúng vì
  công thức của file gốc so khớp chính xác chữ; các cột trạng thái khác chỉ gợi ý mềm.

## Cấu trúc code

- `models.py` – bảng dữ liệu (SQLite) cho Đơn vị + 9 biểu + Danh mục sản phẩm.
- `schema.py` – khai báo nhãn, kiểu nhập liệu, cột Excel tương ứng của từng trường.
- `registry.py` – nối schema với model.
- `export.py` – engine ghi dữ liệu vào bản sao file mẫu (`template_xlsx/`), giữ nguyên công thức.
- `app.py` – route Flask (CRUD theo đơn vị, danh mục sản phẩm, tổng hợp & xuất).
- `templates/` – giao diện (form nhập liệu sinh tự động từ schema).
