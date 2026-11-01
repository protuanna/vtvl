# Dinh nghia 9 bieu nghiep vu VTVL + danh muc san pham dung chung.
# "col" la cot trong file mau "03. Bo bieu VTVL 19-10-2026_30000_dong.xlsx" ma
# truong nay se duoc ghi vao khi xuat Excel (chi cac cot trang/nhap lieu - khong
# dong vao cot cong thuc da khoa). Du lieu bat dau tu dong 3, toi da dong 30002
# (30.000 dong).
#
# "help" la mo ta ngan cho tung truong, lay sat theo cot "Chuan nhap de cong
# thuc chay dung" trong tai lieu "Huong_dan_chuan_nhap_du_lieu_Bo_bieu_VTVL.docx"
# (Ban To chuc Trung uong). Hien duoi moi o nhap trong giao dien.

F_TEXT = "text"
F_TEXTAREA = "textarea"
F_NUMBER = "number"
F_INT = "int"
F_REF = "ref"  # goi y ma da co san (tra tu bieu khac qua "ref"), van cho nhap tu do
F_SELECT = "select"  # bat buoc chon trong danh sach (dung voi data validation list cua file goc)
F_SELECT_SOFT = "select_soft"  # goi y bang datalist, van cho nhap tu do

# Cac danh sach gia tri duoi day lay dung tu Phan IV "GIA TRI CHUAN PHAI GO DUNG
# TUNG CHU" cua tai lieu "Huong_dan_chuan_nhap_du_lieu_Bo_bieu_VTVL.docx" (Ban To
# chuc Trung uong, To cong tac xay dung VTVL/bien che 2027-2031).
CHE_DO_LAM_VIEC_OPTIONS = ["Phổ biến (40 giờ/tuần)", "Đặc thù"]
PHUONG_PHAP_OPTIONS = ["PP1", "PP2", "PP3", "PP1+PP2", "PP1+PP3", "PP2+PP3"]
D2_TRANG_THAI_OPTIONS = [
    "Đang xây dựng",
    "Đã hoàn thành ở đơn vị",
    "Đã thẩm định",
    "Đã được quyết định",
]
C3_TRANG_THAI_OPTIONS = ["Tính thử", "Chưa đủ dữ liệu", "Đã kiểm chứng", "Không áp dụng"]
B2_TRANG_THAI_OPTIONS = ["Đã kiểm chứng", "Chưa kiểm chứng", "Ước lượng tham khảo"]
XAC_NHAN_GIO_RIENG_OPTIONS = ["Đã kiểm chứng"]
D1_TRANG_THAI_OPTIONS = ["Đã kiểm chứng"]
A1_TRANG_THAI_OPTIONS = [
    "Tiếp tục thực hiện",
    "Mới tiếp nhận",
    "Đã chuyển giao",
    "Không còn thực hiện",
]
B1_TRANG_THAI_CAN_CU_OPTIONS = ["Kiểm đếm, ước lượng thực tế", "Có định mức"]
LOAI_TAC_DONG_OPTIONS = [
    "Số hóa",
    "Cải tiến quy trình",
    "Thuê dịch vụ/đặt hàng",
    "Tự chủ",
    "Xã hội hóa",
    "Công việc mới phát sinh",
]
CO_SO_B1_OPTIONS = ["Trước thay đổi, có điều chỉnh", "Đã đo sau thay đổi (không trừ)"]
DOI_CHIEU_TRUNG_OPTIONS = ["Đã đối chiếu"]
DM_SAN_PHAM_TRANG_THAI_OPTIONS = ["Đã xác nhận", "Đang kiểm chứng", "Không áp dụng"]

BIEU_SCHEMAS = [
    {
        "key": "a1",
        "label": "A1 – Danh mục nhiệm vụ",
        "sheet_name": "1. A1_NhiemVu",
        "model_name": "NhiemVu",
        "step": 1,
        "step_note": "Rà soát nhiệm vụ của đơn vị. Không phụ thuộc biểu nào khác — làm trước tiên.",
        "fields": [
            {"name": "ma_nhiem_vu", "label": "Mã nhiệm vụ", "col": "A", "type": F_TEXT, "required": True,
             "help": "Duy nhất trong biểu. Đặt theo một quy ước thống nhất của đơn vị. Không để trống."},
            {"name": "ten_nhiem_vu", "label": "Tên nhiệm vụ", "col": "B", "type": F_TEXT, "required": True,
             "help": "Ghi đúng tên nhiệm vụ theo văn bản giao."},
            {"name": "co_quan", "label": "Cơ quan", "col": "C", "type": F_TEXT,
             "help": "Tên cơ quan, đơn vị thực hiện nhiệm vụ."},
            {"name": "can_cu", "label": "Căn cứ", "col": "D", "type": F_TEXT,
             "help": "Điều, khoản, tên văn bản giao nhiệm vụ."},
            {"name": "ky", "label": "Kỳ", "col": "E", "type": F_TEXT,
             "help": "Kỳ số liệu của bản ghi. Dùng một cách ghi duy nhất cho toàn file (vd: 01/10/2025–30/9/2026)."},
            {"name": "trang_thai", "label": "Trạng thái", "col": "F", "type": F_SELECT_SOFT, "options": A1_TRANG_THAI_OPTIONS,
             "help": "Tình trạng nhiệm vụ trong kỳ. Ghi thống nhất cho cả biểu."},
            {"name": "nguon", "label": "Nguồn", "col": "G", "type": F_TEXT,
             "help": "Văn bản hoặc hồ sơ làm căn cứ ghi nhận nhiệm vụ."},
        ],
    },
    {
        "key": "a2",
        "label": "A2 – Công việc và sản phẩm",
        "sheet_name": "2. A2_CongViec",
        "model_name": "CongViec",
        "step": 2,
        "step_note": "Rã nhiệm vụ (A1) thành công việc, gắn với VTVL (C1). Nên nhập C1 trước A2.",
        "fields": [
            {"name": "ma_cong_viec", "label": "Mã công việc", "col": "A", "type": F_TEXT, "required": True,
             "help": "Duy nhất trong biểu. Là khóa để B1 và B2_CT gọi công việc. Không để trống."},
            {"name": "ma_nhiem_vu", "label": "Mã nhiệm vụ", "col": "B", "type": F_REF, "ref": "a1", "required": True,
             "help": "Trùng đúng mã đã có ở A1."},
            {"name": "ma_vtvl", "label": "Mã VTVL", "col": "C", "type": F_REF, "ref": "c1", "required": True,
             "help": "Trùng đúng mã đã có ở C1 (cột A). B1, B2, C3 dùng cột này để tra VTVL. Mỗi công việc chỉ gắn một VTVL."},
            {"name": "ten_cong_viec", "label": "Tên công việc", "col": "D", "type": F_TEXT,
             "help": "Mô tả công việc cụ thể."},
            {"name": "san_pham", "label": "Sản phẩm", "col": "E", "type": F_TEXT,
             "help": "Tên sản phẩm đầu ra của công việc."},
            {"name": "don_vi_tinh", "label": "Đơn vị tính", "col": "F", "type": F_TEXT,
             "help": "Đơn vị đếm sản phẩm. Phải cùng đơn vị với Số SP/năm ở B1."},
            {"name": "chat_luong", "label": "Chất lượng", "col": "G", "type": F_TEXT,
             "help": "Yêu cầu chất lượng sản phẩm."},
            {"name": "can_cu_nguon", "label": "Căn cứ (nguồn)", "col": "H", "type": F_TEXT,
             "help": "Quy trình hoặc văn bản làm căn cứ."},
        ],
    },
    {
        "key": "b1",
        "label": "B1 – Khối lượng",
        "sheet_name": "3. B1_KhoiLuong",
        "model_name": "KhoiLuong",
        "step": 3,
        "step_note": "Khối lượng sản phẩm và thời gian cho từng công việc. Chỉ dòng \"ĐẠT\" mới được cộng vào C3.",
        "fields": [
            {"name": "ma_cong_viec", "label": "Mã công việc", "col": "A", "type": F_REF, "ref": "a2", "required": True,
             "help": "Trùng đúng mã có ở A2. Mỗi công việc một dòng. Không tìm thấy ở A2 thì báo \"THIẾU MÃ VTVL\"."},
            {"name": "so_sp_nam", "label": "Số SP/năm", "col": "C", "type": F_NUMBER,
             "help": "Số sản phẩm trong kỳ, nhập số ≥ 0. Có thể để trống chỉ khi chỉ ghi giờ bổ sung ở cột Giờ bổ sung."},
            {"name": "tb_phut_sp", "label": "Trung bình phút/SP (đo thực tế)", "col": "D", "type": F_NUMBER,
             "help": "Số phút trung bình cho một sản phẩm theo đo thực tế, số > 0. Dùng khi không có phút chuẩn đã xác nhận."},
            {"name": "gio_bo_sung", "label": "Giờ bổ sung (công việc liên quan)", "col": "E", "type": F_NUMBER,
             "help": "Giờ của công việc liên quan chưa nằm trong giờ sản phẩm, số ≥ 0. Nên ghi rõ nguồn ở cột Nguồn thời gian."},
            {"name": "nguon_sl", "label": "Nguồn SL", "col": "H", "type": F_TEXT,
             "help": "Nguồn của số lượng sản phẩm. Bắt buộc; để trống báo \"THIẾU NGUỒN SẢN LƯỢNG/CÔNG VIỆC\"."},
            {"name": "nguon_thoi_gian", "label": "Nguồn thời gian", "col": "I", "type": F_TEXT,
             "help": "Phiếu đo hoặc văn bản định mức làm căn cứ thời gian. Bắt buộc khi chọn \"Có định mức\"."},
            {"name": "trang_thai_can_cu_tg", "label": "Trạng thái căn cứ thời gian", "col": "J", "type": F_SELECT_SOFT, "options": B1_TRANG_THAI_CAN_CU_OPTIONS,
             "help": "Chọn đúng một trong hai. Để trống báo \"CHƯA CHỌN TRẠNG THÁI\"."},
            {"name": "ky", "label": "Kỳ", "col": "K", "type": F_TEXT,
             "help": "Kỳ số liệu của sản lượng và thời gian. Bắt buộc."},
            {"name": "ma_sp_danh_muc", "label": "Mã SP trong danh mục", "col": "L", "type": F_REF, "ref": "san_pham",
             "help": "Có thể để trống. Nếu điền, phải đúng mã ở DM_SAN_PHAM để các cột tự tra hoạt động."},
            {"name": "giai_trinh_khac_chuan", "label": "Giải trình khác chuẩn", "col": "S", "type": F_TEXTAREA,
             "help": "Chỉ điền khi muốn dùng phút đo thực tế khác phút chuẩn đã xác nhận. Phải bắt đầu bằng \"Đã xác nhận:\" (có dấu hai chấm)."},
        ],
    },
    {
        "key": "b2ct",
        "label": "B2_CT – Chi tiết tác động",
        "sheet_name": "4. B2_CT_TacDong",
        "model_name": "TacDong",
        "step": 4,
        "step_note": "Chỉ nhập khi có tác động thật (số hóa, thuê dịch vụ, cải tiến...). Biểu B2_ThayDoi tự tổng hợp theo công việc, không cần nhập.",
        "fields": [
            {"name": "ma_cong_viec", "label": "Mã công việc", "col": "A", "type": F_REF, "ref": "a2", "required": True,
             "help": "Trùng đúng mã ở A2. Một công việc có thể có nhiều dòng (mỗi dòng một công đoạn)."},
            {"name": "loai_tac_dong", "label": "Loại tác động", "col": "C", "type": F_SELECT_SOFT, "options": LOAI_TAC_DONG_OPTIONS,
             "help": "Chọn đúng một trong sáu loại. Loại khác bị báo \"LOẠI TÁC ĐỘNG KHÔNG HỢP LỆ\"."},
            {"name": "cong_doan", "label": "Công đoạn", "col": "D", "type": F_TEXT,
             "help": "Tên công đoạn chịu tác động, bắt buộc. Không trùng giữa các dòng cùng công việc, cùng kỳ."},
            {"name": "sp_ap_dung", "label": "SP áp dụng", "col": "E", "type": F_TEXT,
             "help": "Số sản phẩm chịu tác động, số > 0. Không vượt tổng Số SP/năm của công việc đó ở B1."},
            {"name": "phut_truoc", "label": "Phút trước", "col": "F", "type": F_NUMBER,
             "help": "Phút/SP của công đoạn trước thay đổi."},
            {"name": "phut_sau", "label": "Phút sau", "col": "G", "type": F_NUMBER,
             "help": "Phút/SP của công đoạn sau thay đổi. Nếu lớn hơn Phút trước thì bắt buộc ghi Giờ việc mới."},
            {"name": "gio_viec_moi", "label": "Giờ việc mới", "col": "I", "type": F_NUMBER,
             "help": "Giờ phát sinh thêm do thay đổi, số ≥ 0. Với \"Công việc mới phát sinh\" có thể chỉ điền cột này."},
            {"name": "nguon_bang_chung", "label": "Nguồn/bằng chứng", "col": "J", "type": F_TEXT,
             "help": "Biên bản nghiệm thu, hợp đồng, log, kết quả theo dõi. Bắt buộc."},
            {"name": "ky", "label": "Kỳ", "col": "K", "type": F_TEXT,
             "help": "Kỳ phát sinh/đo tác động. Bắt buộc."},
            {"name": "co_so_b1", "label": "Cơ sở B1", "col": "L", "type": F_SELECT_SOFT, "options": CO_SO_B1_OPTIONS,
             "help": "\"Trước thay đổi, có điều chỉnh\" = giờ giảm sẽ được trừ; \"Đã đo sau thay đổi\" = giờ giảm tính bằng 0. Bắt buộc."},
            {"name": "trang_thai", "label": "Trạng thái", "col": "M", "type": F_SELECT_SOFT, "options": B2_TRANG_THAI_OPTIONS,
             "help": "Chỉ \"Đã kiểm chứng\" mới được tính vào kết quả. \"Ước lượng tham khảo\" không tính."},
            {"name": "kinh_phi_truoc", "label": "Kinh phí trước", "col": "Q", "type": F_NUMBER,
             "help": "Thông tin tham khảo, không tham gia công thức kiểm tra."},
            {"name": "kinh_phi_sau", "label": "Kinh phí sau", "col": "R", "type": F_NUMBER,
             "help": "Thông tin tham khảo, cùng đơn vị tiền tệ với Kinh phí trước."},
            {"name": "doi_chieu_trung", "label": "Đối chiếu trùng/phạm vi (cán bộ xác nhận)", "col": "S", "type": F_SELECT_SOFT, "options": DOI_CHIEU_TRUNG_OPTIONS,
             "help": "Chỉ gõ đúng \"Đã đối chiếu\" sau khi đã soát trùng công đoạn và phạm vi. Để trống báo \"CHỜ SOÁT TRÙNG\"."},
        ],
    },
    {
        "key": "c1",
        "label": "C1 – Vị trí việc làm",
        "sheet_name": "6. C1_VTVL",
        "model_name": "Vtvl",
        "step": 5,
        "step_note": "Danh mục VTVL và quỹ giờ — nguồn tra cứu của hầu hết các biểu khác. Nên nhập trước A2.",
        "fields": [
            {"name": "ma_vtvl", "label": "Mã VTVL", "col": "A", "type": F_TEXT, "required": True,
             "help": "Duy nhất trong biểu. Là khóa tra cứu của A2, C2, C3, D1, D2. Không để trống."},
            {"name": "ten_vtvl", "label": "Tên VTVL", "col": "B", "type": F_TEXT, "required": True,
             "help": "Tên vị trí việc làm theo danh mục được phê duyệt."},
            {"name": "loai_nhan_luc", "label": "Loại nhân lực", "col": "C", "type": F_TEXT,
             "help": "Công chức; Viên chức; Hợp đồng lao động... Ghi thống nhất, cùng cách ghi với D1."},
            {"name": "quyet_dinh", "label": "Quyết định", "col": "D", "type": F_TEXT,
             "help": "Quyết định phê duyệt danh mục VTVL (số, ngày, cơ quan)."},
            {"name": "can_cu_quy_gio", "label": "Căn cứ xác định quỹ giờ", "col": "F", "type": F_TEXT,
             "help": "Văn bản làm căn cứ. Bắt buộc khi dùng giờ đặc thù ở cột Giờ riêng được xác nhận."},
            {"name": "trang_thai", "label": "Trạng thái", "col": "G", "type": F_TEXT,
             "help": "Tình trạng VTVL (vd: Tiếp tục sử dụng, Cần cập nhật). Ghi thống nhất."},
            {"name": "gio_rieng_xac_nhan", "label": "Giờ riêng được xác nhận", "col": "H", "type": F_NUMBER,
             "help": "Chỉ điền khi VTVL có quỹ giờ riêng khác quỹ giờ chung (1.904 giờ/năm). Số > 0."},
            {"name": "xac_nhan_gio_rieng", "label": "Xác nhận giờ riêng", "col": "I", "type": F_SELECT_SOFT, "options": XAC_NHAN_GIO_RIENG_OPTIONS,
             "help": "Gõ đúng \"Đã kiểm chứng\" khi giờ riêng đã được cấp có thẩm quyền xác nhận."},
            {"name": "che_do_lam_viec", "label": "Chế độ làm việc", "col": "J", "type": F_SELECT, "options": CHE_DO_LAM_VIEC_OPTIONS, "required": True,
             "help": "Chọn đúng một trong hai. Bắt buộc."},
        ],
    },
    {
        "key": "c2",
        "label": "C2 – Mô tả năng lực",
        "sheet_name": "7. C2_NangLuc",
        "model_name": "NangLuc",
        "step": 6,
        "step_note": "Mô tả năng lực theo VTVL. Chưa có công thức nào phụ thuộc biểu này — nhập để lưu hồ sơ.",
        "fields": [
            {"name": "ma_vtvl", "label": "Mã VTVL", "col": "A", "type": F_REF, "ref": "c1", "required": True,
             "help": "Trùng đúng mã ở C1."},
            {"name": "trach_nhiem", "label": "Trách nhiệm", "col": "B", "type": F_TEXTAREA,
             "help": "Trách nhiệm chính của vị trí."},
            {"name": "nang_luc_chung", "label": "Năng lực chung", "col": "C", "type": F_TEXTAREA,
             "help": "Các năng lực chung theo khung năng lực."},
            {"name": "chuyen_mon", "label": "Chuyên môn", "col": "D", "type": F_TEXTAREA,
             "help": "Yêu cầu kiến thức, kỹ năng chuyên môn."},
            {"name": "lanh_dao", "label": "Lãnh đạo", "col": "E", "type": F_TEXTAREA,
             "help": "Năng lực lãnh đạo, quản lý; ghi \"Không áp dụng\" nếu vị trí không cần."},
            {"name": "can_cu", "label": "Căn cứ", "col": "F", "type": F_TEXT,
             "help": "Văn bản làm căn cứ."},
            {"name": "nguon_mo_ta", "label": "Nguồn bản mô tả", "col": "G", "type": F_TEXT,
             "help": "Hồ sơ hoặc bản mô tả VTVL gốc."},
        ],
    },
    {
        "key": "c3",
        "label": "C3 – Nhu cầu",
        "sheet_name": "8. C3_NhuCau",
        "model_name": "NhuCau",
        "step": 7,
        "step_note": "Tổng hợp nhu cầu từng VTVL. Giờ B1/B2 tự lấy — chỉ cần chọn phương pháp và kết luận số người. Làm sau khi B1/B2 đã \"ĐẠT\".",
        "fields": [
            {"name": "ma_vtvl", "label": "Mã VTVL", "col": "A", "type": F_REF, "ref": "c1", "required": True,
             "help": "Duy nhất trong biểu, phải có ở C1. Không có báo \"MÃ VTVL CHƯA CÓ Ở C1\"."},
            {"name": "phuong_phap", "label": "Phương pháp", "col": "B", "type": F_SELECT, "options": PHUONG_PHAP_OPTIONS, "required": True,
             "help": "PP1 = theo cơ cấu/chức danh; PP2 = giờ lao động ÷ quỹ giờ; PP3 = ước lượng chuyên môn. Có thể chọn kết hợp."},
            {"name": "nguoi_pp1", "label": "Người PP1", "col": "H", "type": F_NUMBER,
             "help": "Số người theo PP1. Bắt buộc khi phương pháp có PP1."},
            {"name": "nguoi_pp3", "label": "Số người theo đánh giá chuyên môn có căn cứ (PP3)", "col": "I", "type": F_NUMBER,
             "help": "Số người theo PP3. Bắt buộc khi phương pháp có PP3; phải giải trình ở cột Nguồn/giải trình."},
            {"name": "nguoi_can_bo_tri", "label": "Người cần bố trí", "col": "J", "type": F_NUMBER,
             "help": "Kết luận số người cần bố trí do cán bộ nghiệp vụ chốt (số nguyên, không âm). Không phải số tự tính."},
            {"name": "nguon_giai_trinh", "label": "Nguồn/giải trình", "col": "K", "type": F_TEXTAREA,
             "help": "Căn cứ, cách làm tròn, lý do kết luận. Bắt buộc."},
            {"name": "trang_thai", "label": "Trạng thái", "col": "L", "type": F_SELECT, "options": C3_TRANG_THAI_OPTIONS, "required": True,
             "help": "Chỉ \"Đã kiểm chứng\" mới chuyển sang được D2 (phương án biên chế)."},
        ],
    },
    {
        "key": "d1",
        "label": "D1 – Nhân lực",
        "sheet_name": "9. D1_NhanLuc",
        "model_name": "NhanLuc",
        "step": 8,
        "step_note": "Nguồn lực hiện có, mỗi dòng một người. Cần danh sách nhân sự thật, chốt tại 30/9/2026.",
        "fields": [
            {"name": "ma_nguon_luc", "label": "Mã nguồn lực", "col": "A", "type": F_TEXT, "required": True,
             "help": "Duy nhất trong biểu. Trùng mã báo \"TRÙNG MÃ NGUỒN LỰC\"."},
            {"name": "ma_vtvl", "label": "Mã VTVL", "col": "B", "type": F_REF, "ref": "c1", "required": True,
             "help": "Trùng đúng mã ở C1."},
            {"name": "loai_nhan_luc", "label": "Loại nhân lực", "col": "C", "type": F_TEXT,
             "help": "Cùng cách ghi với C1 (Công chức, Viên chức...)."},
            {"name": "gio_danh_cho_vtvl", "label": "Giờ dành cho VTVL", "col": "E", "type": F_NUMBER,
             "help": "Số giờ trong năm người này làm cho VTVL, số ≥ 0."},
            {"name": "gio_viec_khac", "label": "Giờ việc khác", "col": "F", "type": F_NUMBER,
             "help": "Số giờ trong năm làm việc khác ngoài VTVL. Tổng với Giờ dành cho VTVL không vượt giờ quy đổi/năm."},
            {"name": "noi_tinh_bc", "label": "Nơi tính BC", "col": "G", "type": F_TEXT,
             "help": "Đơn vị tính biên chế/báo cáo. Bắt buộc, trống báo \"THIẾU NƠI TÍNH BIÊN CHẾ\"."},
            {"name": "nguon", "label": "Nguồn", "col": "H", "type": F_TEXT,
             "help": "Nguồn dữ liệu và thời điểm chốt. Bắt buộc."},
            {"name": "trang_thai", "label": "Trạng thái", "col": "I", "type": F_SELECT_SOFT, "options": D1_TRANG_THAI_OPTIONS,
             "help": "Gõ đúng \"Đã kiểm chứng\" khi số liệu đã được xác nhận. Khác đi sẽ báo \"Chưa đủ căn cứ\"."},
            {"name": "gio_rieng_theo_che_do", "label": "Giờ riêng theo chế độ", "col": "K", "type": F_NUMBER,
             "help": "Chỉ điền khi người này có quỹ giờ riêng khác quỹ giờ của VTVL (số > 0)."},
        ],
    },
    {
        "key": "d2",
        "label": "D2 – Phương án",
        "sheet_name": "10. D2_PhuongAn",
        "model_name": "PhuongAn",
        "step": 9,
        "step_note": "Phương án bố trí từng năm 2027–2031. Chỉ có ý nghĩa khi C3 của VTVL đã \"Đã kiểm chứng\". Làm sau cùng.",
        "fields": [
            {"name": "ma_vtvl", "label": "Mã VTVL", "col": "A", "type": F_REF, "ref": "c3", "required": True,
             "help": "Trùng đúng mã ở C3."},
            {"name": "nam", "label": "Năm (2027-2031)", "col": "B", "type": F_INT, "required": True,
             "help": "Số nguyên 2027 đến 2031. Cặp (Mã VTVL, Năm) không được trùng."},
            {"name": "so_nguoi_can_bo_tri_nam", "label": "Số người cần bố trí trong năm phương án", "col": "C", "type": F_NUMBER,
             "help": "Ước tính/kết luận số người cần cho năm này (bản mẫu hiện dùng không có sheet dự báo riêng nên đây là ô nhập tay)."},
            {"name": "so_nguoi_co_the_bo_tri", "label": "Số người có thể bố trí (cùng đơn vị tính)", "col": "D", "type": F_NUMBER,
             "help": "Số người, không tính theo giờ, số ≥ 0, cùng đơn vị tính với cột trên."},
            {"name": "bien_che_giao", "label": "Biên chế được giao theo đúng cấp", "col": "E", "type": F_NUMBER,
             "help": "Số biên chế được giao theo đúng cấp có thẩm quyền, số ≥ 0. Khác với số người cần bố trí."},
            {"name": "dieu_chuyen", "label": "Điều chuyển/bố trí lại", "col": "F", "type": F_TEXT,
             "help": "Mô tả phương án điều chuyển, bố trí lại; ghi \"Không phát sinh\" nếu không có."},
            {"name": "de_nghi_dieu_chinh", "label": "Đề nghị điều chỉnh", "col": "G", "type": F_TEXT,
             "help": "Số người đề nghị tăng (+) hoặc giảm (−); không có thì nhập 0."},
            {"name": "nguon_giai_trinh", "label": "Nguồn/giải trình", "col": "H", "type": F_TEXTAREA,
             "help": "Giải trình phương án. Bắt buộc, trống báo \"THIẾU GIẢI TRÌNH\"."},
            {"name": "trang_thai", "label": "Trạng thái", "col": "I", "type": F_SELECT, "options": D2_TRANG_THAI_OPTIONS, "required": True,
             "help": "Chọn đúng giai đoạn thực tế của phương án."},
            {"name": "van_ban_tham_dinh", "label": "Văn bản thẩm định/quyết định (số, ngày, cơ quan)", "col": "N", "type": F_TEXT,
             "help": "Bắt buộc khi trạng thái là \"Đã thẩm định\" hoặc \"Đã được quyết định\"."},
        ],
    },
]

# Ghi chu: ban mau "19-10-2026_30000_dong" ma du an dang dung KHONG co sheet
# DU_BAO_5_NAM (sheet nay chi xuat hien o ban sua doi sau, vd v2_2026-10-23 nhac
# trong video huong dan). Vi vay D2_PhuongAn cot C o day la o nhap tay, khong
# phai cong thuc tu dong. Neu sau nay chuyen sang ban co DU_BAO_5_NAM, can them
# lai bieu nay va doi cot D2-C ve "Tu tinh" (bo khoi fields o tren).

# Danh muc san pham (DM_SAN_PHAM) - dung chung toan he thong, do dau moi quan ly,
# khong gan theo don vi. Khong thuoc trinh tu 9 buoc nhung nen co truoc B1 neu
# don vi da co phut chuan duoc xac nhan.
SAN_PHAM_SCHEMA = {
    "key": "san_pham",
    "label": "Danh mục sản phẩm (DM_SAN_PHAM)",
    "sheet_name": "11. DM_SAN_PHAM",
    "model_name": "SanPham",
    "step_note": "Danh mục sản phẩm chuẩn dùng để tra phút chuẩn cho B1. Nhập trước B1 nếu đã có phút chuẩn được xác nhận; không bắt buộc.",
    "fields": [
        {"name": "ma_sp", "label": "Mã sản phẩm chuẩn", "col": "A", "type": F_TEXT, "required": True,
         "help": "Duy nhất trong biểu. B1 cột \"Mã SP trong danh mục\" phải trùng đúng mã này."},
        {"name": "ten_sp", "label": "Tên/loại sản phẩm", "col": "B", "type": F_TEXT, "required": True,
         "help": "Tên sản phẩm theo văn bản xác nhận."},
        {"name": "khoi_nganh", "label": "Khối/ngành/lĩnh vực", "col": "C", "type": F_TEXT,
         "help": "Phạm vi khối, ngành, lĩnh vực áp dụng."},
        {"name": "nhom", "label": "Nhóm I–VI (nếu áp dụng)", "col": "D", "type": F_TEXT,
         "help": "Ghi nhóm theo danh mục 06 nhóm (I đến VI) hoặc \"Không phân nhóm\"."},
        {"name": "phut_sp_chuan", "label": "Phút/SP chuẩn", "col": "E", "type": F_NUMBER,
         "help": "Số phút chuẩn, số > 0. Sản phẩm chưa có định mức thì để trống."},
        {"name": "van_ban_nguon", "label": "Văn bản/nguồn xác nhận", "col": "F", "type": F_TEXT,
         "help": "Văn bản xác nhận phút chuẩn (số, ngày, cơ quan)."},
        {"name": "trang_thai", "label": "Trạng thái", "col": "G", "type": F_SELECT_SOFT, "options": DM_SAN_PHAM_TRANG_THAI_OPTIONS,
         "help": "Gõ đúng \"Đã xác nhận\" khi phút chuẩn đã được cấp có thẩm quyền xác nhận. B1 chỉ dùng phút chuẩn khi là \"Đã xác nhận\"."},
        {"name": "co_quan_xac_nhan", "label": "Cơ quan xác nhận", "col": "H", "type": F_TEXT,
         "help": "Cơ quan ban hành, xác nhận."},
        {"name": "ngay_hieu_luc", "label": "Ngày hiệu lực", "col": "I", "type": F_TEXT,
         "help": "Ngày bắt đầu áp dụng, theo dạng ngày/tháng/năm."},
        {"name": "pham_vi_ghi_chu", "label": "Phạm vi áp dụng/ghi chú", "col": "J", "type": F_TEXT,
         "help": "Phạm vi áp dụng, ngoại lệ."},
    ],
}

MAX_DATA_ROWS = 30000  # dong 3 .. 30002 trong file mau 30.000 dong
START_ROW = 3
