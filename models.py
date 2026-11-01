from flask_sqlalchemy import SQLAlchemy

db = SQLAlchemy()


class DonVi(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    ma_co_quan = db.Column(db.String(64), unique=True, nullable=False)
    ten_co_quan = db.Column(db.String(255), nullable=False)


class NhiemVu(db.Model):  # Bieu A1
    id = db.Column(db.Integer, primary_key=True)
    don_vi_id = db.Column(db.Integer, db.ForeignKey("don_vi.id"), nullable=False)
    ma_nhiem_vu = db.Column(db.String(64))
    ten_nhiem_vu = db.Column(db.String(500))
    co_quan = db.Column(db.String(255))
    can_cu = db.Column(db.String(500))
    ky = db.Column(db.String(100))
    trang_thai = db.Column(db.String(100))
    nguon = db.Column(db.String(500))


class CongViec(db.Model):  # Bieu A2
    id = db.Column(db.Integer, primary_key=True)
    don_vi_id = db.Column(db.Integer, db.ForeignKey("don_vi.id"), nullable=False)
    ma_cong_viec = db.Column(db.String(64))
    ma_nhiem_vu = db.Column(db.String(64))
    ma_vtvl = db.Column(db.String(64))
    ten_cong_viec = db.Column(db.String(500))
    san_pham = db.Column(db.String(500))
    don_vi_tinh = db.Column(db.String(100))
    chat_luong = db.Column(db.String(255))
    can_cu_nguon = db.Column(db.String(500))


class KhoiLuong(db.Model):  # Bieu B1
    id = db.Column(db.Integer, primary_key=True)
    don_vi_id = db.Column(db.Integer, db.ForeignKey("don_vi.id"), nullable=False)
    ma_cong_viec = db.Column(db.String(64))
    so_sp_nam = db.Column(db.Float)
    tb_phut_sp = db.Column(db.Float)
    gio_bo_sung = db.Column(db.Float)
    nguon_sl = db.Column(db.String(500))
    nguon_thoi_gian = db.Column(db.String(500))
    trang_thai_can_cu_tg = db.Column(db.String(255))
    ky = db.Column(db.String(100))
    ma_sp_danh_muc = db.Column(db.String(64))
    giai_trinh_khac_chuan = db.Column(db.String(500))


class TacDong(db.Model):  # Bieu B2_CT_TacDong
    id = db.Column(db.Integer, primary_key=True)
    don_vi_id = db.Column(db.Integer, db.ForeignKey("don_vi.id"), nullable=False)
    ma_cong_viec = db.Column(db.String(64))
    loai_tac_dong = db.Column(db.String(255))
    cong_doan = db.Column(db.String(500))
    sp_ap_dung = db.Column(db.String(255))
    phut_truoc = db.Column(db.Float)
    phut_sau = db.Column(db.Float)
    gio_viec_moi = db.Column(db.Float)
    nguon_bang_chung = db.Column(db.String(500))
    ky = db.Column(db.String(100))
    co_so_b1 = db.Column(db.String(255))
    trang_thai = db.Column(db.String(100))
    kinh_phi_truoc = db.Column(db.Float)
    kinh_phi_sau = db.Column(db.Float)
    doi_chieu_trung = db.Column(db.String(500))


class Vtvl(db.Model):  # Bieu C1
    id = db.Column(db.Integer, primary_key=True)
    don_vi_id = db.Column(db.Integer, db.ForeignKey("don_vi.id"), nullable=False)
    ma_vtvl = db.Column(db.String(64))
    ten_vtvl = db.Column(db.String(500))
    loai_nhan_luc = db.Column(db.String(255))
    quyet_dinh = db.Column(db.String(500))
    can_cu_quy_gio = db.Column(db.String(500))
    trang_thai = db.Column(db.String(100))
    gio_rieng_xac_nhan = db.Column(db.Float)
    xac_nhan_gio_rieng = db.Column(db.String(100))
    che_do_lam_viec = db.Column(db.String(100))


class NangLuc(db.Model):  # Bieu C2
    id = db.Column(db.Integer, primary_key=True)
    don_vi_id = db.Column(db.Integer, db.ForeignKey("don_vi.id"), nullable=False)
    ma_vtvl = db.Column(db.String(64))
    trach_nhiem = db.Column(db.Text)
    nang_luc_chung = db.Column(db.Text)
    chuyen_mon = db.Column(db.Text)
    lanh_dao = db.Column(db.Text)
    can_cu = db.Column(db.String(500))
    nguon_mo_ta = db.Column(db.String(500))


class NhuCau(db.Model):  # Bieu C3
    id = db.Column(db.Integer, primary_key=True)
    don_vi_id = db.Column(db.Integer, db.ForeignKey("don_vi.id"), nullable=False)
    ma_vtvl = db.Column(db.String(64))
    phuong_phap = db.Column(db.String(10))
    nguoi_pp1 = db.Column(db.Float)
    nguoi_pp3 = db.Column(db.Float)
    nguoi_can_bo_tri = db.Column(db.Float)
    nguon_giai_trinh = db.Column(db.String(500))
    trang_thai = db.Column(db.String(100))


class NhanLuc(db.Model):  # Bieu D1
    id = db.Column(db.Integer, primary_key=True)
    don_vi_id = db.Column(db.Integer, db.ForeignKey("don_vi.id"), nullable=False)
    ma_nguon_luc = db.Column(db.String(64))
    ma_vtvl = db.Column(db.String(64))
    loai_nhan_luc = db.Column(db.String(255))
    gio_danh_cho_vtvl = db.Column(db.Float)
    gio_viec_khac = db.Column(db.Float)
    noi_tinh_bc = db.Column(db.String(255))
    nguon = db.Column(db.String(500))
    trang_thai = db.Column(db.String(100))
    gio_rieng_theo_che_do = db.Column(db.Float)


class PhuongAn(db.Model):  # Bieu D2
    id = db.Column(db.Integer, primary_key=True)
    don_vi_id = db.Column(db.Integer, db.ForeignKey("don_vi.id"), nullable=False)
    ma_vtvl = db.Column(db.String(64))
    nam = db.Column(db.Integer)
    so_nguoi_can_bo_tri_nam = db.Column(db.Float)
    so_nguoi_co_the_bo_tri = db.Column(db.Float)
    bien_che_giao = db.Column(db.Float)
    dieu_chuyen = db.Column(db.String(500))
    de_nghi_dieu_chinh = db.Column(db.String(500))
    nguon_giai_trinh = db.Column(db.String(500))
    trang_thai = db.Column(db.String(100))
    van_ban_tham_dinh = db.Column(db.String(500))


class SanPham(db.Model):  # DM_SAN_PHAM - danh muc dung chung, khong gan don vi
    id = db.Column(db.Integer, primary_key=True)
    ma_sp = db.Column(db.String(64))
    ten_sp = db.Column(db.String(500))
    khoi_nganh = db.Column(db.String(255))
    nhom = db.Column(db.String(50))
    phut_sp_chuan = db.Column(db.Float)
    van_ban_nguon = db.Column(db.String(500))
    trang_thai = db.Column(db.String(100))
    co_quan_xac_nhan = db.Column(db.String(255))
    ngay_hieu_luc = db.Column(db.String(50))
    pham_vi_ghi_chu = db.Column(db.String(500))
