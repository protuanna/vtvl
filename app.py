import os
import tempfile
import uuid

from flask import (
    Flask,
    abort,
    flash,
    redirect,
    render_template,
    request,
    send_file,
    url_for,
)

from models import DonVi, db
from registry import BIEU_BY_KEY, BIEU_LIST, SAN_PHAM_ENTRY
from export import ExportError, export_workbook
from ai_import import ImportError_, extract_rows, load_source
from xlsx_import import XlsxImportError, extract_rows_from_xlsx

# Luu tam ket qua AI trich xuat cho den khi nguoi dung xac nhan/huy (bo nho tien trinh, du cho 1 app local).
IMPORT_STAGING = {}

BASE_DIR = os.path.dirname(__file__)

# Uu tien DATABASE_URL tu bien moi truong (Postgres/Supabase khi chay tren Vercel).
# Khong co thi fallback ve SQLite local cho de chay/dev tren may.
database_url = os.environ.get("DATABASE_URL")
if database_url:
    # Supabase/Heroku hay tra ve prefix "postgres://", SQLAlchemy 2.x can "postgresql://".
    if database_url.startswith("postgres://"):
        database_url = database_url.replace("postgres://", "postgresql://", 1)
    # Chi dinh ro driver psycopg2 (da cai qua requirements.txt) de tranh SQLAlchemy
    # tu dong thu driver "psycopg" (v3, chua cai) truoc khi fallback.
    if database_url.startswith("postgresql://"):
        database_url = database_url.replace("postgresql://", "postgresql+psycopg2://", 1)
else:
    os.makedirs(os.path.join(BASE_DIR, "instance"), exist_ok=True)
    database_url = "sqlite:///" + os.path.join(BASE_DIR, "instance", "vtvl.db")

app = Flask(__name__)
app.config["SQLALCHEMY_DATABASE_URI"] = database_url
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
app.secret_key = os.environ.get("SECRET_KEY", "vtvl-tool-local-dev")  # dat SECRET_KEY rieng khi deploy that

db.init_app(app)
app.jinja_env.globals["getattr"] = getattr


def parse_field_value(field, raw):
    raw = (raw or "").strip()
    if raw == "":
        return None
    if field["type"] == "number":
        return float(raw.replace(",", "."))
    if field["type"] == "int":
        return int(raw)
    return raw


def bieu_or_404(bieu_key):
    bieu = BIEU_BY_KEY.get(bieu_key)
    if not bieu:
        abort(404)
    return bieu


def get_ref_options(bieu, don_vi_id):
    """Voi moi truong kieu F_REF trong bieu, tra ve danh sach ma da co san de goi y
    (lay tu bieu nguon duoc tro toi qua "ref"; DM_SAN_PHAM la danh muc dung chung)."""
    options = {}
    for f in bieu["fields"]:
        if f["type"] != "ref":
            continue
        ref_key = f["ref"]
        ref_bieu = SAN_PHAM_ENTRY if ref_key == "san_pham" else BIEU_BY_KEY.get(ref_key)
        if not ref_bieu:
            options[f["name"]] = []
            continue
        code_field = ref_bieu["fields"][0]["name"]
        model = ref_bieu["model"]
        query = model.query
        if ref_key != "san_pham":
            query = query.filter_by(don_vi_id=don_vi_id)
        values = [getattr(r, code_field) for r in query.all()]
        options[f["name"]] = sorted({v for v in values if v})
    return options


def fill_row_from_form(row_obj, fields, form):
    for f in fields:
        try:
            value = parse_field_value(f, form.get(f["name"]))
        except ValueError:
            raise ValueError(f"Giá trị của trường '{f['label']}' không hợp lệ.")
        if f.get("required") and value is None:
            raise ValueError(f"Trường '{f['label']}' là bắt buộc.")
        setattr(row_obj, f["name"], value)


# ---------- Trang chu: quan ly don vi ----------

@app.route("/")
def index():
    don_vis = DonVi.query.order_by(DonVi.ma_co_quan).all()
    return render_template("index.html", don_vis=don_vis)


@app.route("/don-vi/them", methods=["POST"])
def don_vi_them():
    ma = (request.form.get("ma_co_quan") or "").strip()
    ten = (request.form.get("ten_co_quan") or "").strip()
    if not ma or not ten:
        flash("Vui lòng nhập đủ Mã cơ quan và Tên cơ quan.", "error")
        return redirect(url_for("index"))
    if DonVi.query.filter_by(ma_co_quan=ma).first():
        flash(f"Mã cơ quan '{ma}' đã tồn tại.", "error")
        return redirect(url_for("index"))
    db.session.add(DonVi(ma_co_quan=ma, ten_co_quan=ten))
    db.session.commit()
    flash(f"Đã tạo đơn vị '{ten}'.", "ok")
    return redirect(url_for("index"))


@app.route("/don-vi/<int:don_vi_id>/xoa", methods=["POST"])
def don_vi_xoa(don_vi_id):
    don_vi = DonVi.query.get_or_404(don_vi_id)
    for bieu in BIEU_LIST:
        bieu["model"].query.filter_by(don_vi_id=don_vi_id).delete()
    db.session.delete(don_vi)
    db.session.commit()
    flash(f"Đã xóa đơn vị '{don_vi.ten_co_quan}' và toàn bộ dữ liệu liên quan.", "ok")
    return redirect(url_for("index"))


@app.route("/don-vi/<int:don_vi_id>")
def don_vi_dashboard(don_vi_id):
    don_vi = DonVi.query.get_or_404(don_vi_id)
    counts = {b["key"]: b["model"].query.filter_by(don_vi_id=don_vi_id).count() for b in BIEU_LIST}
    return render_template("don_vi_dashboard.html", don_vi=don_vi, bieu_list=BIEU_LIST, counts=counts)


@app.route("/don-vi/<int:don_vi_id>/xuat")
def don_vi_xuat(don_vi_id):
    DonVi.query.get_or_404(don_vi_id)
    try:
        file_path, filename = export_workbook([don_vi_id])
    except ExportError as e:
        flash(str(e), "error")
        return redirect(url_for("don_vi_dashboard", don_vi_id=don_vi_id))
    return send_file(file_path, as_attachment=True, download_name=filename)


# ---------- CRUD chung cho 9 bieu, pham vi theo 1 don vi ----------

@app.route("/don-vi/<int:don_vi_id>/<bieu_key>")
def bieu_list_view(don_vi_id, bieu_key):
    don_vi = DonVi.query.get_or_404(don_vi_id)
    bieu = bieu_or_404(bieu_key)
    rows = bieu["model"].query.filter_by(don_vi_id=don_vi_id).order_by(bieu["model"].id).all()
    return render_template(
        "bieu_list.html", don_vi=don_vi, bieu=bieu, rows=rows,
        action_url=url_for("bieu_them", don_vi_id=don_vi_id, bieu_key=bieu_key),
        edit_url_name="bieu_sua", delete_url_name="bieu_xoa", don_vi_id=don_vi_id,
        ref_options=get_ref_options(bieu, don_vi_id),
    )


@app.route("/don-vi/<int:don_vi_id>/<bieu_key>/them", methods=["POST"])
def bieu_them(don_vi_id, bieu_key):
    DonVi.query.get_or_404(don_vi_id)
    bieu = bieu_or_404(bieu_key)
    row_obj = bieu["model"](don_vi_id=don_vi_id)
    try:
        fill_row_from_form(row_obj, bieu["fields"], request.form)
    except ValueError as e:
        flash(str(e), "error")
        return redirect(url_for("bieu_list_view", don_vi_id=don_vi_id, bieu_key=bieu_key))
    db.session.add(row_obj)
    db.session.commit()
    flash("Đã thêm dòng mới.", "ok")
    return redirect(url_for("bieu_list_view", don_vi_id=don_vi_id, bieu_key=bieu_key))


@app.route("/don-vi/<int:don_vi_id>/<bieu_key>/<int:row_id>/sua", methods=["GET", "POST"])
def bieu_sua(don_vi_id, bieu_key, row_id):
    don_vi = DonVi.query.get_or_404(don_vi_id)
    bieu = bieu_or_404(bieu_key)
    row_obj = bieu["model"].query.filter_by(id=row_id, don_vi_id=don_vi_id).first_or_404()
    if request.method == "POST":
        try:
            fill_row_from_form(row_obj, bieu["fields"], request.form)
        except ValueError as e:
            flash(str(e), "error")
            return redirect(url_for("bieu_sua", don_vi_id=don_vi_id, bieu_key=bieu_key, row_id=row_id))
        db.session.commit()
        flash("Đã lưu thay đổi.", "ok")
        return redirect(url_for("bieu_list_view", don_vi_id=don_vi_id, bieu_key=bieu_key))
    return render_template(
        "bieu_edit.html", don_vi=don_vi, bieu=bieu, row=row_obj,
        back_url=url_for("bieu_list_view", don_vi_id=don_vi_id, bieu_key=bieu_key),
        ref_options=get_ref_options(bieu, don_vi_id),
    )


@app.route("/don-vi/<int:don_vi_id>/<bieu_key>/<int:row_id>/xoa", methods=["POST"])
def bieu_xoa(don_vi_id, bieu_key, row_id):
    bieu = bieu_or_404(bieu_key)
    row_obj = bieu["model"].query.filter_by(id=row_id, don_vi_id=don_vi_id).first_or_404()
    db.session.delete(row_obj)
    db.session.commit()
    flash("Đã xóa dòng.", "ok")
    return redirect(url_for("bieu_list_view", don_vi_id=don_vi_id, bieu_key=bieu_key))


# ---------- Nhap du lieu tu tai lieu (AI doc PDF/DOCX, bao gom ca ban scan) ----------

@app.route("/don-vi/<int:don_vi_id>/<bieu_key>/import-ai", methods=["GET", "POST"])
def bieu_import_ai(don_vi_id, bieu_key):
    don_vi = DonVi.query.get_or_404(don_vi_id)
    bieu = bieu_or_404(bieu_key)

    if request.method == "GET":
        return render_template("import_ai_upload.html", don_vi=don_vi, bieu=bieu)

    uploaded = request.files.get("file")
    if not uploaded or not uploaded.filename:
        flash("Vui lòng chọn một file PDF hoặc DOCX.", "error")
        return redirect(url_for("bieu_import_ai", don_vi_id=don_vi_id, bieu_key=bieu_key))

    suffix = os.path.splitext(uploaded.filename)[1].lower()
    tmp_dir = tempfile.mkdtemp(prefix="vtvl_import_")
    tmp_path = os.path.join(tmp_dir, "source" + suffix)
    uploaded.save(tmp_path)

    try:
        source = load_source(tmp_path)
        rows = extract_rows(bieu, source)
    except ImportError_ as e:
        flash(str(e), "error")
        return redirect(url_for("bieu_import_ai", don_vi_id=don_vi_id, bieu_key=bieu_key))

    if not rows:
        flash("AI không trích xuất được dòng dữ liệu nào phù hợp từ file này.", "error")
        return redirect(url_for("bieu_import_ai", don_vi_id=don_vi_id, bieu_key=bieu_key))

    staging_id = uuid.uuid4().hex
    IMPORT_STAGING[staging_id] = {
        "don_vi_id": don_vi_id,
        "bieu_key": bieu_key,
        "rows": rows,
        "note": source.get("note", ""),
        "source_name": uploaded.filename,
    }
    return redirect(url_for("bieu_import_ai_review", don_vi_id=don_vi_id, bieu_key=bieu_key, staging_id=staging_id))


@app.route("/don-vi/<int:don_vi_id>/<bieu_key>/import-ai/<staging_id>", methods=["GET", "POST"])
def bieu_import_ai_review(don_vi_id, bieu_key, staging_id):
    don_vi = DonVi.query.get_or_404(don_vi_id)
    bieu = bieu_or_404(bieu_key)
    staged = IMPORT_STAGING.get(staging_id)
    if not staged or staged["don_vi_id"] != don_vi_id or staged["bieu_key"] != bieu_key:
        flash("Phiên nhập liệu đã hết hạn hoặc không hợp lệ. Vui lòng tải file lại.", "error")
        return redirect(url_for("bieu_import_ai", don_vi_id=don_vi_id, bieu_key=bieu_key))

    if request.method == "POST":
        imported = 0
        for idx in range(len(staged["rows"])):
            if not request.form.get(f"include_{idx}"):
                continue
            row_obj = bieu["model"](don_vi_id=don_vi_id)
            try:
                fill_row_from_form(
                    row_obj, bieu["fields"],
                    {f["name"]: request.form.get(f"row_{idx}_{f['name']}") for f in bieu["fields"]},
                )
            except ValueError as e:
                flash(f"Dòng {idx + 1}: {e}", "error")
                continue
            db.session.add(row_obj)
            imported += 1
        db.session.commit()
        IMPORT_STAGING.pop(staging_id, None)
        flash(f"Đã nhập {imported} dòng vào {bieu['label']}.", "ok")
        return redirect(url_for("bieu_list_view", don_vi_id=don_vi_id, bieu_key=bieu_key))

    return render_template(
        "import_ai_review.html", don_vi=don_vi, bieu=bieu, staging_id=staging_id,
        rows=staged["rows"], note=staged["note"], source_name=staged["source_name"],
        ref_options=get_ref_options(bieu, don_vi_id),
    )


# ---------- Nhap tu file Excel co san (khach da lam truoc theo dung mau) ----------

@app.route("/don-vi/<int:don_vi_id>/<bieu_key>/import-excel", methods=["GET", "POST"])
def bieu_import_excel(don_vi_id, bieu_key):
    don_vi = DonVi.query.get_or_404(don_vi_id)
    bieu = bieu_or_404(bieu_key)

    if request.method == "GET":
        return render_template("import_excel_upload.html", don_vi=don_vi, bieu=bieu)

    uploaded = request.files.get("file")
    if not uploaded or not uploaded.filename or not uploaded.filename.lower().endswith(".xlsx"):
        flash("Vui lòng chọn một file Excel (.xlsx).", "error")
        return redirect(url_for("bieu_import_excel", don_vi_id=don_vi_id, bieu_key=bieu_key))

    tmp_dir = tempfile.mkdtemp(prefix="vtvl_xlsx_import_")
    tmp_path = os.path.join(tmp_dir, "source.xlsx")
    uploaded.save(tmp_path)

    try:
        rows = extract_rows_from_xlsx(bieu, tmp_path)
    except XlsxImportError as e:
        flash(str(e), "error")
        return redirect(url_for("bieu_import_excel", don_vi_id=don_vi_id, bieu_key=bieu_key))

    staging_id = uuid.uuid4().hex
    IMPORT_STAGING[staging_id] = {
        "don_vi_id": don_vi_id,
        "bieu_key": bieu_key,
        "rows": rows,
        "note": f"Đọc trực tiếp từ sheet '{bieu['sheet_name']}' của file tải lên (không qua AI).",
        "source_name": uploaded.filename,
    }
    return redirect(url_for("bieu_import_ai_review", don_vi_id=don_vi_id, bieu_key=bieu_key, staging_id=staging_id))


@app.route("/don-vi/<int:don_vi_id>/import-excel", methods=["GET", "POST"])
def don_vi_import_excel(don_vi_id):
    don_vi = DonVi.query.get_or_404(don_vi_id)

    if request.method == "GET":
        return render_template("import_excel_upload.html", don_vi=don_vi)

    uploaded = request.files.get("file")
    if not uploaded or not uploaded.filename:
        flash("Vui lòng chọn một file Excel (.xlsx).", "error")
        return redirect(url_for("don_vi_import_excel", don_vi_id=don_vi_id))
    if not uploaded.filename.lower().endswith(".xlsx"):
        flash("Chỉ hỗ trợ file .xlsx theo đúng cấu trúc file mẫu chính thức.", "error")
        return redirect(url_for("don_vi_import_excel", don_vi_id=don_vi_id))

    tmp_dir = tempfile.mkdtemp(prefix="vtvl_xlsx_import_")
    tmp_path = os.path.join(tmp_dir, "source.xlsx")
    uploaded.save(tmp_path)

    results = []  # (bieu, status, count_or_message, staging_id)
    for bieu in BIEU_LIST:
        try:
            rows = extract_rows_from_xlsx(bieu, tmp_path)
        except XlsxImportError as e:
            results.append({"bieu": bieu, "ok": False, "message": str(e), "staging_id": None})
            continue
        staging_id = uuid.uuid4().hex
        IMPORT_STAGING[staging_id] = {
            "don_vi_id": don_vi_id,
            "bieu_key": bieu["key"],
            "rows": rows,
            "note": f"Đọc trực tiếp từ sheet '{bieu['sheet_name']}' của file tải lên (không qua AI).",
            "source_name": uploaded.filename,
        }
        results.append({"bieu": bieu, "ok": True, "count": len(rows), "staging_id": staging_id})

    if not any(r["ok"] for r in results):
        flash("Không đọc được dữ liệu từ bất kỳ sheet nào trong file. Kiểm tra lại file có đúng cấu trúc mẫu không.", "error")
        return redirect(url_for("don_vi_import_excel", don_vi_id=don_vi_id))

    return render_template("import_excel_summary.html", don_vi=don_vi, results=results, source_name=uploaded.filename)


# ---------- Danh muc san pham (dung chung, khong gan don vi) ----------

@app.route("/danh-muc-san-pham")
def san_pham_list_view():
    bieu = SAN_PHAM_ENTRY
    rows = bieu["model"].query.order_by(bieu["model"].id).all()
    return render_template("san_pham_list.html", bieu=bieu, rows=rows, ref_options={})


@app.route("/danh-muc-san-pham/them", methods=["POST"])
def san_pham_them():
    bieu = SAN_PHAM_ENTRY
    row_obj = bieu["model"]()
    try:
        fill_row_from_form(row_obj, bieu["fields"], request.form)
    except ValueError as e:
        flash(str(e), "error")
        return redirect(url_for("san_pham_list_view"))
    db.session.add(row_obj)
    db.session.commit()
    flash("Đã thêm sản phẩm.", "ok")
    return redirect(url_for("san_pham_list_view"))


@app.route("/danh-muc-san-pham/<int:row_id>/sua", methods=["GET", "POST"])
def san_pham_sua(row_id):
    bieu = SAN_PHAM_ENTRY
    row_obj = bieu["model"].query.get_or_404(row_id)
    if request.method == "POST":
        try:
            fill_row_from_form(row_obj, bieu["fields"], request.form)
        except ValueError as e:
            flash(str(e), "error")
            return redirect(url_for("san_pham_sua", row_id=row_id))
        db.session.commit()
        flash("Đã lưu thay đổi.", "ok")
        return redirect(url_for("san_pham_list_view"))
    return render_template("san_pham_edit.html", bieu=bieu, row=row_obj, ref_options={})


@app.route("/danh-muc-san-pham/<int:row_id>/xoa", methods=["POST"])
def san_pham_xoa(row_id):
    bieu = SAN_PHAM_ENTRY
    row_obj = bieu["model"].query.get_or_404(row_id)
    db.session.delete(row_obj)
    db.session.commit()
    flash("Đã xóa sản phẩm.", "ok")
    return redirect(url_for("san_pham_list_view"))


# ---------- Tong hop nhieu don vi va xuat Excel chung ----------

@app.route("/tong-hop", methods=["GET", "POST"])
def tong_hop():
    don_vis = DonVi.query.order_by(DonVi.ma_co_quan).all()
    if request.method == "POST":
        selected_ids = [int(x) for x in request.form.getlist("don_vi_ids")]
        if not selected_ids:
            flash("Vui lòng chọn ít nhất một đơn vị.", "error")
            return redirect(url_for("tong_hop"))
        try:
            file_path, filename = export_workbook(selected_ids)
        except ExportError as e:
            flash(str(e), "error")
            return redirect(url_for("tong_hop"))
        return send_file(file_path, as_attachment=True, download_name=filename)
    return render_template("tong_hop.html", don_vis=don_vis)


# Tao bang neu chua co. Goi o day (khong chi trong __main__) de chay duoc khi Vercel
# import thang module "app" lam ham serverless (khong qua "python app.py").
with app.app_context():
    db.create_all()

if __name__ == "__main__":
    app.run(debug=True, port=8000, use_reloader=False)
