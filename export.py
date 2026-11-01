import datetime
import os
import re
import tempfile
import zipfile
from xml.sax.saxutils import escape as xml_escape

from models import DonVi, SanPham
from registry import BIEU_LIST, SAN_PHAM_ENTRY
from schema import MAX_DATA_ROWS, START_ROW

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "template_xlsx", "03. Bo bieu VTVL 19-10-2026_30000_dong.xlsx")

# File mau vat ly hoa san 30.000 dong voi cong thuc "shared" cho moi sheet nghiep
# vu (vai chuc MB/sheet). Thu vien openpyxl phai parse toan bo cay cong thuc khi
# doc va *giai nen* cong thuc shared thanh van ban day du khi ghi lai, khien thao
# tac mat 2-3 phut va phinh file gap ~3 lan. Vi vay o day ta thao tac truc tiep
# tren XML ben trong file .xlsx (zip): chi tim va thay cac the <c> cua dung cac
# o nhap lieu can ghi, giu nguyen 100% moi o cong thuc va moi sheet khac. Nhanh
# (duoi 1 giay) va khong lam phinh file.

ROW_RE = re.compile(r'<row r="(\d+)"[^>]*(?:/>|>.*?</row>)', re.DOTALL)
CELL_RE_TEMPLATE = r'<c r="{col}{row}"([^>]*?)(?:/>|>(.*?)</c>)'


class ExportError(Exception):
    """Loi nghiep vu khi xuat Excel (vuot gioi han dong, thieu du lieu...)."""


def _col_row_rel_path(zf):
    """Doc workbook.xml + rels de map ten sheet logic -> duong dan file XML trong zip."""
    wb_xml = zf.read("xl/workbook.xml").decode("utf-8")
    rels_xml = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8")

    name_to_rid = dict(re.findall(r'<sheet name="([^"]+)"[^>]*r:id="(rId\d+)"', wb_xml))
    rid_to_target = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="(worksheets/sheet\d+\.xml)"', rels_xml))

    name_to_path = {}
    for name, rid in name_to_rid.items():
        target = rid_to_target.get(rid)
        if target:
            name_to_path[name] = "xl/" + target
    return name_to_path


def _cell_xml(col, row, style_attrs, value):
    ref = f"{col}{row}"
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return f'<c r="{ref}"{style_attrs}><v>{value}</v></c>'
    text = xml_escape(str(value))
    return f'<c r="{ref}"{style_attrs} t="inlineStr"><is><t xml:space="preserve">{text}</t></is></c>'


def _apply_writes_to_sheet_xml(xml_text, writes):
    """writes: {row_num(int): {col_letter(str): value}}. Tra ve xml_text da sua."""
    if not writes:
        return xml_text

    parts = []
    last_end = 0
    seen_rows = set()

    for m in ROW_RE.finditer(xml_text):
        row_num = int(m.group(1))
        parts.append(xml_text[last_end:m.start()])
        row_text = m.group(0)
        if row_num in writes:
            seen_rows.add(row_num)
            for col, value in writes[row_num].items():
                cell_re = re.compile(CELL_RE_TEMPLATE.format(col=col, row=row_num), re.DOTALL)
                cm = cell_re.search(row_text)
                if not cm:
                    raise ExportError(
                        f"Không tìm thấy ô {col}{row_num} trong file mẫu để ghi dữ liệu. "
                        f"File mẫu có thể đã bị thay đổi cấu trúc."
                    )
                style_attrs = cm.group(1) or ""
                new_cell = _cell_xml(col, row_num, style_attrs, value)
                row_text = row_text[: cm.start()] + new_cell + row_text[cm.end() :]
            parts.append(row_text)
        else:
            parts.append(row_text)
        last_end = m.end()

    parts.append(xml_text[last_end:])

    missing = set(writes.keys()) - seen_rows
    if missing:
        raise ExportError(
            f"Không tìm thấy dòng {sorted(missing)[:5]}... trong sheet để ghi dữ liệu "
            f"(vượt quá vùng dòng có sẵn của file mẫu)."
        )
    return "".join(parts)


def _rows_to_writes(fields, rows):
    writes = {}
    for i, row_obj in enumerate(rows):
        excel_row = START_ROW + i
        cells = {}
        for f in fields:
            value = getattr(row_obj, f["name"])
            if value is None or value == "":
                continue
            cells[f["col"]] = value
        if cells:
            writes[excel_row] = cells
    return writes


def export_workbook(don_vi_ids):
    """Xuat file Excel dung chuan mau 03.Bo bieu VTVL cho mot hoac nhieu don vi.

    don_vi_ids: danh sach id DonVi theo dung thu tu muon ghep noi trong file xuat ra.
    Tra ve (file_path, download_filename).
    """
    if not don_vi_ids:
        raise ExportError("Chưa chọn đơn vị nào để xuất.")

    with zipfile.ZipFile(TEMPLATE_PATH, "r") as zin:
        name_to_path = _col_row_rel_path(zin)
        sheet_xml_cache = {}

        def get_sheet_xml(sheet_name):
            path = name_to_path.get(sheet_name)
            if not path:
                raise ExportError(f"Không tìm thấy sheet '{sheet_name}' trong file mẫu.")
            if path not in sheet_xml_cache:
                sheet_xml_cache[path] = zin.read(path).decode("utf-8")
            return path, sheet_xml_cache[path]

        for bieu in BIEU_LIST:
            model = bieu["model"]
            rows = []
            for don_vi_id in don_vi_ids:
                rows.extend(
                    model.query.filter_by(don_vi_id=don_vi_id).order_by(model.id).all()
                )
            max_rows = bieu.get("max_rows", MAX_DATA_ROWS)
            if len(rows) > max_rows:
                raise ExportError(
                    f"Biểu '{bieu['label']}' có {len(rows)} dòng dữ liệu, vượt quá giới hạn "
                    f"{max_rows} dòng (dòng {START_ROW}-{START_ROW + max_rows - 1}) của file mẫu. "
                    f"Theo cẩm nang: khi vượt số dòng phải báo cơ quan đầu mối mở rộng file có "
                    f"kiểm thử, không tự chèn dòng."
                )
            writes = _rows_to_writes(bieu["fields"], rows)
            path, xml_text = get_sheet_xml(bieu["sheet_name"])
            sheet_xml_cache[path] = _apply_writes_to_sheet_xml(xml_text, writes)

        # Danh muc san pham la du lieu dung chung, luon ghi toan bo khong phan theo don vi.
        sp_rows = SanPham.query.order_by(SanPham.id).all()
        max_rows = SAN_PHAM_ENTRY.get("max_rows", MAX_DATA_ROWS)
        if len(sp_rows) > max_rows:
            raise ExportError(
                f"Danh mục sản phẩm có {len(sp_rows)} dòng, vượt quá giới hạn {max_rows} dòng."
            )
        sp_writes = _rows_to_writes(SAN_PHAM_ENTRY["fields"], sp_rows)
        path, xml_text = get_sheet_xml(SAN_PHAM_ENTRY["sheet_name"])
        sheet_xml_cache[path] = _apply_writes_to_sheet_xml(xml_text, sp_writes)

        today = datetime.date.today().strftime("%Y%m%d")
        if len(don_vi_ids) == 1:
            don_vi = DonVi.query.get(don_vi_ids[0])
            safe_name = (don_vi.ma_co_quan if don_vi else str(don_vi_ids[0])).replace(" ", "_")
            download_filename = f"BoBieuVTVL_{safe_name}_{today}.xlsx"
        else:
            download_filename = f"BoBieuVTVL_TongHop_{today}.xlsx"

        tmp_dir = tempfile.mkdtemp(prefix="vtvl_export_")
        file_path = os.path.join(tmp_dir, download_filename)

        with zipfile.ZipFile(file_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = sheet_xml_cache[item.filename].encode("utf-8") if item.filename in sheet_xml_cache else zin.read(item.filename)
                zout.writestr(item, data)

    return file_path, download_filename
