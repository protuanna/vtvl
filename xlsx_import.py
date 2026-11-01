import openpyxl

from schema import MAX_DATA_ROWS, START_ROW


class XlsxImportError(Exception):
    """Loi khi doc file Excel nguoi dung tai len de nhap du lieu co san."""


def extract_rows_from_xlsx(bieu, file_path):
    """Doc 1 sheet trong file .xlsx nguoi dung tai len, theo dung cot da khai bao
    trong bieu['fields'] (cot A-S nhu file mau chinh thuc). Tra ve list[dict]."""
    try:
        wb = openpyxl.load_workbook(file_path, data_only=True, read_only=True)
    except Exception as e:
        raise XlsxImportError(f"Không mở được file Excel: {e}")

    sheet_name = bieu["sheet_name"]
    if sheet_name not in wb.sheetnames:
        raise XlsxImportError(
            f"File không có sheet '{sheet_name}'. File tải lên phải theo đúng cấu trúc "
            f"file mẫu chính thức (tên sheet, cột giống hệt)."
        )
    ws = wb[sheet_name]

    rows = []
    empty_streak = 0
    for i in range(MAX_DATA_ROWS):
        excel_row = START_ROW + i
        row_data = {}
        has_value = False
        for f in bieu["fields"]:
            col = f["col"]
            cell = ws[f"{col}{excel_row}"]
            value = cell.value
            if value is not None and value != "":
                has_value = True
                if f["type"] in ("number", "int") and not isinstance(value, (int, float)):
                    value = str(value)
                row_data[f["name"]] = value
        if not has_value:
            empty_streak += 1
            if empty_streak >= 5:  # qua 5 dong trong lien tiep thi coi nhu het du lieu
                break
            continue
        empty_streak = 0
        rows.append(row_data)

    if not rows:
        raise XlsxImportError(
            f"Không tìm thấy dòng dữ liệu nào trong sheet '{sheet_name}' (dòng {START_ROW} trở đi)."
        )
    return rows
