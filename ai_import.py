import base64
import json
import os
import re

import fitz  # pymupdf
from docx import Document
from pypdf import PdfReader

MAX_IMAGE_PAGES = 20  # gioi han so trang scan gui qua API (chi phi + kich thuoc request)
MIN_CHARS_PER_PAGE = 200  # duoi nguong nay coi la file scan, khong co lop van ban
MAX_TEXT_CHARS = 60000  # cat bot neu van ban qua dai

ANTHROPIC_MODEL = os.environ.get("ANTHROPIC_MODEL", "claude-sonnet-4-5")


class ImportError_(Exception):
    """Loi trong qua trinh trich xuat/import tai lieu."""


def _extract_pdf_text(path):
    reader = PdfReader(path)
    pages_text = [(p.extract_text() or "") for p in reader.pages]
    full = "\n".join(pages_text)
    avg_per_page = len(full) / max(len(pages_text), 1)
    return full, avg_per_page, len(pages_text)


def _render_pdf_images(path, max_pages=MAX_IMAGE_PAGES):
    doc = fitz.open(path)
    images = []
    for i in range(min(len(doc), max_pages)):
        pix = doc[i].get_pixmap(dpi=170)
        images.append(pix.tobytes("png"))
    return images, len(doc)


def _extract_docx_text(path):
    doc = Document(path)
    parts = [p.text for p in doc.paragraphs if p.text.strip()]
    for table in doc.tables:
        for row in table.rows:
            parts.append(" | ".join(c.text.strip() for c in row.cells))
    return "\n".join(parts)


def load_source(file_path):
    """Doc file nguon (pdf/docx), tra ve dict {mode: 'text'|'images', text?, images?, note}."""
    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".docx":
        text = _extract_docx_text(file_path)
        if len(text.strip()) < 20:
            raise ImportError_("Không đọc được nội dung văn bản từ file .docx này.")
        return {"mode": "text", "text": text[:MAX_TEXT_CHARS], "note": f"{len(text)} ký tự từ file .docx"}

    if ext == ".pdf":
        text, avg_per_page, n_pages = _extract_pdf_text(file_path)
        if avg_per_page >= MIN_CHARS_PER_PAGE:
            return {
                "mode": "text",
                "text": text[:MAX_TEXT_CHARS],
                "note": f"{n_pages} trang, trích xuất trực tiếp (file có lớp văn bản).",
            }
        images, n_pages = _render_pdf_images(file_path)
        if n_pages > MAX_IMAGE_PAGES:
            note = (
                f"File scan {n_pages} trang — chỉ đọc {MAX_IMAGE_PAGES} trang đầu bằng AI-vision "
                f"(giới hạn để kiểm soát chi phí/kích thước). Có thể chạy lại cho phần còn lại."
            )
        else:
            note = f"File scan {n_pages} trang — đọc toàn bộ bằng AI-vision (không có lớp văn bản)."
        return {"mode": "images", "images": images, "note": note}

    raise ImportError_(f"Chưa hỗ trợ định dạng file '{ext}'. Chỉ hỗ trợ .pdf và .docx.")


def _build_field_spec(bieu):
    lines = []
    for f in bieu["fields"]:
        extra = ""
        if f.get("options"):
            extra = " (chọn đúng 1 trong: " + ", ".join(f["options"]) + ")"
        req = " [bắt buộc]" if f.get("required") else ""
        lines.append(f'- "{f["name"]}": {f["label"]}{req}{extra}')
    return "\n".join(lines)


def _build_prompt(bieu):
    field_spec = _build_field_spec(bieu)
    field_names = ", ".join(f'"{f["name"]}"' for f in bieu["fields"])
    return f"""Bạn là trợ lý trích xuất dữ liệu từ văn bản hành chính Việt Nam để nhập vào biểu "{bieu['label']}".

Danh sách trường dữ liệu cần điền cho MỖI dòng (đúng tên khóa JSON, giữ nguyên tiếng Việt có dấu trong giá trị):
{field_spec}

Yêu cầu:
- Đọc kỹ nội dung tài liệu được cung cấp, xác định các mục/dòng dữ liệu phù hợp với biểu trên.
- Trả về DUY NHẤT một mảng JSON (JSON array), mỗi phần tử là một object chỉ chứa các khóa: {field_names}.
- Chỉ điền giá trị khi có căn cứ rõ ràng trong tài liệu; để null nếu không chắc chắn hoặc tài liệu không đề cập — KHÔNG bịa thêm dữ liệu.
- Giữ nguyên văn phong, số hiệu văn bản, ngày tháng như trong tài liệu gốc.
- Không kèm giải thích, không markdown, không ```json — chỉ trả về JSON thuần túy bắt đầu bằng [ và kết thúc bằng ]."""


def _parse_json_array(raw_text):
    raw_text = raw_text.strip()
    match = re.search(r"\[.*\]", raw_text, re.DOTALL)
    if not match:
        raise ImportError_("AI không trả về dữ liệu JSON hợp lệ. Thử lại hoặc chọn file khác.")
    try:
        data = json.loads(match.group(0))
    except json.JSONDecodeError as e:
        raise ImportError_(f"Không đọc được JSON do AI trả về: {e}")
    if not isinstance(data, list):
        raise ImportError_("AI không trả về một danh sách (array) như yêu cầu.")
    return data


def extract_rows(bieu, source):
    """Goi Claude API de trich xuat du lieu theo schema cua bieu. Tra ve list[dict]."""
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise ImportError_(
            "Chưa cấu hình ANTHROPIC_API_KEY trên máy chạy app. "
            "Đặt biến môi trường ANTHROPIC_API_KEY (xem README.md) rồi khởi động lại app."
        )
    import anthropic

    client = anthropic.Anthropic(api_key=api_key)
    prompt = _build_prompt(bieu)

    if source["mode"] == "text":
        content = [{"type": "text", "text": prompt + "\n\n--- NỘI DUNG TÀI LIỆU ---\n" + source["text"]}]
    else:
        content = [{"type": "text", "text": prompt + "\n\nNội dung tài liệu là các trang ảnh scan dưới đây:"}]
        for img_bytes in source["images"]:
            content.append({
                "type": "image",
                "source": {
                    "type": "base64",
                    "media_type": "image/png",
                    "data": base64.b64encode(img_bytes).decode("ascii"),
                },
            })

    try:
        resp = client.messages.create(
            model=ANTHROPIC_MODEL,
            max_tokens=8000,
            messages=[{"role": "user", "content": content}],
        )
    except Exception as e:
        raise ImportError_(f"Gọi API Claude thất bại: {e}")

    raw_text = "".join(block.text for block in resp.content if getattr(block, "type", None) == "text")
    rows = _parse_json_array(raw_text)

    allowed_names = {f["name"] for f in bieu["fields"]}
    cleaned = []
    for row in rows:
        if not isinstance(row, dict):
            continue
        cleaned.append({k: v for k, v in row.items() if k in allowed_names})
    return cleaned
