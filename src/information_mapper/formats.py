"""支持的输入/输出格式定义与判定工具。

统一维护"能读什么、能写什么、格式属于哪一类"，供 readers / writers /
converters / cli 共享，避免各处硬编码后缀名。
"""

from __future__ import annotations

from pathlib import Path

# ---------------------------------------------------------------- 输入格式
TABLE_INPUT_SUFFIXES: frozenset[str] = frozenset({".xlsx", ".xlsm", ".xls", ".csv", ".tsv"})
DOCUMENT_INPUT_SUFFIXES: frozenset[str] = frozenset({".docx", ".pdf", ".html", ".htm"})
TEXT_INPUT_SUFFIXES: frozenset[str] = frozenset({".txt", ".md", ".json", ".log"})
# 图片需要 OCR 识别文字（安装 ocr 可选依赖后可用）
IMAGE_INPUT_SUFFIXES: frozenset[str] = frozenset(
    {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp"}
)

INPUT_SUFFIXES: frozenset[str] = (
    TABLE_INPUT_SUFFIXES
    | DOCUMENT_INPUT_SUFFIXES
    | TEXT_INPUT_SUFFIXES
    | IMAGE_INPUT_SUFFIXES
)

# ---------------------------------------------------------------- 输出格式
TABLE_OUTPUT_SUFFIXES: frozenset[str] = frozenset({".xlsx", ".csv"})
DOCUMENT_OUTPUT_SUFFIXES: frozenset[str] = frozenset({".docx"})
TEXT_OUTPUT_SUFFIXES: frozenset[str] = frozenset({".txt", ".md", ".json", ".html"})

OUTPUT_SUFFIXES: frozenset[str] = (
    TABLE_OUTPUT_SUFFIXES | DOCUMENT_OUTPUT_SUFFIXES | TEXT_OUTPUT_SUFFIXES
)

# ------------------------------------------------------------ 人类可读说明
FORMAT_NOTES: dict[str, str] = {
    ".xlsx": "Excel 工作簿（可读可写，支持多工作表）",
    ".xlsm": "启用宏的 Excel 工作簿（读取）",
    ".xls": "旧版 Excel 工作簿（读取，需要 xlrd）",
    ".csv": "逗号分隔文本表格（可读可写，输出带 BOM 便于 Excel 打开）",
    ".tsv": "制表符分隔文本表格（读取）",
    ".docx": "Word 文档（可读可写，支持段落键值、表格、模板）",
    ".pdf": "PDF 文档（读取，需要 pypdf）",
    ".txt": "纯文本（可读可写）",
    ".md": "Markdown（可读可写，输出为表格/正文）",
    ".json": "JSON 记录数组（可读可写）",
    ".html": "HTML 表格/网页（可读可写）",
    ".png": "PNG 图片（OCR 识别文字）",
    ".jpg": "JPEG 图片（OCR 识别文字）",
    ".jpeg": "JPEG 图片（OCR 识别文字）",
    ".bmp": "BMP 图片（OCR 识别文字）",
    ".tif": "TIFF 图片（OCR 识别文字，适合多页扫描件）",
    ".tiff": "TIFF 图片（OCR 识别文字，适合多页扫描件）",
    ".webp": "WebP 图片（OCR 识别文字）",
}

# 归档 / 归类
KIND_TABLE = "table"
KIND_DOCUMENT = "document"
KIND_TEXT = "text"
KIND_UNKNOWN = "unknown"


def normalize_suffix(path: str | Path) -> str:
    """返回小写的文件后缀（含点），例如 ``.XLSX`` -> ``.xlsx``。"""
    return Path(path).suffix.lower()


def format_kind(suffix: str | Path) -> str:
    """判断格式类别：table / document / text / unknown。

    图片归入 document 类：需要经 OCR 提取文本后再解析。
    """
    normalized = normalize_suffix(suffix)
    if normalized in TABLE_INPUT_SUFFIXES:
        return KIND_TABLE
    if normalized in DOCUMENT_INPUT_SUFFIXES or normalized in IMAGE_INPUT_SUFFIXES:
        return KIND_DOCUMENT
    if normalized in TEXT_INPUT_SUFFIXES:
        return KIND_TEXT
    return KIND_UNKNOWN


def is_supported_input(path: str | Path) -> bool:
    return normalize_suffix(path) in INPUT_SUFFIXES


def is_supported_output(path: str | Path) -> bool:
    return normalize_suffix(path) in OUTPUT_SUFFIXES


def support_matrix() -> list[tuple[str, str, str, str]]:
    """返回 ``(后缀, 说明, 能否作为输入, 能否作为输出)`` 的列表，用于打印帮助。"""
    rows: list[tuple[str, str, str, str]] = []
    for suffix in sorted(FORMAT_NOTES, key=lambda item: (item not in OUTPUT_SUFFIXES, item)):
        rows.append(
            (
                suffix,
                FORMAT_NOTES[suffix],
                "是" if suffix in INPUT_SUFFIXES else "否",
                "是" if suffix in OUTPUT_SUFFIXES else "否",
            )
        )
    return rows


def input_suffixes_hint() -> str:
    return "、".join(sorted(INPUT_SUFFIXES))


def output_suffixes_hint() -> str:
    return "、".join(sorted(OUTPUT_SUFFIXES))
