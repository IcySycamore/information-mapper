"""多格式读取测试。"""

from __future__ import annotations

from pathlib import Path

import pytest

from information_mapper.readers import (
    CONTENT_COLUMN,
    SOURCE_COLUMN,
    content_of,
    is_plain_content,
    iter_input_files,
    ocr_available,
    read_any,
    read_csv,
    read_excel,
    read_folder,
    read_image,
    read_json,
    read_text,
    read_word,
)


def test_read_excel_returns_records(xlsx_file: Path) -> None:
    records = read_excel(xlsx_file)
    assert len(records) == 2
    assert records[0]["姓名"] == "张三"


def test_read_csv_detects_gbk(gbk_csv_file: Path) -> None:
    records = read_csv(gbk_csv_file)
    assert records == [{"姓名": "赵六", "电话": "13600000000"}]


def test_read_word_parses_paragraphs_and_table(docx_file: Path) -> None:
    records = read_word(docx_file)
    assert {"姓名": "孙七", "联系电话": "13500000000"} in records
    assert any(record.get("姓名") == "王五" for record in records)


def test_read_text_parses_key_values(text_file: Path) -> None:
    assert read_text(text_file) == [{"姓名": "周八", "联系电话": "13400000000"}]


def test_read_text_without_structure_returns_content(plain_text_file: Path) -> None:
    records = read_text(plain_text_file)
    assert is_plain_content(records)
    assert "会议室" in content_of(records)


def test_read_text_keeps_structured_fields_and_body(tmp_path: Path) -> None:
    path = tmp_path / "混合.txt"
    path.write_text("姓名：张三\n正文第一段\n正文第二段\n", encoding="utf-8")
    records = read_text(path)
    assert records[0]["姓名"] == "张三"
    assert "正文第一段" in records[0][CONTENT_COLUMN]


def test_read_json_list(json_file: Path) -> None:
    assert read_json(json_file) == [{"姓名": "吴九", "联系电话": "13300000000"}]


def test_read_any_rejects_unsupported(tmp_path: Path) -> None:
    path = tmp_path / "说明.bin"
    path.write_bytes(b"\x00\x01")
    with pytest.raises(ValueError, match="不支持的输入格式"):
        read_any(path)


def test_iter_input_files_filters_and_sorts(inbox: Path) -> None:
    files = iter_input_files(inbox)
    names = [path.name for path in files]
    assert "报表A.xlsx" in names
    assert "报表B.csv" in names  # 子目录中的文件同样被收集
    assert files == sorted(files)  # 结果按路径稳定排序


def test_iter_input_files_respects_patterns(inbox: Path) -> None:
    files = iter_input_files(inbox, patterns=["*.xlsx"])
    assert [path.name for path in files] == ["报表A.xlsx"]


def test_iter_input_files_without_recursive(inbox: Path) -> None:
    names = [path.name for path in iter_input_files(inbox, recursive=False)]
    assert "报表B.csv" not in names


def test_read_folder_adds_source_column(inbox: Path) -> None:
    failed: list[tuple[Path, str]] = []
    records = read_folder(inbox, failed=failed)
    assert any(record[SOURCE_COLUMN] == "报表A.xlsx" for record in records)
    # 伪造的 .docx 无法解析，应被跳过而不是中断整批
    assert [path.name for path, _ in failed] == ["临时文件.docx"]


def test_read_folder_can_skip_source_column(inbox: Path) -> None:
    records = read_folder(inbox, source_column=None)
    assert all(SOURCE_COLUMN not in record for record in records)


# ---------------------------------------------------------------- 图片与 OCR
def test_image_suffixes_are_supported_inputs() -> None:
    from information_mapper.formats import INPUT_SUFFIXES, format_kind

    for suffix in (".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp"):
        assert suffix in INPUT_SUFFIXES
        assert format_kind(suffix) == "document"


def test_ocr_available_reports_boolean() -> None:
    assert isinstance(ocr_available(), bool)


def test_read_image_without_engine_raises(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    """未安装 OCR 组件时，应给出可操作的错误提示。"""
    from information_mapper import readers

    monkeypatch.setattr(readers, "_OCR_ENGINE", None)
    monkeypatch.setattr(readers.importlib.util, "find_spec", lambda name: None)

    def fake_engine() -> None:
        raise ValueError("识别图片或扫描件需要 OCR 组件")

    monkeypatch.setattr(readers, "_ocr_engine", fake_engine)

    path = tmp_path / "sample.png"
    path.write_bytes(b"\x89PNG")
    with pytest.raises(ValueError, match="OCR 组件"):
        read_image(path)


@pytest.mark.skipif(not ocr_available(), reason="未安装 OCR 组件")
def test_read_image_extracts_text(tmp_path: Path) -> None:
    """图片经 OCR 识别后，应能解析出记录。"""
    from PIL import Image, ImageDraw, ImageFont

    image = Image.new("RGB", (760, 200), "white")
    draw = ImageDraw.Draw(image)
    try:
        font = ImageFont.truetype("arial.ttf", 30)
    except OSError:  # pragma: no cover - 缺少字体时退回默认字体
        font = ImageFont.load_default()
    draw.text((30, 40), "Phone: 13800000000", fill="black", font=font)

    path = tmp_path / "scan.png"
    image.save(path)

    records = read_image(path)

    assert records
    assert any("13800000000" in str(value) for record in records for value in record.values())
