"""常见文档格式转换。

任意受支持输入 → 任意受支持输出：

- 表格类输入（xlsx/csv/docx 表格）→ 保留为记录，可写 xlsx/csv/docx/txt/md/json/html
- 正文类输入（txt/md/pdf）→ 原文直转 txt/md/docx/html
- 提供字段映射配置时，先做字段映射再输出
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable, Sequence

from .core import map_records
from .formats import OUTPUT_SUFFIXES, normalize_suffix
from .merger import collect_headers
from .readers import (
    content_of,
    is_plain_content,
    iter_input_files,
    read_any,
)
from .writers import write_output, write_text_document

# 可以"原文直转"的目标格式
_TEXT_TARGET_SUFFIXES = frozenset({".txt", ".md", ".docx", ".html", ".htm"})


@dataclass
class ConvertResult:
    """单个文件的转换结果。"""

    source: Path
    target: Path
    ok: bool = True
    error: str = ""

    def summary(self) -> str:
        status = "成功" if self.ok else "失败"
        detail = f"（{self.error}）" if self.error else ""
        return f"[{status}] {self.source.name} → {self.target.name}{detail}"


def normalize_target_format(value: str) -> str:
    """把 ``csv`` / ``.csv`` / ``CSV`` 统一成 ``.csv``。"""
    suffix = value.strip().lower()
    if not suffix:
        raise ValueError("目标格式不能为空")
    if not suffix.startswith("."):
        suffix = f".{suffix}"
    if suffix not in OUTPUT_SUFFIXES:
        raise ValueError(
            f"不支持的目标格式：{suffix}，请使用 {'、'.join(sorted(OUTPUT_SUFFIXES))}"
        )
    return suffix


def convert_file(
    source: str | Path,
    target: str | Path,
    *,
    mapping: dict[str, Any] | None = None,
    title: str | None = None,
) -> ConvertResult:
    """转换单个文档，返回转换结果。"""
    src = Path(source)
    dst = Path(target)
    suffix = normalize_suffix(dst)

    if suffix not in OUTPUT_SUFFIXES:
        raise ValueError(
            f"不支持的输出格式：{suffix or dst.name}，"
            f"请使用 {'、'.join(sorted(OUTPUT_SUFFIXES))}"
        )
    if not src.exists():
        raise ValueError(f"找不到输入文件：{src}")

    try:
        records = read_any(src)
        target_headers = list((mapping or {}).get("target_headers") or [])
        if target_headers:
            records = map_records(records, target_headers, (mapping or {}).get("field_mapping") or {})
            write_output(records, target_headers, dst, title=title)
        elif is_plain_content(records) and suffix in _TEXT_TARGET_SUFFIXES:
            write_text_document(content_of(records), dst, title=title or src.stem)
        else:
            write_output(records, collect_headers(records), dst, title=title)
    except Exception as error:  # noqa: BLE001 - 批量时需要记录失败项
        return ConvertResult(source=src, target=dst, ok=False, error=str(error))
    return ConvertResult(source=src, target=dst, ok=True)


def convert_folder(
    source_dir: str | Path,
    target_dir: str | Path,
    target_format: str,
    *,
    mapping: dict[str, Any] | None = None,
    recursive: bool = True,
    patterns: Iterable[str] | None = None,
    overwrite: bool = False,
    keep_structure: bool = True,
) -> list[ConvertResult]:
    """批量转换目录下的所有文档到同一目标格式。"""
    suffix = normalize_target_format(target_format)
    root = Path(source_dir)
    destination_root = Path(target_dir)
    files = iter_input_files(root, recursive=recursive, patterns=patterns)

    results: list[ConvertResult] = []
    for path in files:
        if keep_structure and recursive:
            relative = path.relative_to(root).with_suffix(suffix)
        else:
            relative = Path(path.name).with_suffix(suffix)
        destination = _unique_path(destination_root / relative, overwrite=overwrite)
        results.append(convert_file(path, destination, mapping=mapping))
    return results


def _unique_path(path: Path, *, overwrite: bool) -> Path:
    """目标已存在时自动加序号，避免覆盖已有成果。"""
    if overwrite or not path.exists():
        return path
    index = 1
    while True:
        candidate = path.with_name(f"{path.stem}({index}){path.suffix}")
        if not candidate.exists():
            return candidate
        index += 1


def summarize(results: Sequence[ConvertResult]) -> str:
    """把批量转换结果整理成可读摘要。"""
    ok = sum(1 for item in results if item.ok)
    lines = [f"转换完成：成功 {ok} 个，失败 {len(results) - ok} 个"]
    lines.extend(item.summary() for item in results if not item.ok)
    return "\n".join(lines)
