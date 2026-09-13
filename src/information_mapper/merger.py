"""把大批量输入文档合并为一个输出文档。

覆盖三类合并场景：

1. ``merge_tables``：多个 Excel / CSV / Word 表格 / PDF → 一张汇总表（单文件）
2. ``merge_word_documents``：多个 Word → 一个 Word（保留段落与表格）
3. ``merge_text_documents``：多个 txt / md / pdf / docx → 一篇长文档

``merge_any`` 会按输入与输出的类型自动挑选合适的合并方式。
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable, Sequence

from docx import Document

from .formats import normalize_suffix
from .readers import (
    SOURCE_COLUMN,
    content_of,
    iter_input_files,
    read_any,
    read_text_content,
)
from .writers import write_output, write_text_document

TABLE_SUFFIXES = frozenset({".xlsx", ".xlsm", ".xls", ".csv", ".tsv"})
PLAIN_SUFFIXES = frozenset({".txt", ".md", ".html", ".htm", ".json"})
_ALLOWED_MODES = frozenset({"auto", "table", "word", "text"})


@dataclass
class MergeResult:
    """合并结果，供 CLI / 脚本打印汇总信息。"""

    sources: list[Path]
    output: Path
    headers: list[str] = field(default_factory=list)
    record_count: int = 0
    failed: list[tuple[Path, str]] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return bool(self.sources)

    def summary(self) -> str:
        lines = [
            f"已合并 {len(self.sources)} 份文件 → {self.output.name}",
            f"记录数：{self.record_count}",
        ]
        if self.headers:
            lines.append(f"表头：{' | '.join(self.headers)}")
        if self.failed:
            lines.append(f"跳过 {len(self.failed)} 份：")
            lines.extend(f"  - {path.name}：{reason}" for path, reason in self.failed)
        return "\n".join(lines)


# ------------------------------------------------------------------ 公共工具
def collect_headers(
    records: Sequence[dict[str, Any]],
    preferred: Iterable[str] | None = None,
) -> list[str]:
    """汇总表头：``preferred`` 优先，其余按首次出现顺序补全。"""
    headers: list[str] = []
    seen: set[str] = set()
    for name in preferred or []:
        if name and name not in seen:
            headers.append(name)
            seen.add(name)
    for record in records:
        for name in record:
            if name not in seen:
                headers.append(name)
                seen.add(name)
    return headers


def resolve_inputs(
    inputs: Iterable[str | Path],
    *,
    patterns: Iterable[str] | None = None,
    recursive: bool = True,
) -> list[Path]:
    """把"文件或目录"的混合输入展开为待处理文件列表（去重并排序）。"""
    resolved: list[Path] = []
    seen: set[Path] = set()
    for item in inputs:
        path = Path(item)
        if path.is_dir():
            candidates = iter_input_files(path, recursive=recursive, patterns=patterns)
        elif path.is_file():
            candidates = [path]
        else:
            raise ValueError(f"找不到输入：{path}")
        for candidate in candidates:
            key = candidate.resolve()
            if key not in seen:
                seen.add(key)
                resolved.append(candidate)
    return resolved


def pick_mode(input_paths: Sequence[Path], output_path: str | Path, mode: str) -> str:
    """推断合并方式；显式传入 table/word/text 时原样返回。"""
    if mode not in _ALLOWED_MODES:
        raise ValueError(f"不支持的合并方式：{mode}，可选 auto/table/word/text")

    output_suffix = normalize_suffix(output_path)
    if mode == "auto":
        if output_suffix in {".xlsx", ".xlsm", ".csv"}:
            return "table"
        if all(normalize_suffix(path) == ".docx" for path in input_paths):
            return "word"
        if any(normalize_suffix(path) in TABLE_SUFFIXES for path in input_paths):
            return "table"
        return "text"
    if mode == "table" and output_suffix in {".docx", ".txt", ".md", ".html", ".htm", ".json"}:
        # Word/文本输出同样支持"记录 → 表格"，无需改写模式
        return "table"
    return mode


def merge_any(
    inputs: Iterable[str | Path],
    output_path: str | Path,
    *,
    mode: str = "auto",
    headers: Iterable[str] | None = None,
    source_column: str | None = SOURCE_COLUMN,
    title: str | None = None,
    sheet_name: str = "汇总",
    patterns: Iterable[str] | None = None,
    recursive: bool = True,
    add_source_heading: bool = True,
) -> MergeResult:
    """批量合并为一个输出文档，自动选择合并方式。"""
    input_paths = resolve_inputs(inputs, patterns=patterns, recursive=recursive)
    if not input_paths:
        raise ValueError("没有找到可合并的输入文档")

    resolved_mode = pick_mode(input_paths, output_path, mode)
    if resolved_mode == "word":
        return merge_word_documents(
            input_paths,
            output_path,
            title=title,
            add_source_heading=add_source_heading,
        )
    if resolved_mode == "text":
        return merge_text_documents(
            input_paths,
            output_path,
            title=title,
            add_source_heading=add_source_heading,
        )
    return merge_tables(
        input_paths,
        output_path,
        headers=headers,
        source_column=source_column,
        title=title,
        sheet_name=sheet_name,
    )


# ------------------------------------------------------------------ 三种合并
def merge_tables(
    input_paths: Sequence[str | Path],
    output_path: str | Path,
    *,
    headers: Iterable[str] | None = None,
    source_column: str | None = SOURCE_COLUMN,
    title: str | None = None,
    sheet_name: str = "汇总",
    append: bool = False,
) -> MergeResult:
    """把所有输入文档的信息汇总成一张表，写出为单个文件。"""
    sources: list[Path] = []
    failed: list[tuple[Path, str]] = []
    all_records: list[dict[str, Any]] = []

    for item in input_paths:
        path = Path(item)
        try:
            records = read_any(path)
        except Exception as error:  # noqa: BLE001 - 单份文件失败不应中断整批
            failed.append((path, str(error)))
            continue
        sources.append(path)
        for record in records:
            row = dict(record)
            if source_column:
                row = {source_column: path.name, **row}
            all_records.append(row)

    resolved_headers = collect_headers(all_records, headers)
    target = Path(output_path)
    write_output(
        all_records,
        resolved_headers,
        target,
        title=title or f"批量汇总（共 {len(all_records)} 条）",
        sheet_name=sheet_name,
        append=append,
    )
    return MergeResult(
        sources=sources,
        output=target,
        headers=resolved_headers,
        record_count=len(all_records),
        failed=failed,
    )


def merge_word_documents(
    input_paths: Sequence[str | Path],
    output_path: str | Path,
    *,
    title: str | None = None,
    add_source_heading: bool = True,
    page_break_between: bool = True,
) -> MergeResult:
    """把多个 Word 文档按顺序拼成一个 Word，保留段落与表格内容。"""
    target = Document()
    if title:
        target.add_heading(title, level=0)

    sources: list[Path] = []
    failed: list[tuple[Path, str]] = []
    for index, item in enumerate(input_paths):
        path = Path(item)
        try:
            source = Document(str(path))
        except Exception as error:  # noqa: BLE001
            failed.append((path, str(error)))
            continue
        if index and page_break_between:
            target.add_page_break()
        if add_source_heading:
            target.add_heading(path.stem, level=1)
        _copy_docx_body(source, target)
        sources.append(path)

    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    target.save(output)
    return MergeResult(sources=sources, output=output, record_count=len(sources), failed=failed)


def merge_text_documents(
    input_paths: Sequence[str | Path],
    output_path: str | Path,
    *,
    title: str | None = None,
    add_source_heading: bool = True,
    separator: str = "\n\n",
) -> MergeResult:
    """把多份文档的正文依次拼接成一篇长文档。"""
    sources: list[Path] = []
    failed: list[tuple[Path, str]] = []
    chunks: list[str] = []

    for item in input_paths:
        path = Path(item)
        try:
            text = extract_plain_text(path)
        except Exception as error:  # noqa: BLE001
            failed.append((path, str(error)))
            continue
        sources.append(path)
        if not text.strip():
            continue
        if add_source_heading:
            heading = "##" if normalize_suffix(output_path) == ".md" else ""
            chunks.append(f"{heading} {path.stem}".strip() + "\n\n" + text.strip())
        else:
            chunks.append(text.strip())

    body = separator.join(chunks)
    output = Path(output_path)
    write_text_document(body, output, title=title or "合并文档")
    return MergeResult(sources=sources, output=output, record_count=len(sources), failed=failed)


def extract_plain_text(path: str | Path) -> str:
    """把任意受支持文档转成纯文本，用于正文合并与预览。"""
    target = Path(path)
    suffix = normalize_suffix(target)
    if suffix == ".docx":
        document = Document(str(target))
        return "\n".join(
            paragraph.text for paragraph in document.paragraphs if paragraph.text.strip()
        )
    if suffix in PLAIN_SUFFIXES and suffix not in {".html", ".htm", ".json"}:
        return read_text_content(target)
    return content_of(read_any(target))


# ------------------------------------------------------------------ 内部工具
def _copy_docx_body(source: Document, target: Document) -> None:
    for paragraph in source.paragraphs:
        if not paragraph.text.strip():
            continue
        target.add_paragraph(paragraph.text, style=_safe_style(target, paragraph.style))
    for table in source.tables:
        rows = len(table.rows)
        columns = len(table.columns)
        if not rows or not columns:
            continue
        new_table = target.add_table(rows=rows, cols=columns)
        try:
            new_table.style = "Table Grid"
        except KeyError:  # pragma: no cover - 模板缺少该样式时忽略
            pass
        for row_index, row in enumerate(table.rows):
            for column_index, cell in enumerate(row.cells):
                if column_index < columns:
                    new_table.cell(row_index, column_index).text = cell.text


def _safe_style(target: Document, style: Any) -> str | None:
    try:
        name = style.name
    except AttributeError:  # pragma: no cover
        return None
    try:
        target.styles[name]
    except KeyError:
        return None
    return name
