"""命令行入口。

子命令一览：

- ``map``      字段映射（把来源字段整理成目标表头）
- ``convert``  常见文档格式互相转换（单文件 / 整目录批量）
- ``merge``    大批量输入文档 → 一个输出文档
- ``report``   报表统计核算：多份报表汇总成一张总表
- ``scan``     批量解析目录文档，提取信息成汇总表
- ``classify`` 文件自动归档整理
- ``upload``   自动上传（网址与 API Key 为占位符）

``info-map --input ... --output ... --mapping ...`` 的旧式用法仍然可用。
"""

from __future__ import annotations

import argparse
import json
import sys
from dataclasses import asdict
from datetime import datetime
from pathlib import Path
from typing import Any, Sequence

from . import __version__
from .classify import ClassifyConfig
from .classify import organize as organize_files
from .classify import summarize as summarize_moves
from .converters import convert_file, convert_folder
from .converters import normalize_target_format, summarize as summarize_converts
from .core import map_records
from .formats import support_matrix
from .merger import collect_headers, merge_any, merge_tables, resolve_inputs
from .readers import iter_input_files, read_folder
from .uploader import (
    DEFAULT_API_TOKEN,
    DEFAULT_UPLOAD_URL,
    ENV_TOKEN,
    ENV_URL,
    SUPPORTED_METHODS,
    UploadConfig,
    UploadResult,
    upload_file,
    upload_paths,
)
from .uploader import summarize as summarize_uploads
from .writers import write_output

SUBCOMMANDS = ("map", "convert", "merge", "report", "scan", "classify", "upload", "formats")
SOURCE_COLUMN = "源文件"


# ---------------------------------------------------------------- 参数定义
def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="info-map",
        description="信息映射与办公文档批量处理工具（转换 / 合并 / 归档 / 上传）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "示例：\n"
            "  info-map map --input 报名表.xlsx --output 汇总.xlsx --mapping config/mapping.json\n"
            "  info-map convert --input data --output output/csv --to csv\n"
            "  info-map merge --input data/inbox --output output/总表.xlsx\n"
            "  info-map report --input data/报表 --output output/报表汇总.xlsx\n"
            "  info-map classify --input data/inbox --config config/classify.example.json --apply\n"
            "  info-map upload --input output --config config/upload.example.json\n"
        ),
    )
    parser.add_argument("--version", action="version", version=f"information-mapper {__version__}")
    subparsers = parser.add_subparsers(dest="command", metavar="子命令")

    # ---- map
    map_parser = subparsers.add_parser("map", help="字段映射：把来源字段整理成目标表头")
    _add_input_args(map_parser, multiple=False)
    _add_output_arg(map_parser, help_text="输出文件（.xlsx/.csv/.docx/.txt/.md/.json/.html）")
    map_parser.add_argument("--mapping", type=Path, required=True, help="字段映射 JSON 配置")
    map_parser.add_argument("--title", help="输出文档标题")
    map_parser.add_argument("--sheet", default="Sheet1", help="Excel 工作表名（默认 Sheet1）")
    map_parser.set_defaults(func=_cmd_map)

    # ---- convert
    convert_parser = subparsers.add_parser("convert", help="常见文档格式互相转换")
    _add_input_args(convert_parser, multiple=True, help_text="输入文件或目录（可多次传入）")
    _add_output_arg(convert_parser, help_text="输出文件（单文件）或输出目录（批量）")
    convert_parser.add_argument("--to", help="目标格式，如 csv / xlsx / docx / md（批量时必填）")
    convert_parser.add_argument("--mapping", type=Path, help="可选的字段映射 JSON 配置")
    convert_parser.add_argument("--title", help="输出文档标题")
    convert_parser.add_argument("--pattern", action="append", help="目录内文件通配符，如 *.xlsx（可多次）")
    convert_parser.add_argument("--no-recursive", action="store_true", help="不递归子目录")
    convert_parser.add_argument("--overwrite", action="store_true", help="目标同名时直接覆盖")
    convert_parser.set_defaults(func=_cmd_convert)

    # ---- merge
    merge_parser = subparsers.add_parser("merge", help="大批量输入文档 → 一个输出文档")
    _add_input_args(merge_parser, multiple=True, help_text="输入文件或目录（可多次传入）")
    _add_output_arg(merge_parser, help_text="唯一的输出文件")
    merge_parser.add_argument(
        "--mode",
        choices=("auto", "table", "word", "text"),
        default="auto",
        help="合并方式：table 汇总成一张表 / word 拼接 Word / text 拼接正文（默认 auto）",
    )
    merge_parser.add_argument("--headers", help="指定表头顺序，用英文逗号分隔")
    merge_parser.add_argument("--title", help="输出文档标题")
    merge_parser.add_argument("--sheet", default="汇总", help="Excel 工作表名（默认 汇总）")
    merge_parser.add_argument("--pattern", action="append", help="目录内文件通配符，如 *.xlsx（可多次）")
    merge_parser.add_argument("--no-recursive", action="store_true", help="不递归子目录")
    merge_parser.add_argument(
        "--no-source-column", action="store_true", help="不添加来源文件列（默认会加一列记录数据来源）"
    )
    merge_parser.set_defaults(func=_cmd_merge)

    # ---- report
    report_parser = subparsers.add_parser("report", help="报表统计核算：多份报表汇总成一张总表")
    _add_input_args(report_parser, multiple=True, help_text="报表文件或报表目录（可多次传入）")
    _add_output_arg(report_parser, required=False, help_text="输出文件（默认 output/报表汇总_日期.xlsx）")
    report_parser.add_argument("--headers", help="指定表头顺序，用英文逗号分隔")
    report_parser.add_argument("--title", help="输出文档标题")
    report_parser.add_argument("--sheet", default="汇总", help="Excel 工作表名（默认 汇总）")
    report_parser.add_argument("--pattern", action="append", help="目录内文件通配符（可多次）")
    report_parser.add_argument("--no-recursive", action="store_true", help="不递归子目录")
    report_parser.add_argument(
        "--no-source-column", action="store_true", help="不添加来源文件列（默认会加一列记录数据来源）"
    )
    report_parser.set_defaults(func=_cmd_report)

    # ---- scan
    scan_parser = subparsers.add_parser("scan", help="批量解析目录文档，提取信息成汇总表")
    _add_input_args(scan_parser, multiple=False, help_text="待扫描的目录")
    _add_output_arg(scan_parser, help_text="输出文件")
    scan_parser.add_argument("--mapping", type=Path, help="可选的字段映射 JSON 配置")
    scan_parser.add_argument("--title", help="输出文档标题")
    scan_parser.add_argument("--pattern", action="append", help="只扫描匹配的文件，如 *.docx（可多次）")
    scan_parser.add_argument("--no-recursive", action="store_true", help="不递归子目录")
    scan_parser.set_defaults(func=_cmd_scan)

    # ---- classify
    classify_parser = subparsers.add_parser("classify", help="文件自动归档整理")
    _add_input_args(classify_parser, multiple=False, help_text="待整理的目录")
    classify_parser.add_argument("--config", type=Path, required=True, help="归档规则 JSON 配置")
    classify_parser.add_argument("--root", type=Path, help="归档根目录（默认 输入目录/归档）")
    classify_parser.add_argument("--apply", action="store_true", help="真正执行移动（默认只预演）")
    classify_parser.add_argument("--no-recursive", action="store_true", help="不递归子目录")
    classify_parser.set_defaults(func=_cmd_classify)

    # ---- upload
    upload_parser = subparsers.add_parser("upload", help="自动上传结果文件（网址/密钥用占位符）")
    _add_input_args(upload_parser, multiple=True, help_text="待上传文件或目录（可多次传入）")
    upload_parser.add_argument("--config", type=Path, help="上传配置 JSON（默认内置占位符）")
    upload_parser.add_argument("--url", help=f"上传接口地址（占位符默认 {DEFAULT_UPLOAD_URL}）")
    upload_parser.add_argument("--token", help=f"API Key（占位符默认 {DEFAULT_API_TOKEN}）")
    upload_parser.add_argument("--method", choices=SUPPORTED_METHODS, help="提交方式：multipart / json")
    upload_parser.add_argument("--field", help="multipart 表单字段名（默认 file）")
    upload_parser.add_argument("--retries", type=int, help="失败重试次数（默认 3）")
    upload_parser.add_argument("--timeout", type=float, help="单次请求超时秒数（默认 30）")
    upload_parser.add_argument("--pattern", action="append", help="目录内文件通配符（可多次）")
    upload_parser.add_argument("--apply", action="store_true", help="真正发送请求（默认只预演）")
    upload_parser.set_defaults(func=_cmd_upload)

    # ---- formats
    formats_parser = subparsers.add_parser("formats", help="列出支持的输入/输出格式")
    formats_parser.set_defaults(func=_cmd_formats)

    return parser


def _add_input_args(parser: argparse.ArgumentParser, *, multiple: bool, help_text: str | None = None) -> None:
    if multiple:
        parser.add_argument(
            "-i",
            "--input",
            action="append",
            required=True,
            help=help_text or "输入文件或目录（可多次传入）",
        )
    else:
        parser.add_argument("-i", "--input", required=True, help=help_text or "输入文件")


def _add_output_arg(parser: argparse.ArgumentParser, *, help_text: str, required: bool = True) -> None:
    parser.add_argument("-o", "--output", required=required, help=help_text)


# ---------------------------------------------------------------- 入口
def main(argv: Sequence[str] | None = None) -> int:
    arguments = list(sys.argv[1:] if argv is None else argv)
    parser = build_parser()

    if not arguments:
        parser.print_help()
        return 0

    if arguments[0] not in SUBCOMMANDS:
        if arguments[0] in {"-h", "--help", "--version"}:
            parser.parse_args(arguments)  # 打印主帮助 / 版本后退出
            return 0
        if arguments[0].startswith("-"):
            return _run_legacy(parser, arguments)
        parser.error(f"未知子命令：{arguments[0]}，可选：{'、'.join(SUBCOMMANDS)}")

    args = parser.parse_args(arguments)
    return int(args.func(args) or 0)


def _run_legacy(parser: argparse.ArgumentParser, arguments: Sequence[str]) -> int:
    """兼容旧式 ``--input/--output/--mapping`` 用法。"""
    legacy = argparse.ArgumentParser(
        prog="info-map",
        description="旧式用法（等价于 info-map map）。推荐使用子命令，完整说明见 info-map --help",
    )
    legacy.add_argument("--input", required=True, type=Path, help="输入的 .xlsx/.xls/.docx 文件")
    legacy.add_argument("--output", required=True, type=Path, help="输出的目标文件")
    legacy.add_argument("--mapping", required=True, type=Path, help="字段映射 JSON 配置")
    legacy.add_argument("--title", help="输出文档标题")
    legacy.add_argument("--sheet", default="Sheet1", help="Excel 工作表名")
    args = legacy.parse_args(arguments)
    return _cmd_map(args)


# ---------------------------------------------------------------- 子命令实现
def _cmd_map(args: argparse.Namespace) -> int:
    mapping = _load_mapping(args.mapping)
    target_headers = list(mapping.get("target_headers") or [])
    if not target_headers:
        raise SystemExit("映射配置缺少 target_headers")

    source = _single_input_file(args.input)
    records = _read_single(source)
    mapped = map_records(records, target_headers, mapping.get("field_mapping") or {})
    output = write_output(
        mapped,
        target_headers,
        Path(args.output),
        title=getattr(args, "title", None),
        sheet_name=getattr(args, "sheet", "Sheet1"),
    )
    _echo(f"处理完成：读取 {len(records)} 条，输出 {len(mapped)} 条")
    _echo(f"文件位置：{output.resolve()}")
    return 0


def _cmd_convert(args: argparse.Namespace) -> int:
    inputs = [Path(item) for item in args.input]
    mapping = _load_mapping(args.mapping)
    patterns = args.pattern
    recursive = not args.no_recursive
    target_format = normalize_target_format(args.to) if args.to else None

    single_file = len(inputs) == 1 and inputs[0].is_file()
    output_path = Path(args.output)

    if single_file and (target_format is None or output_path.suffix.lower() == target_format):
        result = convert_file(inputs[0], output_path, mapping=mapping, title=args.title)
        _echo(result.summary())
        return 0 if result.ok else 1

    if target_format is None:
        raise SystemExit("批量转换必须用 --to 指定目标格式，例如：--to csv")

    results = []
    for item in inputs:
        if item.is_dir():
            results.extend(
                convert_folder(
                    item,
                    output_path,
                    target_format,
                    mapping=mapping,
                    recursive=recursive,
                    patterns=patterns,
                    overwrite=args.overwrite,
                )
            )
        elif item.is_file():
            destination = output_path if output_path.suffix.lower() == target_format else output_path / (
                item.stem + target_format
            )
            results.append(convert_file(item, destination, mapping=mapping))
        else:
            raise SystemExit(f"找不到输入：{item}")

    _echo(summarize_converts(results))
    return 0 if all(item.ok for item in results) else 1


def _cmd_merge(args: argparse.Namespace) -> int:
    inputs = [Path(item) for item in args.input]
    headers = _split_headers(args.headers)
    result = merge_any(
        inputs,
        Path(args.output),
        mode=args.mode,
        headers=headers,
        source_column=None if args.no_source_column else SOURCE_COLUMN,
        title=args.title,
        sheet_name=args.sheet,
        patterns=args.pattern,
        recursive=not args.no_recursive,
    )
    _echo(result.summary())
    _echo(f"文件位置：{result.output.resolve()}")
    return 0


def _cmd_report(args: argparse.Namespace) -> int:
    inputs = [Path(item) for item in args.input]
    files = resolve_inputs(inputs, patterns=args.pattern, recursive=not args.no_recursive)
    if not files:
        raise SystemExit("没有找到待汇总的报表文件")

    output = Path(args.output) if args.output else Path(f"output/报表汇总_{datetime.now():%Y%m%d_%H%M}.xlsx")
    result = merge_tables(
        files,
        output,
        headers=_split_headers(args.headers),
        source_column=None if args.no_source_column else SOURCE_COLUMN,
        title=args.title or "报表汇总",
        sheet_name=args.sheet,
    )
    _echo(result.summary())
    _echo(f"文件位置：{result.output.resolve()}")
    return 0


def _cmd_scan(args: argparse.Namespace) -> int:
    directory = Path(args.input)
    failed: list[tuple[Path, str]] = []
    records = read_folder(
        directory,
        recursive=not args.no_recursive,
        patterns=args.pattern,
        failed=failed,
    )
    files = iter_input_files(directory, recursive=not args.no_recursive, patterns=args.pattern)

    mapping = _load_mapping(args.mapping)
    if mapping and mapping.get("target_headers"):
        target_headers = list(mapping["target_headers"])
        records = map_records(records, target_headers, mapping.get("field_mapping") or {})
    else:
        target_headers = collect_headers(records)

    if not records:
        raise SystemExit(f"在 {directory} 中没有解析到任何记录")

    output = write_output(records, target_headers, Path(args.output), title=args.title or "信息汇总")
    _echo(f"扫描完成：{len(files)} 份文档，共 {len(records)} 条记录")
    for path, reason in failed:
        _echo(f"  已跳过 {path.name}：{reason}")
    _echo(f"文件位置：{output.resolve()}")
    return 0


def _cmd_classify(args: argparse.Namespace) -> int:
    config = ClassifyConfig.from_file(args.config)
    dry_run = not args.apply
    plans = organize_files(
        args.input,
        config,
        dry_run=dry_run,
        root=args.root,
        recursive=not args.no_recursive,
    )
    _echo(summarize_moves(plans, dry_run=dry_run))
    if dry_run:
        _echo("当前为预演模式，确认无误后加 --apply 真正执行归档。")
    return 0


def _cmd_upload(args: argparse.Namespace) -> int:
    config = _build_upload_config(args)
    dry_run = not args.apply
    targets = [Path(item) for item in args.input]

    _echo(f"接口地址：{config.url}")
    _echo(f"提交方式：{config.method}")
    pending = config.placeholders()
    if pending:
        _echo(
            f"提示：{'、'.join(pending)} 仍是占位符，真实上传前请替换"
            f"（命令行 --url/--token，或环境变量 {ENV_URL} / {ENV_TOKEN}）。"
        )

    results: list[UploadResult] = []
    for item in targets:
        if item.is_dir():
            files = iter_input_files(item, recursive=False, patterns=args.pattern)
            results.extend(upload_paths(files, config, dry_run=dry_run))
        elif item.is_file():
            results.append(upload_file(item, config, dry_run=dry_run))
        else:
            raise SystemExit(f"找不到待上传路径：{item}")

    if not results:
        raise SystemExit("没有找到可上传的文件")

    _echo(summarize_uploads(results))
    if dry_run:
        _echo("当前为预演模式（不会真正上传）。确认地址与密钥后加 --apply 执行真实上传。")
        _echo(f"可用环境变量覆盖配置：{ENV_URL} / {ENV_TOKEN}")
    return 0 if all(item.ok for item in results) else 1


def _cmd_formats(_: argparse.Namespace) -> int:
    _echo("格式支持一览（读＝可作为输入，写＝可作为输出）：")
    for suffix, note, can_input, can_output in support_matrix():
        marks = []
        if can_input == "是":
            marks.append("读")
        if can_output == "是":
            marks.append("写")
        _echo(f"  {suffix:<7}{'/'.join(marks):<6}{note}")
    return 0


# ---------------------------------------------------------------- 辅助函数
def _build_upload_config(args: argparse.Namespace) -> UploadConfig:
    config = UploadConfig.from_file(args.config) if args.config else UploadConfig()
    payload: dict[str, Any] = asdict(config)
    for key in ("url", "token", "method", "field", "retries", "timeout"):
        value = getattr(args, key, None)
        if value is not None:
            payload["field_name" if key == "field" else key] = value
    return UploadConfig.from_dict(payload)


def _load_mapping(path: str | Path | None) -> dict[str, Any]:
    if path is None:
        return {}
    config_path = Path(path)
    if not config_path.exists():
        raise SystemExit(f"找不到映射配置：{config_path}")
    return json.loads(config_path.read_text(encoding="utf-8"))


def _split_headers(value: str | None) -> list[str] | None:
    if not value:
        return None
    return [item.strip() for item in value.split(",") if item.strip()]


def _single_input_file(value: str | Path) -> Path:
    path = Path(value)
    if path.is_dir():
        raise SystemExit(f"该命令需要具体文件，传入的是目录：{path}")
    if not path.exists():
        raise SystemExit(f"找不到输入文件：{path}")
    return path


def _read_single(path: Path) -> list[dict[str, Any]]:
    from .readers import read_any

    return read_any(path)


def _echo(message: str) -> None:
    print(message)


if __name__ == "__main__":
    raise SystemExit(main())
