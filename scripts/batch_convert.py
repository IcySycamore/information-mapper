"""批量文档格式转换脚本：把目录下的文档统一转换成目标格式。

用法::

    python scripts/batch_convert.py                       # 用脚本内默认配置
    python scripts/batch_convert.py -i data -o output/csv --to csv
    python scripts/batch_convert.py -i data -o output/docx --to docx --pattern *.xlsx
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from information_mapper.converters import (  # noqa: E402
    convert_folder,
    normalize_target_format,
    summarize,
)

# ────────────────────────── 可按需修改的配置 ──────────────────────────
SOURCE_DIR = "data/inbox"  # 待转换文档所在目录
TARGET_DIR = "output/转换结果"  # 转换结果存放目录
TARGET_FORMAT = "csv"  # 目标格式：xlsx / csv / docx / txt / md / json / html
PATTERNS = None  # 例如 ["*.xlsx"]，None 表示转换全部受支持格式
OVERWRITE = False  # 目标同名文件是否直接覆盖（False 会自动加序号）
# ─────────────────────────────────────────────────────────────────────


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="批量把文档转换成统一格式")
    parser.add_argument("-i", "--input", default=SOURCE_DIR, help=f"输入目录（默认 {SOURCE_DIR}）")
    parser.add_argument("-o", "--output", default=TARGET_DIR, help=f"输出目录（默认 {TARGET_DIR}）")
    parser.add_argument("--to", default=TARGET_FORMAT, help=f"目标格式（默认 {TARGET_FORMAT}）")
    parser.add_argument("--pattern", action="append", help="文件通配符，如 *.xlsx（可多次）")
    parser.add_argument("--no-recursive", action="store_true", help="不递归子目录")
    parser.add_argument("--overwrite", action="store_true", help="同名文件直接覆盖")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    source = Path(args.input)
    if not source.is_dir():
        print(f"输入目录不存在：{source}")
        return 1

    suffix = normalize_target_format(args.to)
    results = convert_folder(
        source,
        args.output,
        suffix,
        recursive=not args.no_recursive,
        patterns=args.pattern or PATTERNS,
        overwrite=args.overwrite or OVERWRITE,
    )

    if not results:
        print(f"目录 {source} 中没有找到可转换的文件")
        return 1

    print(summarize(results))
    print(f"结果目录：{Path(args.output).resolve()}")
    return 0 if all(item.ok for item in results) else 1


if __name__ == "__main__":
    raise SystemExit(main())
