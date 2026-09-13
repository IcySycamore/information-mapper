"""报表统计核算脚本：把目录下所有报表自动汇总成一张总表。

对应实践方案第 1 个场景：**报表统计核算 —— 手工汇总 Excel 报表，耗时高、易出错**。

用法::

    python scripts/report_merge.py                       # 用脚本内默认配置
    python scripts/report_merge.py -i data/报表 -o output/汇总.xlsx
    python scripts/report_merge.py --headers 姓名,身份证号,联系电话,家庭住址
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from information_mapper.merger import merge_tables, resolve_inputs  # noqa: E402

# ────────────────────────── 可按需修改的配置 ──────────────────────────
INBOX_DIR = "data/inbox"  # 报表所在目录
OUTPUT_FILE = "output/报表汇总.xlsx"  # 汇总结果路径
PATTERNS = ["*.xlsx", "*.xls", "*.csv", "*.tsv"]  # 参与汇总的文件类型
SOURCE_COLUMN = "源文件"  # 记录每条数据来自哪份报表；设为 None 则不添加
SHEET_NAME = "汇总"
# ─────────────────────────────────────────────────────────────────────


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="把多份报表汇总成一张总表")
    parser.add_argument("-i", "--input", nargs="+", default=[INBOX_DIR], help="报表文件或目录")
    parser.add_argument("-o", "--output", default=OUTPUT_FILE, help=f"输出文件（默认 {OUTPUT_FILE}）")
    parser.add_argument("--pattern", action="append", help="文件通配符，如 *.xlsx（可多次）")
    parser.add_argument("--headers", help="指定表头顺序，用英文逗号分隔")
    parser.add_argument("--sheet", default=SHEET_NAME, help="Excel 工作表名")
    parser.add_argument("--no-recursive", action="store_true", help="不递归子目录")
    parser.add_argument("--no-source-column", action="store_true", help="不添加来源文件列")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    patterns = args.pattern or PATTERNS
    inputs = resolve_inputs(args.input, patterns=patterns, recursive=not args.no_recursive)

    if not inputs:
        print(f"未找到报表文件，请检查目录与文件类型（当前匹配：{'、'.join(patterns)}）")
        return 1

    headers = [item.strip() for item in args.headers.split(",")] if args.headers else None
    result = merge_tables(
        inputs,
        args.output,
        headers=headers,
        source_column=None if args.no_source_column else SOURCE_COLUMN,
        title="报表汇总",
        sheet_name=args.sheet,
    )

    print(result.summary())
    print(f"文件位置：{result.output.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
