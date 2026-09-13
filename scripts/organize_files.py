"""文件自动归档脚本。

对应实践方案第 3 个场景：**文件资料管理 —— 多版本文件人工归档，流程繁琐**。

用法::

    python scripts/organize_files.py                      # 预演，只打印归档计划
    python scripts/organize_files.py --apply              # 真正移动文件
    python scripts/organize_files.py -i data/inbox --root output/归档 --apply
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from information_mapper.classify import (  # noqa: E402
    ClassifyConfig,
    organize,
    summarize,
)

# ────────────────────────── 可按需修改的配置 ──────────────────────────
SOURCE_DIR = "data/inbox"  # 待整理目录
CONFIG_FILE = "config/classify.example.json"  # 归档规则
ARCHIVE_ROOT = None  # 归档根目录；None 表示"待整理目录/归档"
APPLY = False  # True 表示真正移动文件，False 只预演
# ─────────────────────────────────────────────────────────────────────


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="按规则自动归档整理文件")
    parser.add_argument("-i", "--input", default=SOURCE_DIR, help=f"待整理目录（默认 {SOURCE_DIR}）")
    parser.add_argument("--config", default=CONFIG_FILE, help=f"归档规则 JSON（默认 {CONFIG_FILE}）")
    parser.add_argument("--root", help="归档根目录")
    parser.add_argument("--apply", action="store_true", help="真正执行移动（默认只预演）")
    parser.add_argument("--no-recursive", action="store_true", help="不递归子目录")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    source = Path(args.input)
    if not source.is_dir():
        print(f"待整理目录不存在：{source}")
        return 1

    config_path = Path(args.config)
    if not config_path.exists():
        print(f"归档规则文件不存在：{config_path}")
        return 1

    config = ClassifyConfig.from_file(config_path)
    dry_run = not (args.apply or APPLY)
    plans = organize(
        source,
        config,
        dry_run=dry_run,
        root=args.root or ARCHIVE_ROOT,
        recursive=not args.no_recursive,
    )

    if not plans:
        print(f"目录 {source} 中没有需要归档的文件")
        return 0

    print(summarize(plans, dry_run=dry_run))
    if dry_run:
        print("提示：当前为预演模式，确认计划无误后加 --apply 真正执行归档。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
