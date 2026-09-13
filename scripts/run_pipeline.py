"""一键自动化流程：报表汇总 → 材料合并 → 批量转换 → 自动归档 → 自动上传。

流程配置见 ``config/pipeline.example.json``，其中上传地址与密钥都是占位符。

用法::

    python scripts/run_pipeline.py                              # 按配置跑一遍（上传为预演）
    python scripts/run_pipeline.py --config config/pipeline.example.json
    python scripts/run_pipeline.py --upload-apply               # 上传阶段真正发送请求
    python scripts/run_pipeline.py --only report,merge          # 只跑指定阶段
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from information_mapper.classify import ClassifyConfig  # noqa: E402
from information_mapper.classify import organize as organize_files  # noqa: E402
from information_mapper.classify import summarize as summarize_moves  # noqa: E402
from information_mapper.converters import convert_folder  # noqa: E402
from information_mapper.converters import summarize as summarize_converts  # noqa: E402
from information_mapper.merger import merge_any, merge_tables, resolve_inputs  # noqa: E402
from information_mapper.uploader import UploadConfig, summarize, upload_folder  # noqa: E402

DEFAULT_CONFIG = "config/pipeline.example.json"
ALL_STEPS = ("report", "merge", "convert", "classify", "upload")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="一键跑完整自动化流程")
    parser.add_argument("--config", default=DEFAULT_CONFIG, help=f"流程配置（默认 {DEFAULT_CONFIG}）")
    parser.add_argument("--inbox", help="覆盖配置中的输入目录")
    parser.add_argument("--only", help=f"只跑指定阶段，逗号分隔，可选：{','.join(ALL_STEPS)}")
    parser.add_argument("--upload-apply", action="store_true", help="上传阶段真正发送请求")
    return parser.parse_args()


def load_config(path: str | Path) -> dict:
    config_path = Path(path)
    if not config_path.exists():
        raise SystemExit(f"找不到流程配置：{config_path}")
    return json.loads(config_path.read_text(encoding="utf-8"))


def main() -> int:
    args = parse_args()
    config = load_config(args.config)
    inbox = Path(args.inbox or config.get("inbox") or "data/inbox")
    steps = [item.strip() for item in (args.only or ",".join(ALL_STEPS)).split(",") if item.strip()]

    print("=" * 60)
    print(f"工作目录：{inbox.resolve()}")
    print(f"执行阶段：{'、'.join(steps)}")
    print("=" * 60)

    if not inbox.is_dir() and any(step in steps for step in ("report", "merge", "convert", "classify")):
        print(f"输入目录不存在：{inbox}，请先把待处理文档放进去。")
        return 1

    if "report" in steps:
        _step_report(config, inbox)
    if "merge" in steps:
        _step_merge(config, inbox)
    if "convert" in steps:
        _step_convert(config, inbox)
    if "classify" in steps:
        _step_classify(config, inbox)
    if "upload" in steps:
        _step_upload(config, args.upload_apply)

    print("\n流程结束。")
    return 0


def _step_report(config: dict, inbox: Path) -> None:
    section = dict(config.get("report") or {})
    if not section.get("enabled", True):
        print("\n[1/5] 报表汇总：已跳过（配置中关闭）")
        return
    print("\n[1/5] 报表汇总")
    patterns = section.get("patterns") or ["*.xlsx", "*.csv"]
    files = resolve_inputs([inbox], patterns=patterns, recursive=True)
    if not files:
        print(f"  没有找到匹配的报表（{'、'.join(patterns)}），跳过")
        return
    output = section.get("output") or "output/报表汇总.xlsx"
    result = merge_tables(files, output, title="报表汇总")
    print(f"  {result.record_count} 条记录来自 {len(result.sources)} 份报表")
    print(f"  输出：{result.output.resolve()}")


def _step_merge(config: dict, inbox: Path) -> None:
    section = dict(config.get("merge") or {})
    if not section.get("enabled", True):
        print("\n[2/5] 材料合并：已跳过（配置中关闭）")
        return
    print("\n[2/5] 材料合并")
    patterns = section.get("patterns") or ["*.docx", "*.pdf", "*.txt", "*.md"]
    files = resolve_inputs([inbox], patterns=patterns, recursive=True)
    if not files:
        print("  没有找到可合并的文档，跳过")
        return
    output = section.get("output") or "output/材料合集.docx"
    result = merge_any(files, output, mode=section.get("mode", "auto"))
    print(f"  已合并 {len(result.sources)} 份文档")
    print(f"  输出：{result.output.resolve()}")


def _step_convert(config: dict, inbox: Path) -> None:
    section = dict(config.get("convert") or {})
    if not section.get("enabled", True):
        print("\n[3/5] 批量转换：已跳过（配置中关闭）")
        return
    print("\n[3/5] 批量转换")
    target = section.get("to") or "csv"
    output_dir = section.get("output_dir") or "output/转换结果"
    results = convert_folder(
        inbox,
        output_dir,
        target,
        patterns=section.get("patterns"),
    )
    if not results:
        print("  没有找到可转换的文档，跳过")
        return
    print("  " + summarize_converts(results).replace("\n", "\n  "))
    print(f"  输出目录：{Path(output_dir).resolve()}")


def _step_classify(config: dict, inbox: Path) -> None:
    section = dict(config.get("classify") or {})
    if not section.get("enabled", False):
        print("\n[4/5] 自动归档：已跳过（配置中关闭）")
        return
    print("\n[4/5] 自动归档")
    config_path = Path(section.get("config") or "config/classify.example.json")
    if not config_path.exists():
        print(f"  归档规则不存在：{config_path}，跳过")
        return
    rules = ClassifyConfig.from_file(config_path)
    dry_run = bool(section.get("dry_run", True))
    plans = organize_files(inbox, rules, dry_run=dry_run)
    print("  " + summarize_moves(plans, dry_run=dry_run).replace("\n", "\n  "))


def _step_upload(config: dict, apply: bool) -> None:
    section = dict(config.get("upload") or {})
    if not section.get("enabled", False):
        print("\n[5/5] 自动上传：已跳过（配置中关闭，需在 pipeline 配置里启用）")
        return
    print("\n[5/5] 自动上传")
    config_path = section.get("config")
    upload_config = (
        UploadConfig.from_file(config_path) if config_path and Path(config_path).exists() else UploadConfig()
    )
    dry_run = not apply or bool(section.get("dry_run", True))
    directory = Path(config.get("output_dir") or "output")
    results = upload_folder(directory, upload_config, patterns=section.get("patterns"), dry_run=dry_run)
    if not results:
        print("  没有找到可上传的文件，跳过")
        return
    print("  " + summarize(results).replace("\n", "\n  "))
    if dry_run:
        print("  提示：当前为上传预演，确认配置后加 --upload-apply 执行真实上传。")


if __name__ == "__main__":
    raise SystemExit(main())
