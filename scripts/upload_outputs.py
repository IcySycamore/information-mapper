"""自动上传脚本：把生成好的结果文件自动上传到指定接口。

════════════════ 使用前请修改下面的配置（当前为占位符）════════════════
  UPLOAD_URL : 上传接口地址
  API_TOKEN  : 接口密钥 / API Key
═══════════════════════════════════════════════════════════════════════

用法::

    python scripts/upload_outputs.py                  # 预演（默认，不真正上传）
    python scripts/upload_outputs.py --apply          # 确认后真正上传
    python scripts/upload_outputs.py -i output --pattern *.xlsx --apply

推荐用环境变量注入密钥，避免把真实密钥写进文件::

    PowerShell:
        $env:INFO_MAP_UPLOAD_URL   = "https://真实地址/api/upload"
        $env:INFO_MAP_UPLOAD_TOKEN = "真实密钥"
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from information_mapper.readers import iter_input_files  # noqa: E402
from information_mapper.uploader import (  # noqa: E402
    UploadConfig,
    summarize,
    upload_file,
    upload_paths,
)

# ────────────────────────── 需要修改的配置（占位符） ──────────────────────────
UPLOAD_URL = "https://your-server.example.com/api/upload"  # TODO: 替换为真实上传地址
API_TOKEN = "YOUR_API_TOKEN_HERE"  # TODO: 替换为真实 API Key
UPLOAD_DIR = "output"  # 待上传文件所在目录
METHOD = "multipart"  # multipart（表单上传）或 json（Base64 JSON）
# ─────────────────────────────────────────────────────────────────────────────


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="自动上传结果文件（网址与密钥为占位符）")
    parser.add_argument("-i", "--input", default=UPLOAD_DIR, help=f"待上传的文件或目录（默认 {UPLOAD_DIR}）")
    parser.add_argument("--url", default=UPLOAD_URL, help="上传接口地址")
    parser.add_argument("--token", default=API_TOKEN, help="接口密钥 / API Key")
    parser.add_argument("--method", default=METHOD, choices=("multipart", "json"), help="提交方式")
    parser.add_argument("--field", default="file", help="multipart 表单字段名")
    parser.add_argument("--pattern", action="append", help="只上传匹配的文件，如 *.xlsx（可多次）")
    parser.add_argument("--retries", type=int, default=3, help="失败重试次数（默认 3）")
    parser.add_argument("--timeout", type=float, default=30.0, help="单次请求超时秒数（默认 30）")
    parser.add_argument("--apply", action="store_true", help="真正发送上传请求（默认只预演）")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    config = UploadConfig(
        url=args.url,
        token=args.token,
        method=args.method,
        field_name=args.field,
        retries=args.retries,
        timeout=args.timeout,
    )
    dry_run = not args.apply

    target = Path(args.input)
    print(f"接口地址：{config.url}")
    print(f"提交方式：{config.method}　密钥：{'已配置' if config.token else '未配置'}")
    print(f"模式：{'预演（不会真正上传）' if dry_run else '真实上传'}")

    if target.is_dir():
        files = iter_input_files(target, recursive=False, patterns=args.pattern)
        if not files:
            print(f"目录 {target} 中没有找到可上传的文件")
            return 1
        print(f"待上传 {len(files)} 个文件")
        results = upload_paths(files, config, dry_run=dry_run)
    elif target.is_file():
        results = [upload_file(target, config, dry_run=dry_run)]
    else:
        print(f"找不到待上传路径：{target}")
        return 1

    print(summarize(results))
    if dry_run:
        print("提示：确认无误后加 --apply 执行真实上传。")
        print("提示：密钥也可用环境变量 INFO_MAP_UPLOAD_URL / INFO_MAP_UPLOAD_TOKEN 提供。")
    return 0 if all(item.ok for item in results) else 1


if __name__ == "__main__":
    raise SystemExit(main())
