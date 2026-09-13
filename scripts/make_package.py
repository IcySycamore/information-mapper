"""打 Python 库版（wheel + sdist + 源码包）。

与免安装版（`scripts/make_release.py`）并列，两个版本的产物分开存放：

======  ==================================  ==========================
版本    产物                                适用对象
======  ==================================  ==========================
免安装版 `release/`，exe + 可选 zip           没有 Python 环境的基层使用者
库版    `packages/`，wheel + sdist + 源码包   有 Python 环境，需命令行或二次开发
======  ==================================  ==========================

用法::

    python scripts/make_package.py             # 构建 wheel、sdist 与源码包
    python scripts/make_package.py --no-source # 只要 wheel 与 sdist
    python scripts/make_package.py --isolated  # 用隔离环境构建（需联网下载构建依赖）

产出::

    packages/
        information_mapper-<版本>-py3-none-any.whl   pip 安装用
        information_mapper-<版本>.tar.gz             sdist
        information-mapper-<版本>-source.zip          含 scripts/docs/config 与 .bat，解压即用
"""

from __future__ import annotations

import argparse
import subprocess
import sys
import zipfile
from importlib import util as importlib_util
from pathlib import Path
from typing import Sequence

ROOT = Path(__file__).resolve().parents[1]
PACKAGES_DIR = ROOT / "packages"

# 源码包内容：解压后目录结构与仓库一致，可直接双击「一键配置环境.bat」
SOURCE_ITEMS = (
    "src",
    "scripts",
    "docs",
    "config",
    "tests",
    "pyproject.toml",
    "README.md",
    "一键配置环境.bat",
    "启动界面.bat",
    ".gitignore",
)
SOURCE_EXCLUDES = ("__pycache__", ".pytest_cache", ".egg-info", "output", "materials", "release", "packages")


def version() -> str:
    """读取包版本号。"""
    if str(ROOT / "src") not in sys.path:
        sys.path.insert(0, str(ROOT / "src"))
    from information_mapper import __version__

    return __version__


def build_distributions(isolated: bool) -> list[Path]:
    """调用 build 生成 wheel 与 sdist。"""
    if importlib_util.find_spec("build") is None:
        raise SystemExit("缺少构建工具，请先执行：python -m pip install build")
    PACKAGES_DIR.mkdir(parents=True, exist_ok=True)
    command = [sys.executable, "-m", "build", "--outdir", str(PACKAGES_DIR)]
    if not isolated:
        command.append("--no-isolation")
    print("构建 wheel 与 sdist：", " ".join(command[1:]), "\n")
    finished = subprocess.run(command, cwd=ROOT)
    if finished.returncode != 0 and not isolated:
        print("\n本地构建失败，改用隔离环境重试（需联网）…")
        return build_distributions(isolated=True)
    if finished.returncode != 0:
        raise SystemExit(f"构建失败：build 返回 {finished.returncode}")
    return sorted(PACKAGES_DIR.glob("*.whl")) + sorted(PACKAGES_DIR.glob("*.tar.gz"))


def build_source_zip(tag: str) -> Path:
    """把源码、脚本、文档与批处理打包成可直接使用的源码包。"""
    target = PACKAGES_DIR / f"information-mapper-{tag}-source.zip"
    root_name = f"information-mapper-{tag}"
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as archive:
        for item in SOURCE_ITEMS:
            source = ROOT / item
            if not source.exists():
                continue
            if source.is_file():
                archive.write(source, f"{root_name}/{item}")
                continue
            for path in sorted(source.rglob("*")):
                if not path.is_file() or any(part in SOURCE_EXCLUDES or part.endswith(".egg-info") for part in path.parts):
                    continue
                archive.write(path, f"{root_name}/{path.relative_to(ROOT)}")
    return target


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="打 Python 库版")
    parser.add_argument("--no-source", action="store_true", help="不生成源码包")
    parser.add_argument("--isolated", action="store_true", help="用隔离环境构建 wheel 与 sdist")
    options = parser.parse_args(argv)

    tag = version()
    artifacts = build_distributions(options.isolated)
    if not artifacts:
        raise SystemExit("构建结束但没有产物，请检查上面的输出")
    for artifact in artifacts:
        print(f"已生成 {artifact.relative_to(ROOT)}（{artifact.stat().st_size / 1024:.0f} KB）")

    if not options.no_source:
        source = build_source_zip(tag)
        print(f"已生成 {source.relative_to(ROOT)}（{source.stat().st_size / 1048576:.1f} MB）")

    print(f"\n产出目录：{PACKAGES_DIR}")
    print("免安装 exe 版请用 scripts\\make_release.py，产物在 release/。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
