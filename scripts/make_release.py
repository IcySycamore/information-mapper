"""打包免安装版（PyInstaller）。

用法::

    python scripts/make_release.py                # 构建目录版（推荐，启动快）
    python scripts/make_release.py --onefile      # 构建单文件版（便于发送，启动慢）
    python scripts/make_release.py --zip          # 另附便于分发的压缩包
    python scripts/make_release.py --keep-build   # 保留 build/ 与 dist/ 中间产物

产出目录::

    release/基层办公自动化助手/
        基层办公自动化助手.exe     主程序，双击即用
        使用说明.txt               首次使用提示与技术支持方式
        docs/USER_GUIDE.md         用户手册（软件「使用说明」页也会读取）
        config/*.example.json      示例配置
        README.md
        _internal/                 运行时（仅目录版；勿删、勿改名）

需要额外收集的资源，缺失时对应功能会在打包后失效：

- `rapidocr_onnxruntime` 的 OCR 模型（约 16 MB）与字典：`--collect-all`
- `onnxruntime` 的运行时与扩展（约 37 MB）：`--collect-binaries`
- `pypdfium2_raw` 的 `pdfium.dll`（约 7 MB）：`--collect-all`
- `docs/USER_GUIDE.md`：`--add-data` 打进包内，配合 `gui.manual_roots()` 的
  `sys._MEIPASS` 分支定位

构建结束后会用 ``--self-check`` 启动产物并校验依赖、OCR 模型与手册是否可用，
避免「打包成功但功能缺失」的情况。
"""

from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
import time
import zipfile
from importlib import util as importlib_util
from pathlib import Path
from typing import Sequence

ROOT = Path(__file__).resolve().parents[1]
APP_NAME = "基层办公自动化助手"
BUILD_NAME = "OfficeAssistant"  # 构建名用 ASCII，避开各工具链对非 ASCII 路径的兼容问题
RELEASE_DIR = ROOT / "release" / APP_NAME
MANUAL = ROOT / "docs" / "USER_GUIDE.md"

# 当前环境里用不到、但会被间接拉进来的大包；排除后体积明显下降
EXCLUDES = (
    "matplotlib",
    "IPython",
    "jupyter",
    "notebook",
    "pytest",
    "PyQt5",
    "PyQt6",
    "PySide2",
    "PySide6",
    "wx",
    "tkinter.test",
)

# 构建后从产物中删除的二进制：cv2 的视频编解码库（约 29 MB）在 OCR 场景不会加载，
# 删除只影响 cv2.VideoCapture 等视频接口。若后续要处理视频，清空该元组即可。
TRIM_BINARIES = ("cv2/opencv_videoio_ffmpeg500_64.dll",)


def bundle_args(onefile: bool) -> list[str]:
    """组装 PyInstaller 命令行参数。"""
    args = [
        "--noconfirm",
        "--clean",
        "--windowed",
        "--name",
        BUILD_NAME,
        "--paths",
        str(ROOT / "src"),
        "--distpath",
        str(ROOT / "dist"),
        "--workpath",
        str(ROOT / "build"),
        "--specpath",
        str(ROOT / "build"),
        "--collect-all",
        "rapidocr_onnxruntime",  # OCR 模型 + 字典
        "--collect-binaries",
        "onnxruntime",  # onnxruntime.dll 与 pybind11 扩展
        "--collect-all",
        "pypdfium2_raw",  # pdfium.dll，PDF 渲染用
        "--add-data",
        f"{MANUAL}{os.pathsep}docs",  # 包内手册副本
    ]
    if onefile:
        args.insert(3, "--onefile")
    for module in EXCLUDES:
        args += ["--exclude-module", module]
    args.append(str(ROOT / "scripts" / "gui_app.py"))  # 入口脚本
    return args


def run_build(onefile: bool) -> Path:
    """执行 PyInstaller；返回产物路径（目录版返回目录，单文件版返回 exe）。"""
    command = [sys.executable, "-m", "PyInstaller", *bundle_args(onefile)]
    print("开始打包：", " ".join(command), "\n")
    started = time.perf_counter()
    finished = subprocess.run(command, cwd=ROOT)
    if finished.returncode != 0:
        raise SystemExit(f"打包失败：PyInstaller 返回 {finished.returncode}")

    produced = ROOT / "dist" / (f"{BUILD_NAME}.exe" if onefile else BUILD_NAME)
    if not produced.exists():
        raise SystemExit(f"打包失败：未找到产物 {produced}")
    size = produced.stat().st_size if produced.is_file() else _dir_size(produced)
    print(f"\n打包完成，用时 {time.perf_counter() - started:.1f} 秒，体积 {size / 1048576:.1f} MB")
    return produced


def assemble(produced: Path, onefile: bool) -> Path:
    """把产物与随包分发的外部文件组装到发布目录。"""
    if RELEASE_DIR.exists():
        shutil.rmtree(RELEASE_DIR)
    (RELEASE_DIR / "docs").mkdir(parents=True)
    (RELEASE_DIR / "config").mkdir()

    if onefile:
        shutil.copy2(produced, RELEASE_DIR / f"{APP_NAME}.exe")
    else:
        shutil.copytree(produced, RELEASE_DIR, dirs_exist_ok=True)
        (RELEASE_DIR / f"{BUILD_NAME}.exe").replace(RELEASE_DIR / f"{APP_NAME}.exe")

    shutil.copy2(MANUAL, RELEASE_DIR / "docs" / MANUAL.name)
    shutil.copy2(ROOT / "README.md", RELEASE_DIR / "README.md")
    for item in sorted((ROOT / "config").glob("*.example.json")):
        shutil.copy2(item, RELEASE_DIR / "config" / item.name)
    (RELEASE_DIR / "使用说明.txt").write_text(usage_text(), encoding="utf-8")
    trim(RELEASE_DIR)
    return RELEASE_DIR


def trim(release_dir: Path) -> None:
    """删除产物中用不到的二进制（仅目录版有效；单文件版的资源在包内无法裁剪）。"""
    for relative in TRIM_BINARIES:
        for path in release_dir.rglob(Path(relative).name):
            size = path.stat().st_size
            path.unlink()
            print(f"裁剪 {path.relative_to(release_dir)}（{size / 1048576:.1f} MB）")


def usage_text() -> str:
    """生成随包分发的简短说明（联系方式取自 gui.SUPPORT_CONTACTS，避免两处维护）。"""
    lines = [
        f"{APP_NAME}（免安装版）",
        "",
        f"1. 双击「{APP_NAME}.exe」启动，无需安装 Python 或其他组件。",
        "2. 首次处理图片或扫描件会多等几秒，属 OCR 模型加载，正常现象。",
        "3. 操作提示见右侧「消息」栏与底部状态栏；操作步骤见「使用说明」页。",
        "4. 输出文件默认写入输入文件所在目录，重名时自动追加序号，不覆盖原文件。",
        "5. 目录版请整体复制使用，勿删除 _internal 目录；单文件版只拷 exe 即可。",
        "",
        "技术支持",
    ]
    lines += [f"  {label}：{value}" for label, value in support_contacts().items()]
    lines += [
        "",
        "反馈问题时请附上「使用说明」页「技术支持」按钮中的运行环境信息，",
        "以及右侧「消息」栏中对应的报错记录。",
    ]
    return "\n".join(lines) + "\n"


def support_contacts() -> dict[str, str]:
    """读取界面里定义的联系方式。"""
    if str(ROOT / "src") not in sys.path:
        sys.path.insert(0, str(ROOT / "src"))
    from information_mapper.gui import SUPPORT_CONTACTS

    return dict(SUPPORT_CONTACTS)


def version() -> str:
    """读取包版本号。"""
    if str(ROOT / "src") not in sys.path:
        sys.path.insert(0, str(ROOT / "src"))
    from information_mapper import __version__

    return __version__


def self_check(release_dir: Path) -> int:
    """启动产物做无界面自检，返回 0 表示通过。"""
    report = release_dir / "self-check.txt"
    if report.exists():
        report.unlink()
    print("\n校验产物：启动 exe 执行 --self-check …")
    subprocess.run([str(release_dir / f"{APP_NAME}.exe"), "--self-check"], cwd=release_dir, timeout=300)
    if not report.is_file():
        print("自检未通过：exe 未生成 self-check.txt")
        return 1
    content = report.read_text(encoding="utf-8")
    print(content)
    passed = "自检结果：通过" in content
    report.unlink()
    return 0 if passed else 1


def make_zip(tag: str) -> Path:
    """把发布目录压成一个 zip（上传 GitHub Release 用，勿提交到仓库）。"""
    target = ROOT / "release" / f"{APP_NAME}-{tag}.zip"
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as archive:
        for path in sorted(RELEASE_DIR.rglob("*")):
            if path.is_file():
                archive.write(path, Path(APP_NAME) / path.relative_to(RELEASE_DIR))
    return target


def _dir_size(path: Path) -> int:
    return sum(item.stat().st_size for item in path.rglob("*") if item.is_file())


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="打包免安装版")
    parser.add_argument("--onefile", action="store_true", help="打成单个 exe（启动需解包，较慢）")
    parser.add_argument("--zip", action="store_true", help="额外生成便于分发的 zip")
    parser.add_argument("--keep-build", action="store_true", help="保留 build/ 与 dist/ 中间产物")
    options = parser.parse_args(argv)

    if importlib_util.find_spec("PyInstaller") is None:
        print("缺少 PyInstaller，请先执行：python -m pip install pyinstaller")
        return 1
    if not MANUAL.is_file():
        print(f"缺少用户手册 {MANUAL}，打包后「使用说明」页将退化为内置简版")
        return 1

    produced = run_build(options.onefile)
    release_dir = assemble(produced, options.onefile)
    code = self_check(release_dir)
    if code != 0:
        print("产物自检未通过，请检查上面的报告；如需排查可加 --keep-build。")
        return code

    if options.zip:
        archive = make_zip(version())
        print(f"压缩包：{archive}（{archive.stat().st_size / 1048576:.1f} MB）")
    if not options.keep_build:
        shutil.rmtree(ROOT / "build", ignore_errors=True)
        shutil.rmtree(ROOT / "dist", ignore_errors=True)

    print(f"\n发布目录：{release_dir}（{_dir_size(release_dir) / 1048576:.1f} MB）")
    print("分发时请整个目录一起复制；单文件版只需拷 exe。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
