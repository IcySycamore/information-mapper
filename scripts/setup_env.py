"""一键环境配置（由「一键配置环境.bat」调用）。

流程：

1. 检查当前 Python 版本（需 3.10 或更高）
2. 创建或校验 ``.venv`` 虚拟环境（损坏时自动重建）
3. 安装项目依赖；默认源失败时自动改用国内镜像
4. 安装扩展组件：xlrd（旧版 .xls 读取）、rapidocr-onnxruntime + pypdfium2 + Pillow（图片与扫描件 OCR）、pytest
5. 自检：导入依赖、检查图形界面与 OCR 组件、生成演示数据

所有中文提示都在本脚本内输出，避免批处理的编码问题。
"""

from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
VENV_DIR = ROOT / ".venv"
VENV_PY = VENV_DIR / "Scripts" / "python.exe"
MIRROR_INDEX = "https://pypi.tuna.tsinghua.edu.cn/simple"
STEPS = 4


def say(message: str = "") -> None:
    print(message, flush=True)


def run(args: list[str]) -> int:
    """静默执行子进程，返回退出码。"""
    return subprocess.call(args)


def check_python() -> bool:
    if sys.version_info >= (3, 10):
        say(f"[检查] Python {sys.version.split()[0]}，位置：{sys.executable}")
        return True
    say(f"[错误] 当前 Python 版本为 {sys.version.split()[0]}，需要 3.10 或更高版本。")
    say(f"       位置：{sys.executable}")
    say("       请安装新版本 Python 后重新运行本脚本。")
    return False


def venv_is_usable() -> bool:
    return VENV_PY.is_file() and run([str(VENV_PY), "-c", "import sys"]) == 0


def ensure_venv() -> bool:
    if venv_is_usable():
        say(f"[1/{STEPS}] 虚拟环境已存在且可用，跳过创建")
        return True

    if VENV_DIR.exists():
        say(f"[1/{STEPS}] 检测到不可用的虚拟环境，正在重建 .venv ...")
        shutil.rmtree(VENV_DIR, ignore_errors=True)
    else:
        say(f"[1/{STEPS}] 正在创建虚拟环境 .venv ...")

    if run([sys.executable, "-m", "venv", str(VENV_DIR)]) != 0 or not VENV_PY.is_file():
        say("[错误] 创建虚拟环境失败，请检查磁盘空间与目录权限。")
        return False
    return True


def pip_install(args: list[str]) -> bool:
    """安装依赖；默认源失败时改用镜像源重试一次。"""
    if run([str(VENV_PY), "-m", "pip", "install", *args, "--quiet", "--disable-pip-version-check"]) == 0:
        return True
    say("        默认源安装失败，改用国内镜像重试 ...")
    return (
        run(
            [
                str(VENV_PY),
                "-m",
                "pip",
                "install",
                *args,
                "--disable-pip-version-check",
                "-i",
                MIRROR_INDEX,
            ]
        )
        == 0
    )


def install_requirements() -> bool:
    say(f"[2/{STEPS}] 安装项目依赖（首次执行需要数分钟，请耐心等待）...")
    run([str(VENV_PY), "-m", "pip", "install", "--upgrade", "pip", "--quiet", "--disable-pip-version-check"])
    if not pip_install(["-e", str(ROOT)]):
        say("[错误] 依赖安装失败，请检查网络连接后重新运行本脚本。")
        return False
    return True


def install_extras() -> None:
    say(f"[3/{STEPS}] 安装扩展组件（.xls 读取、OCR 识别、测试工具）...")

    if pip_install(["xlrd"]):
        say("        .xls 读取组件已安装：xlrd")
    else:
        say("        .xls 读取组件安装失败，旧版 .xls 文件将无法读取。")

    if pip_install(["rapidocr-onnxruntime", "pypdfium2", "Pillow"]):
        say("        OCR 组件已安装：可识别图片与扫描版 PDF")
    else:
        say("        OCR 组件安装失败，图片与扫描件将无法识别；其余功能不受影响。")

    if pip_install(["pytest"]):
        say("        测试工具已安装：pytest")
    else:
        say("        测试工具安装失败，不影响软件使用。")


def self_check() -> bool:
    say(f"[4/{STEPS}] 正在自检 ...")

    if run([str(VENV_PY), "-c", "import pandas, openpyxl, docx, pypdf"]) != 0:
        say("[错误] 依赖检查未通过，请重新运行本脚本。")
        return False
    say("        依赖检查通过：pandas / openpyxl / python-docx / pypdf")

    if run([str(VENV_PY), "-c", "import tkinter; tkinter.Tcl()"]) != 0:
        say("        警告：未检测到图形界面组件 tkinter，软件界面可能无法打开")
    else:
        say("        图形界面组件检查通过")

    if run([str(VENV_PY), "-c", "import xlrd"]) == 0:
        say("        .xls 读取组件检查通过：xlrd")
    else:
        say("        提示：未安装 xlrd，旧版 .xls 文件无法读取")

    if run([str(VENV_PY), "-c", "import rapidocr_onnxruntime, pypdfium2, PIL"]) == 0:
        say("        OCR 组件检查通过：rapidocr-onnxruntime / pypdfium2 / Pillow")
    else:
        say("        提示：未安装 OCR 组件，图片与扫描版 PDF 无法识别")

    run([str(VENV_PY), str(ROOT / "scripts" / "make_demo_data.py")])
    return True


def main() -> int:
    say("=" * 62)
    say("  基层办公自动化助手 - 一键环境配置")
    say("=" * 62)
    say()

    if not check_python():
        return 1
    if not ensure_venv():
        return 1
    if not install_requirements():
        return 1
    install_extras()
    if not self_check():
        return 1

    say()
    say("=" * 62)
    say("  环境配置完成")
    say()
    say("  下一步：双击「启动界面.bat」打开软件。")
    say("  建议先在「报表汇总」页用 output 目录中的演示数据试一次。")
    say("=" * 62)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
