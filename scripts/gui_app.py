"""图形界面启动脚本。

用法::

    python scripts/gui_app.py                  # 启动界面
    python scripts/gui_app.py 报表.xlsx         # 启动并把文件带入输入列表
    python scripts/gui_app.py --self-check      # 无界面自检（核对依赖与手册，结果写 self-check.txt）
    双击项目根目录的「启动界面.bat」             # 非技术用户推荐

也可以把文件或文件夹直接拖到此脚本上，会自动填入第一个输入框。
"""

from __future__ import annotations

import sys
import tempfile
from pathlib import Path

if not getattr(sys, "frozen", False):  # 打包后由 PyInstaller 提供模块路径
    sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))


def fatal(title: str, message: str) -> None:
    """启动失败提示：有控制台时输出到 stderr，窗口版则弹系统消息框。"""
    print(f"{title}：{message}", file=sys.stderr)
    if sys.platform.startswith("win") and sys.stdout is None:  # 打包后的窗口版没有控制台
        try:
            import ctypes

            ctypes.windll.user32.MessageBoxW(None, message, title, 0x10)
        except Exception:  # pragma: no cover - 极端环境下消息框也不可用
            pass


def self_check() -> int:
    """无界面自检：核对运行环境、组件与打包进去的资源。

    打包后的窗口版没有控制台，结果写入当前目录的 ``self-check.txt``，
    由 ``scripts/make_release.py`` 在构建后读取判据。
    """
    import importlib.util as importlib_util

    from information_mapper import __version__, readers, gui

    lines = [
        f"版本：{__version__}",
        f"Python：{sys.version.split()[0]}",
        f"打包运行：{'是' if getattr(sys, 'frozen', False) else '否'}",
        f"资源目录：{getattr(sys, '_MEIPASS', '（未打包）')}",
    ]
    failures: list[str] = []

    manual = gui.manual_path()
    lines.append(f"用户手册：{manual if manual else '未找到'}")
    if manual is None:
        failures.append("未找到用户手册")

    modules = {
        "pandas": "pandas",
        "openpyxl": "openpyxl",
        "python-docx": "docx",
        "pypdf": "pypdf",
        "xlrd": "xlrd",
        "Pillow": "PIL",
    }
    for label, module in modules.items():
        available = importlib_util.find_spec(module) is not None
        lines.append(f"{label}：{'已安装' if available else '缺失'}")
        if not available:
            failures.append(f"缺少 {label}")

    lines.append(f"OCR 组件：{'已安装' if readers.ocr_available() else '缺失'}")
    if not readers.ocr_available():
        failures.append("缺少 OCR 组件")

    models = _resource_dir("rapidocr_onnxruntime") / "models"
    model_files = sorted(models.glob("*.onnx")) if models.is_dir() else []
    lines.append(f"OCR 模型：{len(model_files)} 个（{models}）")
    if not model_files:
        failures.append("OCR 模型未随包收集")

    pdfium = _resource_dir("pypdfium2_raw") / "pdfium.dll"
    lines.append(f"PDF 渲染库：{'存在' if pdfium.is_file() else '缺失'}（{pdfium}）")
    if not pdfium.is_file():
        failures.append("缺少 pdfium.dll")

    if model_files:
        try:
            readers._ocr_engine()  # noqa: SLF001 - 自检需真正加载一次模型，确认 onnxruntime 可用
            lines.append("OCR 引擎加载：成功")
        except Exception as error:  # pragma: no cover - 仅在打包缺件时触发
            lines.append(f"OCR 引擎加载：失败（{type(error).__name__}: {error}）")
            failures.append("OCR 引擎无法加载")
        else:
            try:  # 跑一次完整识别链路：Pillow 写图 → cv2 读图 → onnxruntime 推理
                sample = Path(tempfile.gettempdir()) / "info-map-self-check.png"
                make_sample_image(sample)
                records = readers.read_image(sample)
                sample.unlink(missing_ok=True)
                lines.append(f"OCR 识别：成功（返回 {len(records)} 条记录）")
            except Exception as error:  # pragma: no cover - 仅在打包缺件时触发
                lines.append(f"OCR 识别：失败（{type(error).__name__}: {error}）")
                failures.append("OCR 识别链路异常")

    try:  # 真正搭一次界面：窗口版出错不会显示堆栈，只能在这里提前发现
        app = gui.OfficeAssistantApp()
        app.geometry("1040x840+4000+4000")  # 移到可视区域外，避免自检时闪窗
        app.update()
        tab_count = len(app.notebook.tabs())
        app.destroy()
        lines.append(f"界面构建：成功（{tab_count} 个标签页）")
    except Exception as error:
        lines.append(f"界面构建：失败（{type(error).__name__}: {error}）")
        failures.append("界面无法构建")

    lines.append("")
    lines.append("自检结果：通过" if not failures else "自检结果：未通过 —— " + "；".join(failures))
    report = Path.cwd() / "self-check.txt"
    report.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return 0 if not failures else 1


def _resource_dir(package: str) -> Path:
    """取包所在目录（打包后指向解包目录，用于核对数据文件是否被收集）。"""
    import importlib.util as importlib_util

    spec = importlib_util.find_spec(package)
    locations = list(spec.submodule_search_locations or []) if spec else []
    return Path(locations[0]) if locations else Path()


def make_sample_image(path: Path) -> None:
    """生成一张带大号文字的图片，供自检跑通 OCR 链路（文件写系统临时目录）。"""
    from PIL import Image, ImageDraw, ImageFont

    image = Image.new("RGB", (520, 160), "white")
    draw = ImageDraw.Draw(image)
    try:
        font = ImageFont.truetype("arial.ttf", 56)
    except OSError:  # pragma: no cover - 无该字体时退回内置位图字体
        font = ImageFont.load_default()
    draw.text((30, 50), "OCR CHECK 2026", fill="black", font=font)
    image.save(path)


def main() -> int:
    if "--self-check" in sys.argv[1:]:
        return self_check()
    try:
        from information_mapper.gui import main as run_gui
    except ImportError as error:  # 极少数精简版 Python 不带 tkinter
        fatal(
            "启动失败",
            f"{error}\n\n本界面依赖 Python 自带的 tkinter，"
            "请使用完整版 Python（Windows 官方安装包默认包含）。",
        )
        return 1
    return run_gui()


if __name__ == "__main__":
    raise SystemExit(main())
