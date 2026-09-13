"""按标签页抓取主窗口截图，用于 README 的「界面预览」。

用法::

    python scripts/make_screenshots.py                 # 输出到 docs/images/gui-<页名>.png
    python scripts/make_screenshots.py --out 临时目录    # 换个输出目录
    python scripts/make_screenshots.py --tabs report,upload

界面改版后重新运行即可，各页文件名固定，README 中的引用无需改动。

注意事项（均已在脚本内处理）：

- 抓图坐标必须与屏幕物理像素一致，故先把进程设为 DPI 感知（否则高缩放比下会截偏）；
- 取外框矩形而非客户区（`GetAncestor(..., GA_ROOT)` + `GetWindowRect`），保证标题栏一起入图；
- 截当前标签页前要 `update()` 并短暂停留，否则截到未重绘的旧画面。
"""

from __future__ import annotations

import argparse
import ctypes
import sys
import time
from pathlib import Path
from typing import Sequence

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

DEFAULT_OUT = ROOT / "docs" / "images"
# 标签页顺序与界面一致：报表汇总、格式转换、合并文档、批量提取、文件归档、自动上传
TABS = {
    "report": 0,
    "convert": 1,
    "merge": 2,
    "scan": 3,
    "classify": 4,
    "upload": 5,
}
GA_ROOT = 2


class Rect(ctypes.Structure):
    """Win32 RECT。"""

    _fields_ = [
        ("left", ctypes.c_long),
        ("top", ctypes.c_long),
        ("right", ctypes.c_long),
        ("bottom", ctypes.c_long),
    ]


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="抓取界面截图")
    parser.add_argument("--out", type=Path, default=DEFAULT_OUT, help=f"输出目录（默认 {DEFAULT_OUT}）")
    parser.add_argument("--tabs", default=",".join(TABS), help="只抓指定页，逗号分隔")
    parser.add_argument("--size", default="1040x840+80+40", help="窗口几何，默认 1040x840+80+40")
    options = parser.parse_args(argv)

    try:
        from PIL import ImageGrab
    except ImportError:
        print("缺少 Pillow，请先执行：python -m pip install Pillow")
        return 1

    from information_mapper import gui

    ctypes.windll.user32.SetProcessDPIAware()  # 非 Windows 环境下需另行适配
    options.out.mkdir(parents=True, exist_ok=True)

    app = gui.OfficeAssistantApp()
    app.geometry(options.size)
    app.attributes("-topmost", True)  # 避免被其他窗口遮挡
    app.update()
    time.sleep(1.0)  # 等启动耗时提示写入消息区
    app.update()

    window = ctypes.windll.user32.GetAncestor(app.winfo_id(), GA_ROOT)
    rect = Rect()
    ctypes.windll.user32.GetWindowRect(window, ctypes.byref(rect))
    box = (rect.left, rect.top, rect.right, rect.bottom)
    print(f"窗口外框：{box}，屏幕：{app.winfo_screenwidth()}x{app.winfo_screenheight()}")

    for name in [item.strip() for item in options.tabs.split(",") if item.strip()]:
        if name not in TABS:
            print(f"跳过未知页面：{name}（可选：{'、'.join(TABS)}）")
            continue
        app.notebook.select(TABS[name])
        app.update()
        time.sleep(0.6)  # 等标签页重绘完成
        path = options.out / f"gui-{name}.png"
        image = ImageGrab.grab(bbox=box)
        image.save(path)
        print(f"{name}: {image.size[0]}x{image.size[1]}，{path.stat().st_size / 1024:.0f} KB -> {path}")

    app.destroy()
    print("\n完成。README 引用的是 docs/images/gui-<页名>.png，文件名保持固定即可。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
