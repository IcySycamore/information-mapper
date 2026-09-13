"""图形界面启动脚本。

用法::

    python scripts/gui_app.py            # 命令行启动
    双击项目根目录的「启动界面.bat」      # 非技术用户推荐

也可以把文件或文件夹直接拖到此脚本上，会自动填入第一个输入框。
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))


def main() -> int:
    try:
        from information_mapper.gui import main as run_gui
    except ImportError as error:  # 极少数精简版 Python 不带 tkinter
        print(f"启动失败：{error}")
        print("本界面依赖 Python 自带的 tkinter，请使用完整版 Python（Windows 官方安装包默认包含）。")
        input("按回车键退出…")
        return 1
    return run_gui()


if __name__ == "__main__":
    raise SystemExit(main())
