"""图形界面（tkinter，无需额外依赖）。

窗口布局：

    ┌───────────────────────────────┬──────────────┐
    │ 标签页（功能页）               │ 右栏         │
    │                               │  当前功能说明 │
    │                               │  消息与过程   │
    ├───────────────────────────────┴──────────────┤
    │ 状态栏（固定高度，单行）                       │
    └──────────────────────────────────────────────┘

- 状态栏只显示四类简短状态：``就绪`` / ``需要…`` / ``成功 · …`` / ``失败 · …``（以及启动耗时），
  除致命错误外都会在数秒后自动回到「就绪」。
- 完整的操作说明、历史提示与处理明细统一放在右栏，切换标签页时右栏说明同步更新。
- 界面层只收集参数，真正干活的是下面的 ``run_*`` 服务函数（不依赖 tkinter，可单独测试）。
- 耗时任务在后台线程执行；消息按类型着色（成功 / 失败 / 明细），数量过多时自动丢弃最早的记录。

启动方式::

    python scripts/gui_app.py      # 或双击项目根目录的「启动界面.bat」
"""

from __future__ import annotations

import copy
import importlib
import json
import os
import platform
import queue
import re
import subprocess
import sys
import threading
import time
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, font as tkfont, ttk
from typing import Any, Callable, Iterable, Sequence

from . import __version__, rules
from .classify import ClassifyConfig
from .classify import organize as organize_files
from .classify import summarize as summarize_moves
from .converters import convert_file, convert_folder, normalize_target_format
from .core import map_records
from .formats import INPUT_SUFFIXES, support_matrix
from .merger import collect_headers, merge_any, merge_tables, resolve_inputs
from .readers import SOURCE_COLUMN, iter_input_files, read_folder
from .uploader import UploadConfig, ping_server, upload_paths
from .uploader import summarize as summarize_uploads
from .writers import write_output

LogFunc = Callable[[str], None]

# ---------------------------------------------------------------- 常量
STATUS_READY = "就绪"
STATUS_RESET_MS = 4000  # 普通提示在几秒后自动回到「就绪」
POLL_INTERVAL_MS = 80  # 后台消息轮询间隔
SIDE_PANEL_WIDTH = 300  # 右栏固定宽度（像素）
MESSAGE_MAX_LINES = 500  # 右栏消息最多保留的行数
MESSAGE_TRIM_LINES = 150  # 超出上限时一次丢弃的行数

STATUS_STYLES = {
    "info": "Status.TLabel",
    "success": "StatusSuccess.TLabel",
    "error": "StatusError.TLabel",
}

MERGE_MODE_LABELS: dict[str, str] = {
    "auto": "自动判断",
    "table": "汇总成一张表",
    "word": "Word 按顺序拼接",
    "text": "正文按顺序拼接",
}
MERGE_MODE_VALUES: dict[str, str] = {label: value for value, label in MERGE_MODE_LABELS.items()}

UPLOAD_METHOD_LABELS: dict[str, str] = {
    "multipart": "multipart 表单上传",
    "json": "json + Base64",
}
UPLOAD_METHOD_VALUES: dict[str, str] = {label: value for value, label in UPLOAD_METHOD_LABELS.items()}

CONVERT_TARGETS = ("xlsx", "csv", "docx", "txt", "md", "json", "html")

TAB_REPORT = "报表汇总"
TAB_CONVERT = "格式转换"
TAB_MERGE = "合并文档"
TAB_SCAN = "批量提取"
TAB_CLASSIFY = "文件归档"
TAB_UPLOAD = "自动上传"
TAB_MANUAL = "使用说明"
TAB_FORMATS = "支持的格式"

TAB_ORDER = (
    TAB_REPORT,
    TAB_CONVERT,
    TAB_MERGE,
    TAB_SCAN,
    TAB_CLASSIFY,
    TAB_UPLOAD,
    TAB_MANUAL,
    TAB_FORMATS,
)

# 右栏显示的当前功能说明（切换标签页时同步）
TAB_HINTS: dict[str, str] = {
    TAB_REPORT: "添加一份或多份报表 → 填写输出文件 → 点「开始汇总」。\n"
    "输出会多一列「源文件」，标明每条数据来自哪份报表。\n"
    "需要统一列名时，点「字段映射」右侧的「设置…」直接在软件里编辑。",
    TAB_CONVERT: "添加文件或文件夹 → 选择目标格式 → 选择输出文件夹 → 点「开始转换」。\n"
    "同名文件不会被覆盖，会自动加序号。",
    TAB_MERGE: "添加要合并的文档 → 选择合并方式 → 填写输出文件 → 点「开始合并」。\n"
    "输出 .xlsx/.csv 时汇总成一张表；输出 .docx/.md/.txt 时按顺序拼接内容。",
    TAB_SCAN: "选择待扫描目录 → 填写输出文件 → 点「开始提取」。\n"
    "程序读取文档中的「字段：内容」与表格；解析失败的文件会被跳过并在消息区说明。\n"
    "需要统一表头时，点「字段映射」右侧的「设置…」编辑规则。",
    TAB_CLASSIFY: "选择待整理目录 → 点「归档规则」右侧的「设置…」编辑规则 → 点「开始归档」。\n"
    "规则可直接在软件中增删改，保存后自动生成配置文件；移动明细见右栏消息。",
    TAB_UPLOAD: "添加要上传的文件 → 填写接口地址与密钥 → 先点「Ping 测试」确认接口可达，"
    "再点「开始上传」。Ping 只发一个轻量请求，不会上传文件。",
    TAB_MANUAL: "此处显示完整说明书，亦可用系统默认程序打开文档查看；遇到问题点「技术支持」查看联系方式。",
    TAB_FORMATS: "列出各文件后缀的说明，以及能否作为输入（读）或输出（写）。",
}

MANUAL_FILENAME = "USER_GUIDE.md"

# 技术支持：联系方式（可按实际情况修改）
SUPPORT_CONTACTS: dict[str, str] = {
    "QQ": "1284742412",
    "GitHub": "IcySycamore",
    "邮箱": "1284742412@qq.com",
}
SUPPORT_REPOSITORY = "https://github.com/IcySycamore/information-mapper"

_FALLBACK_MANUAL = """# 用户手册（未找到文档）

未在 docs 目录下找到 USER_GUIDE.md，已显示简版说明。

## 操作流程
1. 双击「启动界面.bat」打开软件；
2. 在上方标签页中选择功能；
3. 按右栏提示添加文件、填写输出位置；
4. 点页面下方的执行按钮；过程与结果见右栏消息与状态栏。

## 各标签页用途
- 报表汇总：多份 Excel/CSV 报表合并为一张总表
- 格式转换：文档在 xlsx / csv / docx / txt / md / json / html 之间互转
- 合并文档：大批量文档合并为一个文档
- 批量提取：从 Word/PDF/Excel 中提取信息成表
- 文件归档：按规则把文件分类到子目录
- 自动上传：把结果文件上传到接口（可先做连通性测试）
- 支持的格式：查看可读写的格式
"""


# ============================================================ 服务层（与界面无关）
def split_patterns(text: str) -> list[str] | None:
    """把 ``*.xlsx, *.csv`` 这类输入解析成通配符列表；空文本返回 ``None``（表示不过滤）。"""
    items = [item.strip() for item in re.split(r"[,，;；\s]+", text or "") if item.strip()]
    return items or None


def load_mapping(path: str | Path | None) -> dict[str, Any]:
    """读取字段映射配置；未提供时返回空字典。"""
    if not path:
        return {}
    config_path = Path(path)
    if not config_path.exists():
        raise ValueError(f"找不到字段映射配置：{config_path}")
    return json.loads(config_path.read_text(encoding="utf-8"))


def supported_filetypes() -> list[tuple[str, str]]:
    """文件选择对话框用的类型过滤器。"""
    patterns = " ".join(f"*{suffix}" for suffix in sorted(INPUT_SUFFIXES))
    return [("所有受支持文档", patterns), ("所有文件", "*.*")]


def environment_summary() -> dict[str, str]:
    """收集用于技术支持的运行环境信息。"""
    dependencies = (
        ("pandas", "pandas"),
        ("python-docx", "docx"),
        ("pypdf", "pypdf"),
        ("Pillow", "PIL"),
        ("rapidocr-onnxruntime", "rapidocr_onnxruntime"),
        ("pypdfium2", "pypdfium2"),
        ("xlrd", "xlrd"),
    )
    installed: list[str] = []
    for label, module_name in dependencies:
        try:
            module = importlib.import_module(module_name)
        except ImportError:
            continue
        version = getattr(module, "__version__", "已安装")
        installed.append(f"{label} {version}")
    return {
        "版本": __version__,
        "Python": platform.python_version(),
        "系统": f"{platform.system()} {platform.release()}",
        "依赖": " / ".join(installed) or "未检测到",
    }


def support_text(contacts: dict[str, str] | None = None, environment: dict[str, str] | None = None) -> str:
    """拼装技术支持信息文本（供弹窗显示与复制）。"""
    contact_items = contacts or SUPPORT_CONTACTS
    environment_items = environment or environment_summary()
    lines = [
        f"基层办公自动化助手 v{environment_items.get('版本', __version__)}",
        "",
        "技术支持联系方式",
    ]
    lines.extend(f"  {label}：{value}" for label, value in contact_items.items())
    lines.extend(["", "运行环境"])
    lines.extend(f"  {label}：{value}" for label, value in environment_items.items())
    lines.extend(
        [
            "",
            "反馈方式",
            f"  1. 在项目主页提交 issue：{SUPPORT_REPOSITORY}/issues",
            "  2. 也可通过上面的 QQ 或邮箱联系",
            "",
            "本页内容为联系方式与运行环境，不含报错信息。",
            "反馈时请附上本页信息，并附上右栏「消息」区中对应的报错记录",
            "（在右栏消息中选中相关文本，按 Ctrl+C 复制）。",
        ]
    )
    return "\n".join(lines)


def manual_path() -> Path | None:
    """定位说明书文档：优先项目 docs 目录，其次当前工作目录。"""
    candidates = [
        Path(__file__).resolve().parents[2] / "docs" / MANUAL_FILENAME,
        Path.cwd() / "docs" / MANUAL_FILENAME,
        Path(__file__).resolve().parent / MANUAL_FILENAME,
    ]
    for candidate in candidates:
        if candidate.is_file():
            return candidate
    return None


def read_manual() -> tuple[str, Path | None]:
    """读取说明书内容；找不到文件时返回内置简版。"""
    path = manual_path()
    if path is None:
        return _FALLBACK_MANUAL, None
    return path.read_text(encoding="utf-8"), path


def open_document(path: str | Path) -> bool:
    """用系统默认程序打开文档（Windows 用 startfile，macOS/Linux 用对应命令）。"""
    target = str(Path(path))
    try:
        if sys.platform.startswith("win"):
            os.startfile(target)  # type: ignore[attr-defined]  # noqa: S606
        elif sys.platform == "darwin":
            subprocess.Popen(["open", target])  # noqa: S603, S607
        else:
            subprocess.Popen(["xdg-open", target])  # noqa: S603, S607
    except OSError:
        return False
    return True


def run_report(
    inputs: Sequence[str | Path],
    output: str | Path,
    *,
    mapping_path: str | Path | None = None,
    patterns: Iterable[str] | None = None,
    recursive: bool = True,
    log: LogFunc = print,
) -> str:
    """报表统计核算：多份报表 → 一张总表。"""
    files = resolve_inputs(inputs, patterns=patterns, recursive=recursive)
    if not files:
        raise ValueError("没有找到可汇总的报表文件，请检查输入目录与文件类型")

    log(f"找到 {len(files)} 份报表，开始汇总…")
    mapping = load_mapping(mapping_path)
    result = merge_tables(
        files,
        output,
        headers=mapping.get("target_headers"),
        title="报表汇总",
        sheet_name="汇总",
    )
    for path, reason in result.failed:
        log(f"  跳过 {path.name}：{reason}")
    log(f"完成：{result.record_count} 条记录，来自 {len(result.sources)} 份报表")
    log(f"输出文件：{result.output.resolve()}")
    return f"汇总完成：{result.record_count} 条记录 → {result.output.name}"


def run_convert(
    inputs: Sequence[str | Path],
    output_dir: str | Path,
    target_format: str,
    *,
    mapping_path: str | Path | None = None,
    patterns: Iterable[str] | None = None,
    recursive: bool = True,
    overwrite: bool = False,
    log: LogFunc = print,
) -> str:
    """常见文档格式转换（单文件或整目录）。"""
    suffix = normalize_target_format(target_format)
    mapping = load_mapping(mapping_path) or None
    destination = Path(output_dir)
    results = []

    for item in inputs:
        path = Path(item)
        if path.is_dir():
            log(f"扫描目录：{path}")
            results.extend(
                convert_folder(
                    path,
                    destination,
                    suffix,
                    mapping=mapping,
                    recursive=recursive,
                    patterns=patterns,
                    overwrite=overwrite,
                )
            )
        elif path.is_file():
            target = destination / f"{path.stem}{suffix}"
            results.append(convert_file(path, target, mapping=mapping))
        else:
            log(f"  跳过不存在的路径：{path}")

    if not results:
        raise ValueError("没有找到可转换的文件")

    failed = [item for item in results if not item.ok]
    for item in failed:
        log(f"  失败 {item.source.name}：{item.error}")
    ok = len(results) - len(failed)
    log(f"完成：成功 {ok} 个，失败 {len(failed)} 个")
    log(f"输出目录：{destination.resolve()}")
    return f"转换完成：成功 {ok} 个，失败 {len(failed)} 个"


def run_merge(
    inputs: Sequence[str | Path],
    output: str | Path,
    *,
    mode: str = "auto",
    mapping_path: str | Path | None = None,
    patterns: Iterable[str] | None = None,
    recursive: bool = True,
    source_column: str | None = SOURCE_COLUMN,
    log: LogFunc = print,
) -> str:
    """大批量输入文档 → 一个输出文档。"""
    files = resolve_inputs(inputs, patterns=patterns, recursive=recursive)
    if not files:
        raise ValueError("没有找到可合并的文档，请检查输入目录与文件类型")

    log(f"准备合并 {len(files)} 份文档（方式：{MERGE_MODE_LABELS.get(mode, mode)}）…")
    mapping = load_mapping(mapping_path)
    result = merge_any(
        files,
        output,
        mode=mode,
        headers=mapping.get("target_headers"),
        source_column=source_column,
        title=Path(output).stem,
    )
    for path, reason in result.failed:
        log(f"  跳过 {path.name}：{reason}")
    log(result.summary())
    log(f"输出文件：{result.output.resolve()}")
    return f"合并完成：{len(result.sources)} 份文档 → {result.output.name}"


def run_scan(
    directory: str | Path,
    output: str | Path,
    *,
    mapping_path: str | Path | None = None,
    patterns: Iterable[str] | None = None,
    recursive: bool = True,
    log: LogFunc = print,
) -> str:
    """批量解析目录中的文档，提取信息汇总成一张表。"""
    failed: list[tuple[Path, str]] = []
    records = read_folder(directory, recursive=recursive, patterns=patterns, failed=failed)
    if not records:
        raise ValueError("没有解析到任何记录，请确认目录中有受支持的文档")

    mapping = load_mapping(mapping_path)
    if mapping.get("target_headers"):
        headers = list(mapping["target_headers"])
        records = map_records(records, headers, mapping.get("field_mapping") or {})
    else:
        headers = collect_headers(records)

    for path, reason in failed:
        log(f"  跳过 {path.name}：{reason}")

    output_path = write_output(records, headers, output, title="信息汇总")
    log(f"完成：共 {len(records)} 条记录")
    log(f"输出文件：{output_path.resolve()}")
    return f"提取完成：{len(records)} 条记录 → {output_path.name}"


def run_classify(
    directory: str | Path,
    config_path: str | Path,
    *,
    root: str | Path | None = None,
    recursive: bool = True,
    log: LogFunc = print,
) -> str:
    """按规则自动归档文件（直接执行移动）。"""
    config = ClassifyConfig.from_file(config_path)
    plans = organize_files(directory, config, dry_run=False, root=root or None, recursive=recursive)
    if not plans:
        log("没有需要归档的文件")
        return "没有需要归档的文件"

    for plan in plans:
        log(f"  {plan.summary()}")
    summary = summarize_moves(plans, dry_run=False)
    log(summary)
    moved = sum(1 for plan in plans if plan.applied)
    return f"归档完成：{moved}/{len(plans)} 个文件已归类"


def run_upload(
    paths: Sequence[str | Path],
    *,
    url: str,
    token: str,
    method: str = "multipart",
    field_name: str = "file",
    patterns: Iterable[str] | None = None,
    log: LogFunc = print,
) -> str:
    """把结果文件上传到指定接口。"""
    config = UploadConfig(url=url.strip(), token=token.strip(), method=method, field_name=field_name)
    config.validate()  # 地址/密钥有问题时直接报错，由界面显示在状态栏

    targets: list[Path] = []
    for item in paths:
        path = Path(item)
        if path.is_dir():
            targets.extend(iter_input_files(path, recursive=False, patterns=patterns))
        elif path.is_file():
            targets.append(path)
        else:
            log(f"  跳过不存在的路径：{path}")

    if not targets:
        raise ValueError("没有找到可上传的文件")

    log(f"接口地址：{config.url}")
    log(f"开始上传 {len(targets)} 个文件…")
    results = upload_paths(targets, config)
    for item in results:
        log(f"  {item.summary()}")
    ok = sum(1 for item in results if item.ok)
    log(summarize_uploads(results))
    return f"上传完成：成功 {ok}/{len(results)} 个文件"


def ping_upload(url: str, token: str = "", *, method: str = "multipart", timeout: float = 5.0) -> str:
    """测试上传接口是否可达，返回可直接显示在状态栏的文字。"""
    config = UploadConfig(url=url.strip(), token=token.strip(), method=method)
    return ping_server(config, timeout=timeout).summary()


# ============================================================ 界面控件
class InputPicker(ttk.Frame):
    """可添加多个"文件或目录"的输入选择器。"""

    def __init__(self, master: tk.Misc, *, height: int = 3) -> None:
        super().__init__(master)
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        self.listbox = tk.Listbox(self, height=height, selectmode=tk.EXTENDED, activestyle="none")
        self.listbox.grid(row=0, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.listbox.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.listbox.configure(yscrollcommand=scrollbar.set)

        buttons = ttk.Frame(self)
        buttons.grid(row=1, column=0, columnspan=2, sticky="w", pady=(6, 0))
        ttk.Button(buttons, text="添加文件…", command=self.add_files).pack(side=tk.LEFT)
        ttk.Button(buttons, text="添加文件夹…", command=self.add_directory).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Button(buttons, text="移除所选", command=self.remove_selected).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Button(buttons, text="清空", command=self.clear).pack(side=tk.LEFT, padx=(6, 0))

    def add_files(self) -> None:
        selected = filedialog.askopenfilenames(title="选择文件", filetypes=supported_filetypes())
        self._extend(selected)

    def add_directory(self) -> None:
        directory = filedialog.askdirectory(title="选择文件夹")
        if directory:
            self._extend([directory])

    def remove_selected(self) -> None:
        for index in reversed(self.listbox.curselection()):
            self.listbox.delete(index)

    def clear(self) -> None:
        self.listbox.delete(0, tk.END)

    def paths(self) -> list[str]:
        return list(self.listbox.get(0, tk.END))

    def _extend(self, items: Iterable[str]) -> None:
        existing = set(self.paths())
        for item in items:
            if item and item not in existing:
                self.listbox.insert(tk.END, item)
                existing.add(item)

    def set_paths(self, items: Iterable[str]) -> None:
        self.clear()
        self._extend(items)


class PathRow(ttk.Frame):
    """一行"标签 + 输入框 + 浏览按钮"，用于选择单个文件或目录。"""

    def __init__(
        self,
        master: tk.Misc,
        label: str,
        *,
        mode: str = "save",
        default: str = "",
        title: str = "选择路径",
        browse: bool = True,
        extra: tuple[str, Callable[[], None]] | None = None,
    ) -> None:
        super().__init__(master)
        self.mode = mode
        self.columnconfigure(1, weight=1)
        self.variable = tk.StringVar(value=default)
        ttk.Label(self, text=f"{label}：", width=13, anchor="w").grid(row=0, column=0, sticky="w")
        ttk.Entry(self, textvariable=self.variable).grid(row=0, column=1, sticky="ew", padx=(0, 6))
        column = 2
        if browse:
            ttk.Button(self, text="浏览…", width=8, command=self.browse).grid(row=0, column=column)
            column += 1
        if extra is not None:
            extra_text, extra_command = extra
            ttk.Button(self, text=extra_text, width=8, command=extra_command).grid(
                row=0, column=column, padx=(4, 0)
            )
        self._title = title

    def browse(self) -> None:
        if self.mode == "dir":
            selected = filedialog.askdirectory(title=self._title)
        elif self.mode == "open":
            selected = filedialog.askopenfilename(
                title=self._title,
                filetypes=[("JSON 配置", "*.json"), ("所有文件", "*.*")],
            )
        else:
            selected = filedialog.asksaveasfilename(title=self._title, defaultextension="")
        if selected:
            self.variable.set(selected)

    def get(self) -> str:
        return self.variable.get().strip()

    def set(self, value: str) -> None:
        self.variable.set(value)


class ScrollableTab(ttk.Frame):
    """带垂直滚动的标签页容器（用于控件较多的表单页）。"""

    def __init__(self, master: tk.Misc) -> None:
        super().__init__(master)
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        style = ttk.Style(self)
        background = style.lookup("TFrame", "background") or "#f0f0f0"
        self.canvas = tk.Canvas(self, highlightthickness=0, borderwidth=0, background=background)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.canvas.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.canvas.configure(yscrollcommand=scrollbar.set)

        self.content = ttk.Frame(self.canvas, padding=12)
        self._window = self.canvas.create_window((0, 0), window=self.content, anchor="nw")
        self.content.bind("<Configure>", self._sync_scrollregion)
        self.canvas.bind("<Configure>", self._sync_width)

    def _sync_scrollregion(self, _event: tk.Event | None = None) -> None:
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _sync_width(self, event: tk.Event) -> None:
        self.canvas.itemconfigure(self._window, width=event.width)

    def bind_wheel(self, widget: tk.Misc | None = None) -> None:
        """给容器内所有控件绑定鼠标滚轮（子控件是动态创建的，需构建完再调用）。"""
        target = widget or self.content
        target.bind("<MouseWheel>", self._on_wheel, add="+")
        for child in target.winfo_children():
            self.bind_wheel(child)

    def _on_wheel(self, event: tk.Event) -> None:
        box = self.canvas.bbox("all")
        if not box or box[3] <= self.canvas.winfo_height():
            return  # 内容没超出可视区域，不需要滚动
        self.canvas.yview_scroll(-1 if event.delta > 0 else 1, "units")


# ============================================================ 规则编辑对话框
def _activate_modal(window: tk.Toplevel) -> None:
    """置为模态并聚焦；无显示环境（如自动化测试）下忽略失败。"""
    try:
        window.grab_set()
        window.focus_force()
    except tk.TclError:  # pragma: no cover - 窗口尚未映射
        pass


class SupportDialog(tk.Toplevel):
    """技术支持弹窗：联系方式与运行环境，可一键复制。"""

    def __init__(self, master: tk.Misc, *, text: str) -> None:
        super().__init__(master)
        self.title("技术支持")
        self.transient(master)
        self.geometry("560x460")
        self.minsize(480, 360)
        self.copied = False

        body = ttk.Frame(self, padding=12)
        body.pack(fill=tk.BOTH, expand=True)

        holder = ttk.Frame(body)
        holder.pack(fill=tk.BOTH, expand=True)
        self._text = tk.Text(
            holder, wrap=tk.WORD, relief=tk.FLAT, background="#fbfbfb", padx=10, pady=8
        )
        scrollbar = ttk.Scrollbar(holder, orient=tk.VERTICAL, command=self._text.yview)
        self._text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self._text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self._text.insert("1.0", text)
        self._text.configure(state=tk.DISABLED)
        self._content = text

        footer = ttk.Frame(body)
        footer.pack(fill=tk.X, pady=(8, 0))
        self._copy_button = ttk.Button(footer, text="复制信息", command=self._copy)
        self._copy_button.pack(side=tk.LEFT)
        ttk.Button(footer, text="关闭", command=self.destroy).pack(side=tk.RIGHT)

        self.bind("<Escape>", lambda _event: self.destroy())
        _activate_modal(self)

    def content(self) -> str:
        """当前展示的信息文本。"""
        return self._content

    def _copy(self) -> None:
        self.clipboard_clear()
        self.clipboard_append(self._content)
        self.copied = True
        self._copy_button.configure(text="已复制")


class _FieldsDialog(tk.Toplevel):
    """按字段逐项编辑一条规则。"""

    def __init__(
        self,
        master: tk.Misc,
        *,
        title: str,
        labels: Sequence[str],
        values: Sequence[str] = (),
    ) -> None:
        super().__init__(master)
        self.title(title)
        self.transient(master)
        self.resizable(False, False)
        self.result: tuple[str, ...] | None = None

        body = ttk.Frame(self, padding=14)
        body.pack(fill=tk.BOTH, expand=True)
        body.columnconfigure(1, weight=1)

        self._variables: list[tk.StringVar] = []
        for index, label in enumerate(labels):
            ttk.Label(body, text=f"{label}：").grid(row=index, column=0, sticky="w", pady=3)
            variable = tk.StringVar(value=values[index] if index < len(values) else "")
            ttk.Entry(body, textvariable=variable, width=44).grid(
                row=index, column=1, sticky="ew", pady=3, padx=(6, 0)
            )
            self._variables.append(variable)

        footer = ttk.Frame(self, padding=(14, 0, 14, 14))
        footer.pack(fill=tk.X)
        ttk.Button(footer, text="确定", command=self._confirm).pack(side=tk.RIGHT)
        ttk.Button(footer, text="取消", command=self.destroy).pack(side=tk.RIGHT, padx=(0, 8))

        self.bind("<Return>", lambda _event: self._confirm())
        self.bind("<Escape>", lambda _event: self.destroy())
        _activate_modal(self)

    def _confirm(self) -> None:
        self.result = tuple(variable.get().strip() for variable in self._variables)
        self.destroy()


class _TableRuleDialog(tk.Toplevel):
    """规则表格编辑器基类：树形列表 + 增删改 + 顺序调整。"""

    columns: tuple[str, ...] = ()
    widths: tuple[int, ...] = ()
    edit_labels: tuple[str, ...] = ()
    edit_title = "编辑规则"
    hint = ""

    def __init__(self, master: tk.Misc, *, title: str, rows: Sequence[Sequence[str]]) -> None:
        super().__init__(master)
        self.title(title)
        self.transient(master)
        self.geometry("780x470")
        self.minsize(660, 400)
        self.result: dict[str, Any] | None = None
        self._rows: list[list[str]] = [list(row) for row in rows]

        container = ttk.Frame(self, padding=12)
        container.pack(fill=tk.BOTH, expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(1, weight=1)

        self._build_header(container)

        self.tree = ttk.Treeview(container, columns=self.columns, show="headings", height=8)
        for name, width in zip(self.columns, self.widths):
            self.tree.heading(name, text=name)
            self.tree.column(name, width=width, anchor="w")
        self.tree.grid(row=1, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(container, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar.grid(row=1, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.bind("<Double-1>", lambda _event: self._edit())

        buttons = ttk.Frame(container)
        buttons.grid(row=2, column=0, columnspan=2, sticky="w", pady=(8, 0))
        for text, command in (
            ("添加", self._add),
            ("编辑", self._edit),
            ("删除", self._remove),
            ("上移", lambda: self._move(-1)),
            ("下移", lambda: self._move(1)),
        ):
            ttk.Button(buttons, text=text, width=6, command=command).pack(side=tk.LEFT, padx=(0, 6))

        if self.hint:
            ttk.Label(
                container, text=self.hint, foreground="#666666", wraplength=720, justify="left"
            ).grid(row=3, column=0, columnspan=2, sticky="w", pady=(8, 0))

        footer = ttk.Frame(container)
        footer.grid(row=4, column=0, columnspan=2, sticky="e", pady=(10, 0))
        ttk.Button(footer, text="确定", command=self._confirm).pack(side=tk.RIGHT)
        ttk.Button(footer, text="取消", command=self.destroy).pack(side=tk.RIGHT, padx=(0, 8))

        self._refresh()
        _activate_modal(self)

    # -------------------------------------------------- 子类扩展
    def _build_header(self, container: ttk.Frame) -> None:
        """在表格上方插入全局设置项（默认无）。"""

    def collect(self) -> dict[str, Any]:  # pragma: no cover - 由子类实现
        raise NotImplementedError

    # -------------------------------------------------- 表格操作
    def _refresh(self) -> None:
        self.tree.delete(*self.tree.get_children())
        for row in self._rows:
            self.tree.insert("", tk.END, values=row)

    def _selected_index(self) -> int | None:
        selection = self.tree.selection()
        return self.tree.index(selection[0]) if selection else None

    def _add(self) -> None:
        values = self._ask([""] * len(self.columns))
        if values is None:
            return
        self._rows.append(list(values))
        self._refresh()

    def _edit(self) -> None:
        index = self._selected_index()
        if index is None:
            return
        values = self._ask(self._rows[index])
        if values is None:
            return
        self._rows[index] = list(values)
        self._refresh()
        self.tree.selection_set(self.tree.get_children()[index])

    def _remove(self) -> None:
        index = self._selected_index()
        if index is None:
            return
        del self._rows[index]
        self._refresh()

    def _move(self, delta: int) -> None:
        index = self._selected_index()
        if index is None:
            return
        target = index + delta
        if not 0 <= target < len(self._rows):
            return
        self._rows[index], self._rows[target] = self._rows[target], self._rows[index]
        self._refresh()
        self.tree.selection_set(self.tree.get_children()[target])

    def _ask(self, values: Sequence[str]) -> tuple[str, ...] | None:
        dialog = _FieldsDialog(self, title=self.edit_title, labels=self.edit_labels, values=values)
        self.wait_window(dialog)
        return dialog.result

    def _confirm(self) -> None:
        self.result = self.collect()
        self.destroy()


class MappingDialog(_TableRuleDialog):
    """字段映射编辑器。"""

    columns = ("目标字段", "来源字段候选")
    widths = (200, 480)
    edit_labels = ("目标字段名", "来源字段候选（多个用逗号分隔）")
    edit_title = "编辑字段映射"
    hint = (
        "来源字段候选按顺序匹配，命中第一个存在的列名；"
        "例如“联系电话”一行可填：联系电话, 手机号, 电话。"
    )

    def __init__(self, master: tk.Misc, config: dict[str, Any]) -> None:
        super().__init__(master, title="设置字段映射", rows=rules.mapping_rows(config))

    def collect(self) -> dict[str, Any]:
        return rules.build_mapping([(row[0], row[1]) for row in self._rows])


class ClassifyDialog(_TableRuleDialog):
    """归档规则编辑器。"""

    columns = ("规则名", "目标子目录", "扩展名", "文件名关键词")
    widths = (150, 150, 220, 200)
    edit_labels = (
        "规则名",
        "目标子目录名",
        "匹配的扩展名（逗号分隔，如 .xlsx, .csv）",
        "匹配的文件名关键词（逗号分隔，可留空）",
    )
    edit_title = "编辑归档规则"
    hint = "规则自上而下匹配，第一条命中的生效；未命中任何规则的文件归入“默认归类”。"

    def __init__(self, master: tk.Misc, config: dict[str, Any]) -> None:
        self._target_root = tk.StringVar(value=str(config.get("target_root") or "归档"))
        self._default_target = tk.StringVar(value=str(config.get("default_target") or "99-其他"))
        self._copy = tk.BooleanVar(value=bool(config.get("copy", False)))
        super().__init__(master, title="设置归档规则", rows=rules.classify_rows(config))

    def _build_header(self, container: ttk.Frame) -> None:
        header = ttk.Frame(container)
        header.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        ttk.Label(header, text="归档根目录：").pack(side=tk.LEFT)
        ttk.Entry(header, textvariable=self._target_root, width=16).pack(side=tk.LEFT)
        ttk.Label(header, text="　默认归类：").pack(side=tk.LEFT)
        ttk.Entry(header, textvariable=self._default_target, width=14).pack(side=tk.LEFT)
        ttk.Checkbutton(header, text="复制而非移动", variable=self._copy).pack(side=tk.LEFT, padx=(12, 0))

    def collect(self) -> dict[str, Any]:
        return rules.build_classify(
            [(row[0], row[1], row[2], row[3]) for row in self._rows],
            target_root=self._target_root.get(),
            default_target=self._default_target.get(),
            copy=self._copy.get(),
        )


class OfficeAssistantApp(tk.Tk):
    """主窗口。

    右侧为固定宽度的信息栏（当前功能说明 + 消息列表），
    底部为固定高度的单行状态栏。
    """

    def __init__(self, *, startup_started: float | None = None) -> None:
        self._startup_started = startup_started if startup_started is not None else time.perf_counter()
        super().__init__()

        self.title("基层办公自动化助手")
        self.geometry("1040x840")
        self.minsize(900, 560)

        self._queue: queue.Queue[tuple[str, str]] = queue.Queue()
        self._busy = False
        self._run_buttons: list[ttk.Button] = []
        self._scroll_tabs: list[ScrollableTab] = []
        self._status_job: str | None = None
        self._poll_job: str | None = None
        self._after_jobs: list[str] = []
        self.status_reset_ms = STATUS_RESET_MS  # 测试可调小

        self._apply_fonts()
        self._build_layout()
        self._schedule(120, self._announce_startup)
        self._poll_queue()  # 常驻轮询：后台线程的消息统一在界面线程处理

    # -------------------------------------------------- 样式
    def _apply_fonts(self) -> None:
        families = set(tkfont.families(self))
        family = None
        for candidate in ("Microsoft YaHei UI", "Microsoft YaHei", "微软雅黑", "SimHei", "PingFang SC"):
            if candidate in families:
                family = candidate
                for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont"):
                    tkfont.nametofont(name).configure(family=candidate, size=10)
                break
        self._family = family or tkfont.nametofont("TkDefaultFont").cget("family")

        style = ttk.Style(self)
        for theme in ("vista", "winnative", "clam"):
            if theme in style.theme_names():
                style.theme_use(theme)
                break
        style.configure("Run.TButton", font=(self._family, 10, "bold"))
        style.configure("SideTitle.TLabel", font=(self._family, 10, "bold"), foreground="#1f3d7a")
        style.configure("Status.TLabel", foreground="#1f3d7a")
        style.configure("StatusSuccess.TLabel", foreground="#1b6b32")
        style.configure("StatusError.TLabel", foreground="#a4262c")

    # -------------------------------------------------- 布局
    def _build_layout(self) -> None:
        self.columnconfigure(0, weight=1)  # 左侧工作区（可伸缩）
        self.columnconfigure(1, weight=0, minsize=SIDE_PANEL_WIDTH)  # 右栏（固定宽度）
        self.rowconfigure(0, weight=1)  # 主体
        self.rowconfigure(1, weight=0)  # 状态栏（固定高度）

        body = ttk.Frame(self)
        body.grid(row=0, column=0, sticky="nsew", padx=(12, 6), pady=(8, 4))
        header = ttk.Frame(body)
        header.pack(fill=tk.X)
        ttk.Label(header, text="基层办公自动化助手", font=(self._family, 14, "bold")).pack(anchor="w")

        self.notebook = ttk.Notebook(body)
        self.notebook.pack(fill=tk.BOTH, expand=True, pady=(6, 0))
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)

        self._build_side_panel()

        self._build_status_bar()

        self._build_report_tab()
        self._build_convert_tab()
        self._build_merge_tab()
        self._build_scan_tab()
        self._build_classify_tab()
        self._build_upload_tab()
        self._build_manual_tab()
        self._build_formats_tab()

        for scroll in self._scroll_tabs:  # 控件已就位，挂上滚轮支持
            scroll.bind_wheel()

        self._on_tab_changed()  # 初始化右栏说明

    def _build_status_bar(self) -> None:
        """底部状态栏：固定高度、单行、不随文字换行。"""
        self.status = tk.StringVar(value=STATUS_READY)
        self.status_bar = ttk.Label(
            self,
            textvariable=self.status,
            anchor="w",
            style="Status.TLabel",
            padding=(12, 5),
        )
        self.status_bar.grid(row=1, column=0, columnspan=2, sticky="ew")

    def _build_side_panel(self) -> None:
        """右栏：当前功能说明 + 消息列表（固定宽度）。"""
        panel = ttk.Frame(self)
        panel.grid(row=0, column=1, sticky="nsew", padx=(6, 12), pady=(8, 4))
        self.side_panel = panel

        ttk.Label(panel, text="当前功能", style="SideTitle.TLabel").pack(anchor="w")
        self.hint_text = tk.Text(
            panel,
            width=1,
            height=6,
            wrap=tk.WORD,
            state=tk.DISABLED,
            relief=tk.FLAT,
            background="#f4f6fa",
            spacing1=2,
            spacing3=2,
        )
        self.hint_text.pack(fill=tk.X, pady=(2, 10))

        ttk.Label(panel, text="消息", style="SideTitle.TLabel").pack(anchor="w")
        holder = ttk.Frame(panel)
        holder.pack(fill=tk.BOTH, expand=True, pady=(2, 0))
        self.message_text = tk.Text(
            holder, width=1, wrap=tk.WORD, state=tk.DISABLED, relief=tk.FLAT, background="#fbfbfb"
        )
        scrollbar = ttk.Scrollbar(holder, orient=tk.VERTICAL, command=self.message_text.yview)
        self.message_text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.message_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.message_text.tag_configure("success", foreground="#1b6b32")
        self.message_text.tag_configure("error", foreground="#a4262c")
        self.message_text.tag_configure("detail", foreground="#5a5a5a", lmargin1=16, lmargin2=16)

    def _new_tab(self, title: str) -> ttk.Frame:
        """新建可滚动的表单页，返回放置控件的容器。"""
        scroll = ScrollableTab(self.notebook)
        self.notebook.add(scroll, text=title)
        self._scroll_tabs.append(scroll)
        return scroll.content

    def _new_plain_tab(self, title: str) -> ttk.Frame:
        """新建普通页（内部自带滚动的展示页用，内容可占满整个可用区域）。"""
        frame = ttk.Frame(self.notebook, padding=12)
        self.notebook.add(frame, text=title)
        return frame

    def _run_button(self, parent: tk.Misc, text: str, command: Callable[[], None]) -> None:
        button = ttk.Button(parent, text=text, style="Run.TButton", command=command)
        button.pack(anchor="w", pady=(10, 0))
        self._run_buttons.append(button)

    # -------------------------------------------------- 各功能页
    def _build_report_tab(self) -> None:
        tab = self._new_tab(TAB_REPORT)
        self.report_inputs = InputPicker(tab)
        self.report_inputs.pack(fill=tk.BOTH, expand=True)

        self.report_output = PathRow(tab, "输出文件", mode="save", title="保存汇总表")
        self.report_output.pack(fill=tk.X, pady=(8, 2))
        self.report_mapping = PathRow(
            tab,
            "字段映射",
            mode="open",
            title="选择映射配置",
            extra=("设置…", lambda: self._on_edit_mapping(self.report_mapping)),
        )
        self.report_mapping.pack(fill=tk.X, pady=2)
        self.report_patterns = PathRow(tab, "文件类型", mode="save", browse=False)
        self.report_patterns.pack(fill=tk.X, pady=2)

        self.report_recursive = tk.BooleanVar(value=True)
        ttk.Checkbutton(tab, text="包含子文件夹", variable=self.report_recursive).pack(anchor="w", pady=(6, 0))

        self._run_button(tab, "开始汇总", self._on_run_report)

    def _build_convert_tab(self) -> None:
        tab = self._new_tab(TAB_CONVERT)
        self.convert_inputs = InputPicker(tab)
        self.convert_inputs.pack(fill=tk.BOTH, expand=True)

        target_row = ttk.Frame(tab)
        target_row.pack(fill=tk.X, pady=(8, 2))
        ttk.Label(target_row, text="目标格式：", width=13, anchor="w").pack(side=tk.LEFT)
        self.convert_target = tk.StringVar(value=CONVERT_TARGETS[1])
        ttk.Combobox(
            target_row, textvariable=self.convert_target, values=CONVERT_TARGETS, width=12, state="readonly"
        ).pack(side=tk.LEFT)

        self.convert_output = PathRow(tab, "输出文件夹", mode="dir", title="选择输出文件夹")
        self.convert_output.pack(fill=tk.X, pady=2)
        self.convert_mapping = PathRow(
            tab,
            "字段映射",
            mode="open",
            title="选择映射配置",
            extra=("设置…", lambda: self._on_edit_mapping(self.convert_mapping)),
        )
        self.convert_mapping.pack(fill=tk.X, pady=2)
        self.convert_patterns = PathRow(tab, "文件类型", mode="save", browse=False)
        self.convert_patterns.pack(fill=tk.X, pady=2)

        self.convert_recursive = tk.BooleanVar(value=True)
        ttk.Checkbutton(tab, text="包含子文件夹", variable=self.convert_recursive).pack(anchor="w", pady=(6, 0))

        self._run_button(tab, "开始转换", self._on_run_convert)

    def _build_merge_tab(self) -> None:
        tab = self._new_tab(TAB_MERGE)
        self.merge_inputs = InputPicker(tab)
        self.merge_inputs.pack(fill=tk.BOTH, expand=True)

        mode_row = ttk.Frame(tab)
        mode_row.pack(fill=tk.X, pady=(8, 2))
        ttk.Label(mode_row, text="合并方式：", width=13, anchor="w").pack(side=tk.LEFT)
        self.merge_mode = tk.StringVar(value=MERGE_MODE_LABELS["auto"])
        ttk.Combobox(
            mode_row,
            textvariable=self.merge_mode,
            values=list(MERGE_MODE_VALUES),
            width=20,
            state="readonly",
        ).pack(side=tk.LEFT)

        self.merge_output = PathRow(tab, "输出文件", mode="save", title="保存合并结果")
        self.merge_output.pack(fill=tk.X, pady=2)
        self.merge_patterns = PathRow(tab, "文件类型", mode="save", browse=False)
        self.merge_patterns.pack(fill=tk.X, pady=2)

        self.merge_recursive = tk.BooleanVar(value=True)
        ttk.Checkbutton(tab, text="包含子文件夹", variable=self.merge_recursive).pack(anchor="w", pady=(6, 0))

        self._run_button(tab, "开始合并", self._on_run_merge)

    def _build_scan_tab(self) -> None:
        tab = self._new_tab(TAB_SCAN)
        self.scan_dir = PathRow(tab, "待扫描目录", mode="dir", title="选择待扫描目录")
        self.scan_dir.pack(fill=tk.X, pady=2)
        self.scan_output = PathRow(tab, "输出文件", mode="save", title="保存提取结果")
        self.scan_output.pack(fill=tk.X, pady=2)
        self.scan_mapping = PathRow(
            tab,
            "字段映射",
            mode="open",
            title="选择映射配置",
            extra=("设置…", lambda: self._on_edit_mapping(self.scan_mapping)),
        )
        self.scan_mapping.pack(fill=tk.X, pady=2)
        self.scan_patterns = PathRow(tab, "文件类型", mode="save", browse=False)
        self.scan_patterns.pack(fill=tk.X, pady=2)

        self.scan_recursive = tk.BooleanVar(value=True)
        ttk.Checkbutton(tab, text="包含子文件夹", variable=self.scan_recursive).pack(anchor="w", pady=(6, 0))

        self._run_button(tab, "开始提取", self._on_run_scan)

    def _build_classify_tab(self) -> None:
        tab = self._new_tab(TAB_CLASSIFY)
        self.classify_dir = PathRow(tab, "待整理目录", mode="dir", title="选择待整理目录")
        self.classify_dir.pack(fill=tk.X, pady=2)
        self.classify_config = PathRow(
            tab,
            "归档规则",
            mode="open",
            title="选择归档规则文件",
            extra=("设置…", self._on_edit_classify),
        )
        self.classify_config.pack(fill=tk.X, pady=2)
        self.classify_root = PathRow(tab, "归档根目录", mode="dir", title="选择归档根目录")
        self.classify_root.pack(fill=tk.X, pady=2)

        self._run_button(tab, "开始归档", self._on_run_classify)

    def _build_upload_tab(self) -> None:
        tab = self._new_tab(TAB_UPLOAD)
        self.upload_inputs = InputPicker(tab)
        self.upload_inputs.pack(fill=tk.BOTH, expand=True)

        url_row = ttk.Frame(tab)
        url_row.pack(fill=tk.X, pady=(8, 2))
        ttk.Label(url_row, text="接口地址：", width=13, anchor="w").pack(side=tk.LEFT)
        self.upload_url = tk.StringVar(value="")
        ttk.Entry(url_row, textvariable=self.upload_url).pack(side=tk.LEFT, fill=tk.X, expand=True)

        token_row = ttk.Frame(tab)
        token_row.pack(fill=tk.X, pady=2)
        ttk.Label(token_row, text="API 密钥：", width=13, anchor="w").pack(side=tk.LEFT)
        self.upload_token = tk.StringVar(value="")
        ttk.Entry(token_row, textvariable=self.upload_token, show="*").pack(side=tk.LEFT, fill=tk.X, expand=True)

        method_row = ttk.Frame(tab)
        method_row.pack(fill=tk.X, pady=2)
        ttk.Label(method_row, text="提交方式：", width=13, anchor="w").pack(side=tk.LEFT)
        self.upload_method = tk.StringVar(value=UPLOAD_METHOD_LABELS["multipart"])
        ttk.Combobox(
            method_row,
            textvariable=self.upload_method,
            values=list(UPLOAD_METHOD_VALUES),
            width=22,
            state="readonly",
        ).pack(side=tk.LEFT)

        actions = ttk.Frame(tab)
        actions.pack(anchor="w", pady=(10, 0))
        ping_button = ttk.Button(actions, text="Ping 测试", command=self._on_ping_upload)
        ping_button.pack(side=tk.LEFT)
        self._run_buttons.append(ping_button)
        upload_button = ttk.Button(actions, text="开始上传", style="Run.TButton", command=self._on_run_upload)
        upload_button.pack(side=tk.LEFT, padx=(8, 0))
        self._run_buttons.append(upload_button)

    def _build_manual_tab(self) -> None:
        tab = self._new_plain_tab(TAB_MANUAL)  # 展示页：内容直接占满可用区域
        actions = ttk.Frame(tab)
        actions.pack(fill=tk.X)
        ttk.Button(actions, text="打开说明书", command=self._on_open_manual).pack(side=tk.LEFT)
        ttk.Button(actions, text="重新载入", command=self._load_manual).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(actions, text="技术支持", command=self._on_show_support).pack(side=tk.LEFT, padx=(8, 0))

        holder = ttk.Frame(tab)
        holder.pack(fill=tk.BOTH, expand=True, pady=(8, 0))
        self.manual_text = tk.Text(holder, wrap=tk.WORD, relief=tk.FLAT, background="#ffffff",
                                   padx=10, pady=8)
        scrollbar = ttk.Scrollbar(holder, orient=tk.VERTICAL, command=self.manual_text.yview)
        self.manual_text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.manual_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self._configure_manual_tags()
        self._load_manual()

    def _configure_manual_tags(self) -> None:
        self.manual_text.tag_configure("h1", font=(self._family, 14, "bold"), spacing1=10, spacing3=6)
        self.manual_text.tag_configure("h2", font=(self._family, 12, "bold"), spacing1=10, spacing3=4)
        self.manual_text.tag_configure("h3", font=(self._family, 10, "bold"), spacing1=8, spacing3=2)
        self.manual_text.tag_configure("code", font=("Consolas", 9), lmargin1=24, lmargin2=24)

    def _load_manual(self) -> None:
        content, path = read_manual()
        widget = self.manual_text
        widget.configure(state=tk.NORMAL)
        widget.delete("1.0", tk.END)
        for line in content.splitlines():
            if line.startswith("### "):
                widget.insert(tk.END, f"{line[4:]}\n", "h3")
            elif line.startswith("## "):
                widget.insert(tk.END, f"{line[3:]}\n", "h2")
            elif line.startswith("# "):
                widget.insert(tk.END, f"{line[2:]}\n", "h1")
            elif line.strip().startswith("|"):
                widget.insert(tk.END, line.strip().strip("|").replace("|", "　") + "\n", "code")
            elif line.strip().startswith("```"):
                continue  # 代码块标记不显示
            else:
                widget.insert(tk.END, f"{line}\n")
        widget.configure(state=tk.DISABLED)
        self._manual_file = path
        if path is None:
            self._append_message("未找到 docs/USER_GUIDE.md，已显示内置简版说明。", kind="detail")

    def _on_open_manual(self) -> None:
        if not self._manual_file:
            self.show_message("需要先准备用户手册文档 docs/USER_GUIDE.md")
            return
        if open_document(self._manual_file):
            self.show_success(f"已用系统默认程序打开 {self._manual_file.name}")
        else:
            self.show_error(f"打开失败，请手动打开 {self._manual_file}")

    def _on_show_support(self) -> None:
        """弹出技术支持信息：联系方式与运行环境。"""
        dialog = SupportDialog(self, text=support_text())
        self.wait_window(dialog)
        if dialog.copied:
            self.show_success("技术支持信息已复制到剪贴板")

    def _build_formats_tab(self) -> None:
        tab = self._new_plain_tab(TAB_FORMATS)
        holder = ttk.Frame(tab)
        holder.pack(fill=tk.BOTH, expand=True)
        text = tk.Text(holder, wrap=tk.WORD, relief=tk.FLAT, background="#ffffff", padx=10, pady=8)
        scrollbar = ttk.Scrollbar(holder, orient=tk.VERTICAL, command=text.yview)
        text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        for suffix, note, can_input, can_output in support_matrix():
            marks = []
            if can_input == "是":
                marks.append("读")
            if can_output == "是":
                marks.append("写")
            text.insert(tk.END, f"{suffix:<7}{'/'.join(marks):<6}{note}\n")
        text.configure(state=tk.DISABLED)

    # -------------------------------------------------- 状态与消息
    def _schedule(self, delay_ms: int, callback: Callable[[], None]) -> str:
        """注册延时任务并记录，便于窗口销毁时统一取消。"""
        job = self.after(delay_ms, callback)
        self._after_jobs.append(job)
        return job

    def destroy(self) -> None:
        for job in list(self._after_jobs):
            try:
                self.after_cancel(job)
            except (tk.TclError, ValueError):  # pragma: no cover - 已执行或不存在的任务
                pass
        self._after_jobs.clear()
        if self._poll_job is not None:
            try:
                self.after_cancel(self._poll_job)
            except (tk.TclError, ValueError):  # pragma: no cover
                pass
            self._poll_job = None
        self._status_job = None
        super().destroy()

    def _announce_startup(self) -> None:
        """报告启动耗时；若已有更重要的状态（如错误）则不覆盖。"""
        if self.status.get() != STATUS_READY:
            return
        elapsed = time.perf_counter() - self._startup_started
        self.show_message(f"启动耗时 {elapsed:.2f} 秒")

    def show_ready(self) -> None:
        """回到「就绪」。"""
        self._cancel_status_job()
        self.status.set(STATUS_READY)
        self.status_bar.configure(style="Status.TLabel")

    def show_message(self, text: str) -> None:
        """普通提示（如"需要…"），数秒后自动回到「就绪」。"""
        self._set_status(text, kind="info")

    def show_success(self, text: str) -> None:
        """成功提示，数秒后自动回到「就绪」。"""
        self._set_status(f"成功 · {text}", kind="success")

    def show_error(self, text: str, *, fatal: bool = False) -> None:
        """失败提示。

        ``fatal=True`` 表示软件无法继续运行（例如界面初始化失败），此时提示会一直保留；
        其余错误仍会在数秒后回到「就绪」，完整原因见右栏消息。
        """
        self._set_status(f"失败 · {text}", kind="error", sticky=fatal)

    def _set_status(self, text: str, *, kind: str, sticky: bool = False) -> None:
        self._cancel_status_job()
        self.status.set(text)
        self.status_bar.configure(style=STATUS_STYLES.get(kind, "Status.TLabel"))
        self._append_message(text, kind=kind)
        if not sticky:
            self._status_job = self._schedule(self.status_reset_ms, self.show_ready)

    def _cancel_status_job(self) -> None:
        if self._status_job is not None:
            try:
                self.after_cancel(self._status_job)
            except (tk.TclError, ValueError):  # pragma: no cover - 窗口已销毁
                pass
            self._status_job = None

    def _append_message(self, text: str, *, kind: str = "info") -> None:
        """把提示或处理明细写入右栏消息区（带时间戳，按类型着色）。"""
        stamp = datetime.now().strftime("%H:%M:%S")
        prefix = "  " if kind == "detail" else ""
        widget = self.message_text
        widget.configure(state=tk.NORMAL)
        widget.insert(tk.END, f"[{stamp}] {prefix}{text}\n", kind)
        if int(widget.index("end-1c").split(".")[0]) > MESSAGE_MAX_LINES:
            widget.delete("1.0", f"{MESSAGE_TRIM_LINES}.0")  # 丢弃最早的记录
        widget.see(tk.END)
        widget.configure(state=tk.DISABLED)

    def _set_hint(self, text: str) -> None:
        """更新右栏的"当前功能"说明。"""
        self.hint_text.configure(state=tk.NORMAL)
        self.hint_text.delete("1.0", tk.END)
        self.hint_text.insert(tk.END, text)
        self.hint_text.configure(state=tk.DISABLED)

    def _on_tab_changed(self, _event: tk.Event | None = None) -> None:
        try:
            current = self.notebook.select()
        except tk.TclError:  # pragma: no cover - 窗口正在销毁
            return
        title = self.notebook.tab(current, "text") if current else TAB_REPORT
        self._set_hint(TAB_HINTS.get(title, ""))

    def _current_tab(self) -> str:
        try:
            current = self.notebook.select()
        except tk.TclError:  # pragma: no cover
            return ""
        return self.notebook.tab(current, "text") if current else ""

    # -------------------------------------------------- 事件处理
    def _load_rules(self, path: str, default: dict[str, Any]) -> dict[str, Any]:
        """读取规则文件；为空或读取失败时回退到默认规则。"""
        if not path:
            return copy.deepcopy(default)
        try:
            return rules.read_rules(path)
        except (OSError, ValueError, json.JSONDecodeError) as error:
            self.show_message(f"读取规则文件失败，改用默认规则：{error}")
            return copy.deepcopy(default)

    def _on_edit_mapping(self, row: PathRow) -> None:
        """在软件中编辑字段映射，结果写入临时 JSON。"""
        dialog = MappingDialog(self, self._load_rules(row.get(), rules.DEFAULT_MAPPING))
        self.wait_window(dialog)
        if dialog.result is None:
            return
        path = rules.write_temp_rules("mapping", dialog.result)
        row.set(str(path))
        self.show_success(f"字段映射已保存到临时文件 {path.name}")

    def _on_edit_classify(self) -> None:
        """在软件中编辑归档规则，结果写入临时 JSON。"""
        initial = self._load_rules(self.classify_config.get(), rules.DEFAULT_CLASSIFY)
        dialog = ClassifyDialog(self, initial)
        self.wait_window(dialog)
        if dialog.result is None:
            return
        path = rules.write_temp_rules("classify", dialog.result)
        self.classify_config.set(str(path))
        self.show_success(f"归档规则已保存到临时文件 {path.name}")

    def _on_run_report(self) -> None:
        inputs = self.report_inputs.paths()
        output = self.report_output.get()
        if not inputs:
            self.show_message("需要选择要汇总的报表文件或文件夹")
            return
        if not output:
            self.show_message("需要填写输出文件路径")
            return
        patterns = split_patterns(self.report_patterns.get())
        mapping = self.report_mapping.get()
        recursive = self.report_recursive.get()

        self.run_task(
            lambda log: run_report(
                inputs, output, mapping_path=mapping, patterns=patterns, recursive=recursive, log=log
            )
        )

    def _on_run_convert(self) -> None:
        inputs = self.convert_inputs.paths()
        output = self.convert_output.get()
        if not inputs:
            self.show_message("需要选择要转换的文件或文件夹")
            return
        if not output:
            self.show_message("需要选择输出文件夹")
            return
        target = self.convert_target.get()
        patterns = split_patterns(self.convert_patterns.get())
        mapping = self.convert_mapping.get()
        recursive = self.convert_recursive.get()

        self.run_task(
            lambda log: run_convert(
                inputs,
                output,
                target,
                mapping_path=mapping,
                patterns=patterns,
                recursive=recursive,
                log=log,
            )
        )

    def _on_run_merge(self) -> None:
        inputs = self.merge_inputs.paths()
        output = self.merge_output.get()
        if not inputs:
            self.show_message("需要选择要合并的文档")
            return
        if not output:
            self.show_message("需要填写输出文件路径")
            return
        mode = MERGE_MODE_VALUES.get(self.merge_mode.get(), "auto")
        patterns = split_patterns(self.merge_patterns.get())
        recursive = self.merge_recursive.get()

        self.run_task(
            lambda log: run_merge(
                inputs, output, mode=mode, patterns=patterns, recursive=recursive, log=log
            )
        )

    def _on_run_scan(self) -> None:
        directory = self.scan_dir.get()
        output = self.scan_output.get()
        if not directory:
            self.show_message("需要选择待扫描目录")
            return
        if not output:
            self.show_message("需要填写输出文件路径")
            return
        patterns = split_patterns(self.scan_patterns.get())
        mapping = self.scan_mapping.get()
        recursive = self.scan_recursive.get()

        self.run_task(
            lambda log: run_scan(
                directory, output, mapping_path=mapping, patterns=patterns, recursive=recursive, log=log
            )
        )

    def _on_run_classify(self) -> None:
        directory = self.classify_dir.get()
        config = self.classify_config.get()
        if not directory:
            self.show_message("需要选择待整理目录")
            return
        if not config:
            self.show_message("需要选择归档规则文件")
            return
        root = self.classify_root.get()

        self.run_task(lambda log: run_classify(directory, config, root=root, log=log))

    def _on_run_upload(self) -> None:
        paths = self.upload_inputs.paths()
        url = self.upload_url.get().strip()
        if not paths:
            self.show_message("需要选择要上传的文件或文件夹")
            return
        if not url:
            self.show_message("需要填写接口地址")
            return
        token = self.upload_token.get().strip()
        method = UPLOAD_METHOD_VALUES.get(self.upload_method.get(), "multipart")

        self.run_task(lambda log: run_upload(paths, url=url, token=token, method=method, log=log))

    def _on_ping_upload(self) -> None:
        url = self.upload_url.get().strip()
        if not url:
            self.show_message("需要填写接口地址")
            return
        token = self.upload_token.get().strip()
        method = UPLOAD_METHOD_VALUES.get(self.upload_method.get(), "multipart")

        self.run_task(lambda log: ping_upload(url, token, method=method))

    # -------------------------------------------------- 任务调度
    def run_task(self, task: Callable[[LogFunc], str], *, background: bool = True) -> None:
        """执行任务并把结果写入状态栏与右栏消息。

        默认在后台线程运行（界面不阻塞）；``background=False`` 时同步执行，
        便于测试在无并发依赖的情况下验证"结果 → 状态栏"的完整链路。
        """
        if self._busy:
            self.show_message("当前任务正在运行，请稍候")
            return

        self._busy = True
        self._cancel_status_job()
        self.status.set("正在运行…")
        self.status_bar.configure(style="Status.TLabel")
        for button in self._run_buttons:
            button.state(["disabled"])
        self._append_message("─" * 20, kind="detail")

        if background:
            threading.Thread(target=self._worker, args=(task,), daemon=True).start()
        else:
            self._worker(task)
            self._drain_queue()

    def _worker(self, task: Callable[[LogFunc], str]) -> None:
        try:
            summary = task(lambda text: self._queue.put(("log", text)))
            self._queue.put(("done", summary))
        except MemoryError as error:  # 无法继续运行，提示保持显示
            self._queue.put(("fatal", f"{type(error).__name__}: {error}"))
        except Exception as error:  # noqa: BLE001 - 其余错误展示后自动复位
            self._queue.put(("error", f"{type(error).__name__}: {error}"))
        finally:
            self._queue.put(("finish", ""))

    def _poll_queue(self) -> None:
        """常驻轮询：周期性把后台线程的消息搬到界面线程，并重新排期。"""
        self._drain_queue()
        self._poll_job = self.after(POLL_INTERVAL_MS, self._poll_queue)

    def _drain_queue(self) -> None:
        finished = False
        while True:
            try:
                kind, payload = self._queue.get_nowait()
            except queue.Empty:
                break

            if kind == "log":
                self._append_message(payload, kind="detail")
            elif kind == "done":
                if payload:
                    self.show_success(payload)
            elif kind == "error":
                self.show_error(payload)
            elif kind == "fatal":
                self.show_error(payload, fatal=True)
            elif kind == "finish":
                finished = True

        if finished:
            self._busy = False
            for button in self._run_buttons:
                button.state(["!disabled"])

    # -------------------------------------------------- 小工具


def main(argv: Sequence[str] | None = None) -> int:
    """启动图形界面。"""
    arguments = list(sys.argv[1:] if argv is None else argv)
    started = time.perf_counter()
    try:
        app = OfficeAssistantApp(startup_started=started)
    except Exception as error:  # 界面都起不来时只能写日志并抛出
        print(f"启动失败：{type(error).__name__}: {error}", file=sys.stderr)
        return 1
    if arguments and Path(arguments[0]).exists():
        app.report_inputs.set_paths([arguments[0]])  # 支持把文件/目录带到输入列表
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
