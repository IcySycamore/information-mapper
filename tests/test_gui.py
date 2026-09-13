"""图形界面测试：布局约定、状态栏语义、服务层函数。"""

from __future__ import annotations

import json
import time
import urllib.error
from pathlib import Path

import pandas as pd
import pytest

from information_mapper import gui, rules
from information_mapper.gui import (
    SIDE_PANEL_WIDTH,
    STATUS_READY,
    TAB_HINTS,
    TAB_ORDER,
    manual_path,
    ping_upload,
    read_manual,
    run_classify,
    run_convert,
    run_merge,
    run_report,
    run_scan,
    run_upload,
    split_patterns,
    supported_filetypes,
)

EXPECTED_TABS = [
    "报表汇总",
    "格式转换",
    "合并文档",
    "批量提取",
    "文件归档",
    "自动上传",
    "使用说明",
    "支持的格式",
]

MANUAL_TAB_INDEX = EXPECTED_TABS.index("使用说明")

QUIET = lambda _: None  # noqa: E731


def _pump(app, seconds: float) -> None:
    """驱动事件循环若干秒（tkinter 未进入 mainloop 时用）。"""
    deadline = time.time() + seconds
    while time.time() < deadline:
        app.update()
        time.sleep(0.01)


def _make_app():
    tk = pytest.importorskip("tkinter")
    app = None
    last_error: Exception | None = None
    for _ in range(3):  # 频繁创建/销毁窗口时偶发 TclError，重试即可
        try:
            app = gui.OfficeAssistantApp()
            break
        except tk.TclError as error:
            last_error = error
            time.sleep(0.15)
    if app is None:  # pragma: no cover - 无图形环境
        pytest.skip(f"当前环境无法创建窗口：{last_error}")
    app.geometry("1040x840+4000+4000")  # 移到可视区域之外，既能真实布局又不打扰使用者
    app.update()
    _pump(app, 0.25)  # 等启动回调完成
    return app


# ---------------------------------------------------------------- 代码约定
def test_tabs_and_hints_are_consistent() -> None:
    assert list(TAB_HINTS) == list(TAB_ORDER)
    assert list(TAB_ORDER) == EXPECTED_TABS


def test_gui_no_longer_uses_message_boxes() -> None:
    assert not hasattr(gui, "messagebox")


def test_gui_source_has_no_dry_run_wording() -> None:
    source = Path(gui.__file__).read_text(encoding="utf-8")
    for word in ("预演", "真正上传", "真正执行", "askyesno", "showerror", "showwarning"):
        assert word not in source, f"界面中不应再出现：{word}"


def test_gui_exposes_ping_test() -> None:
    source = Path(gui.__file__).read_text(encoding="utf-8")
    assert "Ping 测试" in source


def test_inputs_are_not_prefilled() -> None:
    source = Path(gui.__file__).read_text(encoding="utf-8")
    assert "DEFAULT_UPLOAD_URL" not in source
    assert "DEFAULT_API_TOKEN" not in source
    assert 'default="output/' not in source


def test_manual_document_is_available() -> None:
    path = manual_path()
    assert path is not None, "未找到 docs/USER_GUIDE.md"
    content, resolved = read_manual()
    assert resolved == path
    for section in ("软件概述", "界面说明", "操作流程", "常见问题处理"):
        assert section in content


# ---------------------------------------------------------------- 辅助函数
def test_split_patterns() -> None:
    assert split_patterns("*.xlsx, *.csv") == ["*.xlsx", "*.csv"]
    assert split_patterns("*.docx；*.pdf") == ["*.docx", "*.pdf"]
    assert split_patterns("") is None


def test_supported_filetypes_covers_inputs() -> None:
    patterns = supported_filetypes()[0][1]
    assert "*.xlsx" in patterns and "*.docx" in patterns and "*.pdf" in patterns


def test_dropdown_labels_are_reversible() -> None:
    assert gui.MERGE_MODE_VALUES["汇总成一张表"] == "table"
    assert gui.UPLOAD_METHOD_VALUES["json + Base64"] == "json"


# ---------------------------------------------------------------- 服务层
def test_run_report_merges_into_one_file(
    xlsx_file: Path, xlsx_file_b: Path, tmp_path: Path
) -> None:
    logs: list[str] = []
    output = tmp_path / "汇总.xlsx"
    summary = run_report([xlsx_file, xlsx_file_b], output, log=logs.append)

    assert output.exists()
    assert "汇总完成" in summary
    frame = pd.read_excel(output)
    assert "源文件" in frame.columns
    assert len(frame) == 3
    assert any("找到 2 份报表" in line for line in logs)


def test_run_report_without_files_raises(tmp_path: Path) -> None:
    empty = tmp_path / "空目录"
    empty.mkdir()
    with pytest.raises(ValueError, match="没有找到可汇总"):
        run_report([empty], tmp_path / "汇总.xlsx", log=QUIET)


def test_run_report_applies_mapping(xlsx_file: Path, tmp_path: Path) -> None:
    mapping = tmp_path / "mapping.json"
    mapping.write_text(json.dumps({"target_headers": ["联系电话", "姓名"]}), encoding="utf-8")
    output = tmp_path / "汇总.xlsx"
    run_report([xlsx_file], output, mapping_path=mapping, log=QUIET)
    assert list(pd.read_excel(output).columns)[:2] == ["联系电话", "姓名"]


def test_run_convert_batch(xlsx_file: Path, tmp_path: Path) -> None:
    summary = run_convert([xlsx_file], tmp_path / "out", "csv", log=QUIET)
    assert "成功 1" in summary
    assert (tmp_path / "out" / "报表A.csv").exists()


def test_run_convert_keeps_going_on_broken_file(inbox: Path, tmp_path: Path) -> None:
    logs: list[str] = []
    summary = run_convert([inbox], tmp_path / "out", "csv", log=logs.append)
    assert "成功" in summary
    assert any("失败" in line for line in logs)


def test_run_merge_word_mode(docx_file: Path, tmp_path: Path) -> None:
    output = tmp_path / "合集.docx"
    summary = run_merge([docx_file], output, mode="word", log=QUIET)
    assert output.exists()
    assert "合并完成" in summary


def test_run_scan_extracts_records(inbox: Path, tmp_path: Path) -> None:
    logs: list[str] = []
    output = tmp_path / "扫描.xlsx"
    summary = run_scan(inbox, output, log=logs.append)
    assert output.exists()
    assert "提取完成" in summary
    assert any("跳过" in line for line in logs)


def _classify_config(tmp_path: Path) -> Path:
    path = tmp_path / "classify.json"
    path.write_text(
        json.dumps(
            {
                "target_root": "归档",
                "default_target": "99-其他",
                "rules": [{"name": "报表", "target": "01-报表", "extensions": [".xlsx"]}],
            }
        ),
        encoding="utf-8",
    )
    return path


def test_run_classify_moves_files_directly(tmp_path: Path) -> None:
    inbox = tmp_path / "inbox"
    inbox.mkdir()
    (inbox / "报表.xlsx").write_bytes(b"x")

    summary = run_classify(inbox, _classify_config(tmp_path), log=QUIET)

    assert (inbox / "归档" / "01-报表" / "报表.xlsx").exists()
    assert not (inbox / "报表.xlsx").exists()
    assert "归档完成" in summary


def test_run_classify_without_files(tmp_path: Path) -> None:
    inbox = tmp_path / "inbox"
    inbox.mkdir()
    assert "没有需要归档" in run_classify(inbox, _classify_config(tmp_path), log=QUIET)


def test_run_upload_rejects_empty_url(xlsx_file: Path) -> None:
    with pytest.raises(ValueError, match="占位符"):
        run_upload([xlsx_file], url="", token="", log=QUIET)


def test_run_upload_without_files_raises(tmp_path: Path) -> None:
    empty = tmp_path / "空目录"
    empty.mkdir()
    with pytest.raises(ValueError, match="没有找到可上传"):
        run_upload([empty], url="https://upload.gov.cn/api", token="t", log=QUIET)


def test_run_upload_sends_files(monkeypatch: pytest.MonkeyPatch, xlsx_file: Path) -> None:
    class FakeResponse:
        status = 200

        def read(self, size: int = -1) -> bytes:
            return b'{"ok": true}'

        def __enter__(self) -> "FakeResponse":
            return self

        def __exit__(self, *args: object) -> bool:
            return False

    monkeypatch.setattr(
        "information_mapper.uploader.urllib.request.urlopen", lambda *a, **k: FakeResponse()
    )
    summary = run_upload(
        [xlsx_file], url="https://upload.gov.cn/api", token="real-token", log=QUIET
    )
    assert "上传完成：成功 1/1" in summary


# ---------------------------------------------------------------- Ping
def test_ping_reports_missing_address() -> None:
    assert "Ping 失败" in ping_upload("")


def test_ping_reports_placeholder_address() -> None:
    assert "Ping 失败" in ping_upload("https://your-server.example.com/api/upload")


def test_ping_success(monkeypatch: pytest.MonkeyPatch) -> None:
    class FakeResponse:
        status = 204

        def __enter__(self) -> "FakeResponse":
            return self

        def __exit__(self, *args: object) -> bool:
            return False

    monkeypatch.setattr(
        "information_mapper.uploader.urllib.request.urlopen", lambda *a, **k: FakeResponse()
    )
    message = ping_upload("https://upload.gov.cn/api", "token")
    assert "Ping 成功" in message and "204" in message


def test_ping_treats_http_error_as_reachable(monkeypatch: pytest.MonkeyPatch) -> None:
    def fake_urlopen(request, timeout=None, context=None):  # noqa: ANN001
        raise urllib.error.HTTPError(request.full_url, 405, "Method Not Allowed", {}, None)

    monkeypatch.setattr("information_mapper.uploader.urllib.request.urlopen", fake_urlopen)
    assert "Ping 成功" in ping_upload("https://upload.gov.cn/api")


def test_ping_reports_unreachable(monkeypatch: pytest.MonkeyPatch) -> None:
    def fake_urlopen(request, timeout=None, context=None):  # noqa: ANN001
        raise urllib.error.URLError("getaddrinfo failed")

    monkeypatch.setattr("information_mapper.uploader.urllib.request.urlopen", fake_urlopen)
    assert "Ping 失败" in ping_upload("https://not-exist.example.org/api")


# ---------------------------------------------------------------- 布局与尺寸
def test_window_builds_all_tabs() -> None:
    app = _make_app()
    try:
        tabs = [app.notebook.tab(tab_id, "text") for tab_id in app.notebook.tabs()]
        assert tabs == EXPECTED_TABS
    finally:
        app.destroy()


def test_status_bar_has_fixed_height() -> None:
    app = _make_app()
    try:
        height = app.status_bar.winfo_height()
        assert 0 < height <= 40, f"状态栏应为固定单行高度，实际 {height}"

        app.show_message("这是一条很长的提示信息，用于确认状态栏高度不会因为文字变长而改变。" * 3)
        app.update()
        assert app.status_bar.winfo_height() == height
    finally:
        app.destroy()


def test_side_panel_has_fixed_width() -> None:
    app = _make_app()
    try:
        width = app.side_panel.winfo_width()
        # 列宽固定为 SIDE_PANEL_WIDTH，扣除左右外边距（6 + 12）即控件实际宽度
        assert SIDE_PANEL_WIDTH - 20 <= width <= SIDE_PANEL_WIDTH

        app.geometry("1280x840+4000+4000")  # 拉宽主窗口，右栏宽度不应变化
        app.update()
        assert app.side_panel.winfo_width() == width
    finally:
        app.destroy()


def test_side_panel_shows_hint_of_current_tab() -> None:
    app = _make_app()
    try:
        for index, title in enumerate(EXPECTED_TABS):
            app.notebook.select(index)
            app.update()
            assert app.hint_text.get("1.0", "end").strip() == TAB_HINTS[title]
    finally:
        app.destroy()


def test_manual_area_is_large_enough() -> None:
    """使用说明页的文档区域应占满可用空间（此前过小、几乎看不见）。"""
    app = _make_app()
    try:
        app.notebook.select(MANUAL_TAB_INDEX)
        app.update()
        assert app.manual_text.winfo_height() > 200
        assert "软件特点" in app.manual_text.get("1.0", "end")
    finally:
        app.destroy()


# ---------------------------------------------------------------- 状态栏语义
def test_status_starts_with_ready() -> None:
    app = _make_app()
    try:
        app.show_ready()
        assert app.status.get() == STATUS_READY
    finally:
        app.destroy()


def test_missing_input_asks_in_status_bar() -> None:
    app = _make_app()
    try:
        app.report_inputs.set_paths([])
        app._on_run_report()
        assert app.status.get().startswith("需要")

        app.report_inputs.set_paths(["某个文件.xlsx"])
        app.report_output.set("")
        app._on_run_report()
        assert "需要填写输出文件" in app.status.get()

        app.upload_inputs.set_paths(["某个文件.xlsx"])
        app.upload_url.set("")
        app._on_run_upload()
        assert "需要填写接口地址" in app.status.get()
    finally:
        app.destroy()


def test_success_message_resets_to_ready() -> None:
    """成功提示应在数秒后自动回到「就绪」。"""
    app = _make_app()
    try:
        app.status_reset_ms = 400
        app.show_success("汇总完成：3 条记录")
        assert app.status.get() == "成功 · 汇总完成：3 条记录"
        _pump(app, 1.2)
        assert app.status.get() == STATUS_READY
    finally:
        app.destroy()


def test_message_and_error_also_reset() -> None:
    """需要输入与失败提示同样会自动复位。"""
    app = _make_app()
    try:
        app.status_reset_ms = 400

        app.show_message("需要选择要汇总的报表文件")
        assert app.status.get() == "需要选择要汇总的报表文件"
        _pump(app, 1.2)
        assert app.status.get() == STATUS_READY

        app.show_error("ValueError: 示例错误")
        assert app.status.get() == "失败 · ValueError: 示例错误"
        _pump(app, 1.2)
        assert app.status.get() == STATUS_READY
    finally:
        app.destroy()


def test_task_result_reaches_status_bar(tmp_path: Path, xlsx_file: Path) -> None:
    """任务结果应写入状态栏，并保留在右栏消息中。

    使用同步执行路径，避免依赖后台线程调度与 tkinter 定时器，
    从而不受测试执行顺序影响。异步路径由界面实际运行覆盖。
    """
    app = _make_app()
    try:
        output = tmp_path / "gui_汇总.xlsx"
        app.run_task(
            lambda log: gui.run_report([str(xlsx_file)], output, log=log),
            background=False,
        )

        assert output.exists()
        assert app.status.get().startswith("成功 · ")
        assert "成功" in app.message_text.get("1.0", "end")
        assert app._busy is False  # 执行结束后恢复可操作状态
    finally:
        app.destroy()


def test_fatal_error_stays_visible() -> None:
    app = _make_app()
    try:
        app.status_reset_ms = 400
        app.show_error("程序无法继续运行", fatal=True)
        _pump(app, 1.0)
        assert app.status.get() == "失败 · 程序无法继续运行"
    finally:
        app.destroy()


def test_task_details_also_go_to_message_panel(tmp_path: Path, xlsx_file: Path) -> None:
    """处理明细与提示统一进右栏消息区，不再保留独立的日志面板。"""
    app = _make_app()
    try:
        assert not hasattr(app, "log_text")
        output = tmp_path / "gui_汇总.xlsx"
        app.run_task(
            lambda log: gui.run_report([str(xlsx_file)], output, log=log),
            background=False,
        )
        content = app.message_text.get("1.0", "end")
        assert "找到 1 份报表" in content  # 过程明细
        assert "成功 · " in content  # 结果提示
    finally:
        app.destroy()


def test_messages_are_listed_in_side_panel() -> None:
    app = _make_app()
    try:
        app.show_message("需要选择要汇总的报表文件")
        app.show_success("汇总完成：3 条记录")
        content = app.message_text.get("1.0", "end")
        assert "需要选择要汇总的报表文件" in content
        assert "成功 · 汇总完成：3 条记录" in content
    finally:
        app.destroy()


def test_startup_time_is_reported() -> None:
    app = _make_app()
    try:
        deadline = time.time() + 5
        while time.time() < deadline and "启动耗时" not in app.message_text.get("1.0", "end"):
            app.update()
            time.sleep(0.02)
        assert "启动耗时" in app.message_text.get("1.0", "end")
    finally:
        app.destroy()


# ---------------------------------------------------------------- 技术支持
def test_support_text_contains_contacts() -> None:
    text = gui.support_text()
    assert "技术支持联系方式" in text
    assert "QQ：1284742412" in text
    assert "GitHub：IcySycamore" in text
    assert "邮箱：1284742412@qq.com" in text
    assert gui.SUPPORT_REPOSITORY in text


def test_environment_summary_reports_versions() -> None:
    summary = gui.environment_summary()
    assert summary["版本"] == gui.__version__
    assert summary["Python"]
    assert summary["依赖"]


def test_support_dialog_shows_contacts() -> None:
    app = _make_app()
    dialog = None
    try:
        dialog = gui.SupportDialog(app, text=gui.support_text())
        content = dialog.content()
        assert "QQ：1284742412" in content
        assert "GitHub：IcySycamore" in content
        assert "邮箱：1284742412@qq.com" in content
    finally:
        if dialog is not None:
            dialog.destroy()
        app.destroy()


def test_manual_tab_has_support_button() -> None:
    import tkinter.ttk as ttk

    app = _make_app()

    def collect(widget) -> list[str]:  # noqa: ANN001
        found: list[str] = []
        for child in widget.winfo_children():
            if isinstance(child, ttk.Button):
                found.append(str(child.cget("text")))
            found.extend(collect(child))
        return found

    try:
        app.notebook.select(MANUAL_TAB_INDEX)
        app.update()
        assert "技术支持" in collect(app.notebook)
    finally:
        app.destroy()


# ---------------------------------------------------------------- 规则编辑
def test_mapping_dialog_collects_rows() -> None:
    app = _make_app()
    dialog = None
    try:
        dialog = gui.MappingDialog(app, rules.DEFAULT_MAPPING)
        assert len(dialog.tree.get_children()) == len(rules.DEFAULT_MAPPING["target_headers"])

        dialog._rows.append(["测试字段", "来源A, 来源B"])
        result = dialog.collect()

        assert result["target_headers"][-1] == "测试字段"
        assert result["field_mapping"]["测试字段"] == ["来源A", "来源B"]
    finally:
        if dialog is not None:
            dialog.destroy()
        app.destroy()


def test_classify_dialog_collects_rules() -> None:
    app = _make_app()
    dialog = None
    try:
        dialog = gui.ClassifyDialog(app, rules.DEFAULT_CLASSIFY)
        assert len(dialog.tree.get_children()) == len(rules.DEFAULT_CLASSIFY["rules"])

        dialog._rows.append(["图片", "05-图片", "png, jpg", ""])
        dialog._copy.set(True)
        result = dialog.collect()

        assert result["copy"] is True
        assert result["rules"][-1]["extensions"] == [".png", ".jpg"]
    finally:
        if dialog is not None:
            dialog.destroy()
        app.destroy()


def test_edit_mapping_writes_temp_rules(monkeypatch: pytest.MonkeyPatch) -> None:
    """在软件中编辑字段映射后，应生成临时 JSON 并回填路径。"""
    app = _make_app()
    created: list[Path] = []
    try:
        class FakeMappingDialog:
            def __init__(self, master, config):  # noqa: ANN001
                assert config["target_headers"]
                self.result = rules.build_mapping([("姓名", "姓名, 户主姓名")])

        monkeypatch.setattr(gui, "MappingDialog", FakeMappingDialog)
        monkeypatch.setattr(app, "wait_window", lambda *args, **kwargs: None)

        app.report_mapping.set("")
        app._on_edit_mapping(app.report_mapping)

        path = Path(app.report_mapping.get())
        created.append(path)
        assert path.parent == rules.temp_rules_dir()
        data = json.loads(path.read_text(encoding="utf-8"))
        assert data["target_headers"] == ["姓名"]
        assert data["field_mapping"]["姓名"] == ["姓名", "户主姓名"]
    finally:
        for item in created:
            item.unlink(missing_ok=True)
        app.destroy()


def test_edit_classify_writes_temp_rules(monkeypatch: pytest.MonkeyPatch) -> None:
    """在软件中编辑归档规则后，应生成临时 JSON 并回填路径。"""
    app = _make_app()
    created: list[Path] = []
    try:
        class FakeClassifyDialog:
            def __init__(self, master, config):  # noqa: ANN001
                assert "rules" in config
                self.result = rules.build_classify([("报表", "01-报表", "xlsx", "")])

        monkeypatch.setattr(gui, "ClassifyDialog", FakeClassifyDialog)
        monkeypatch.setattr(app, "wait_window", lambda *args, **kwargs: None)

        app.classify_config.set("")
        app._on_edit_classify()

        path = Path(app.classify_config.get())
        created.append(path)
        data = json.loads(path.read_text(encoding="utf-8"))
        assert data["rules"][0]["target"] == "01-报表"
    finally:
        for item in created:
            item.unlink(missing_ok=True)
        app.destroy()


def test_load_rules_falls_back_to_default(tmp_path: Path) -> None:
    app = _make_app()
    try:
        assert app._load_rules("", rules.DEFAULT_MAPPING) == rules.DEFAULT_MAPPING

        broken = tmp_path / "broken.json"
        broken.write_text("{ 不是合法 JSON", encoding="utf-8")
        assert app._load_rules(str(broken), rules.DEFAULT_MAPPING) == rules.DEFAULT_MAPPING
    finally:
        app.destroy()


def test_upload_tab_defaults_are_empty() -> None:
    app = _make_app()
    try:
        assert app.upload_url.get() == ""
        assert app.upload_token.get() == ""
    finally:
        app.destroy()
