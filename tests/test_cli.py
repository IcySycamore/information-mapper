"""命令行端到端测试。"""

from __future__ import annotations

import json
from pathlib import Path

import pandas as pd
import pytest

from information_mapper.cli import main


@pytest.fixture
def mapping_file(tmp_path: Path) -> Path:
    path = tmp_path / "mapping.json"
    path.write_text(
        json.dumps(
            {
                "target_headers": ["姓名", "联系电话", "家庭住址"],
                "field_mapping": {
                    "姓名": ["姓名", "户主姓名"],
                    "联系电话": ["联系电话", "电话"],
                    "家庭住址": ["家庭住址", "地址"],
                },
            }
        ),
        encoding="utf-8",
    )
    return path


@pytest.fixture
def classify_config_file(tmp_path: Path) -> Path:
    path = tmp_path / "classify.json"
    path.write_text(
        json.dumps(
            {
                "target_root": "归档",
                "default_target": "99-其他",
                "rules": [
                    {"name": "报表", "target": "01-报表", "extensions": [".xlsx", ".csv"]},
                ],
            }
        ),
        encoding="utf-8",
    )
    return path


def test_cli_without_args_prints_help(capsys: pytest.CaptureFixture[str]) -> None:
    assert main([]) == 0
    assert "子命令" in capsys.readouterr().out


def test_cli_help_uses_main_parser(capsys: pytest.CaptureFixture[str]) -> None:
    with pytest.raises(SystemExit) as excinfo:
        main(["--help"])
    assert excinfo.value.code == 0
    output = capsys.readouterr().out
    assert "子命令" in output
    assert "merge" in output


def test_cli_version(capsys: pytest.CaptureFixture[str]) -> None:
    with pytest.raises(SystemExit) as excinfo:
        main(["--version"])
    assert excinfo.value.code == 0
    assert "information-mapper" in capsys.readouterr().out


def test_cli_formats_lists_supported_types(capsys: pytest.CaptureFixture[str]) -> None:
    assert main(["formats"]) == 0
    output = capsys.readouterr().out
    assert ".xlsx" in output and ".docx" in output and ".pdf" in output


def test_cli_unknown_subcommand_exits() -> None:
    with pytest.raises(SystemExit):
        main(["不存在的命令"])


def test_cli_legacy_arguments_still_work(
    tmp_path: Path, xlsx_file: Path, mapping_file: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    target = tmp_path / "映射结果.xlsx"
    code = main(["--input", str(xlsx_file), "--output", str(target), "--mapping", str(mapping_file)])
    assert code == 0
    assert "处理完成" in capsys.readouterr().out
    frame = pd.read_excel(target)
    assert list(frame.columns) == ["姓名", "联系电话", "家庭住址"]


def test_cli_map_subcommand(
    tmp_path: Path, xlsx_file: Path, mapping_file: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    target = tmp_path / "映射结果.csv"
    code = main(["map", "-i", str(xlsx_file), "-o", str(target), "--mapping", str(mapping_file)])
    assert code == 0
    frame = pd.read_csv(target, encoding="utf-8-sig")
    assert list(frame.columns) == ["姓名", "联系电话", "家庭住址"]


def test_cli_convert_single_file(tmp_path: Path, xlsx_file: Path, capsys: pytest.CaptureFixture[str]) -> None:
    target = tmp_path / "转换.csv"
    assert main(["convert", "-i", str(xlsx_file), "-o", str(target)]) == 0
    assert "成功" in capsys.readouterr().out
    assert target.exists()


def test_cli_convert_batch(
    tmp_path: Path, inbox: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    outcome = tmp_path / "csv"
    code = main(["convert", "-i", str(inbox), "-o", str(outcome), "--to", "csv"])
    assert code == 1  # 含一份损坏文件，整体仍完成
    assert (outcome / "报表A.csv").exists()
    assert "转换完成" in capsys.readouterr().out


def test_cli_merge_into_single_file(
    tmp_path: Path, inbox: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    target = tmp_path / "汇总.xlsx"
    assert main(["merge", "-i", str(inbox), "-o", str(target)]) == 0
    output = capsys.readouterr().out
    assert "已合并" in output
    frame = pd.read_excel(target)
    assert "源文件" in frame.columns


def test_cli_merge_text_mode(tmp_path: Path, plain_text_file: Path, text_file: Path) -> None:
    target = tmp_path / "合集.md"
    assert main(["merge", "-i", str(plain_text_file), "-i", str(text_file), "-o", str(target), "--mode", "text"]) == 0
    assert "会议室" in target.read_text(encoding="utf-8")


def test_cli_report(tmp_path: Path, inbox: Path, capsys: pytest.CaptureFixture[str]) -> None:
    target = tmp_path / "报表汇总.xlsx"
    assert main(["report", "-i", str(inbox), "-o", str(target), "--pattern", "*.xlsx"]) == 0
    assert "已合并" in capsys.readouterr().out
    assert target.exists()


def test_cli_scan(tmp_path: Path, inbox: Path, capsys: pytest.CaptureFixture[str]) -> None:
    target = tmp_path / "扫描结果.xlsx"
    assert main(["scan", "-i", str(inbox), "-o", str(target)]) == 0
    output = capsys.readouterr().out
    assert "扫描完成" in output
    assert "已跳过" in output  # 损坏文件被跳过并提示


def test_cli_classify_dry_run(
    inbox: Path, classify_config_file: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    assert main(["classify", "-i", str(inbox), "--config", str(classify_config_file)]) == 0
    output = capsys.readouterr().out
    assert "归档计划" in output
    assert (inbox / "报表A.xlsx").exists()  # 预演不移动


def test_cli_classify_apply(
    tmp_path: Path, inbox: Path, classify_config_file: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    assert main(["classify", "-i", str(inbox), "--config", str(classify_config_file), "--apply"]) == 0
    assert "已归档" in capsys.readouterr().out
    assert (inbox / "归档" / "01-报表" / "报表A.xlsx").exists()


def test_cli_upload_dry_run(xlsx_file: Path, capsys: pytest.CaptureFixture[str]) -> None:
    assert main(["upload", "-i", str(xlsx_file)]) == 0
    output = capsys.readouterr().out
    assert "预演" in output
    assert "占位符" in output or "your-server.example.com" in output


def test_cli_upload_apply_without_real_config_fails(xlsx_file: Path, capsys: pytest.CaptureFixture[str]) -> None:
    code = main(["upload", "-i", str(xlsx_file), "--apply"])
    assert code == 1
    assert "占位符" in capsys.readouterr().out
