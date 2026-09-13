"""规则配置模块测试：结构互转与临时文件生成。"""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from information_mapper import rules


def test_split_list_accepts_multiple_separators() -> None:
    assert rules.split_list("a, b；c　d") == ["a", "b", "c", "d"]
    assert rules.split_list("") == []
    assert rules.split_list("  ") == []


def test_join_list_round_trip() -> None:
    assert rules.split_list(rules.join_list(["姓名", "户主姓名"])) == ["姓名", "户主姓名"]


def test_normalize_extension() -> None:
    assert rules.normalize_extension("xlsx") == ".xlsx"
    assert rules.normalize_extension(".XLSX") == ".xlsx"


# ---------------------------------------------------------------- 字段映射
def test_mapping_rows_from_config() -> None:
    rows = rules.mapping_rows(rules.DEFAULT_MAPPING)
    assert rows[0] == ("姓名", "姓名, 户主姓名, 人员姓名")
    assert len(rows) == len(rules.DEFAULT_MAPPING["target_headers"])


def test_mapping_rows_keeps_undeclared_fields() -> None:
    config = {"target_headers": ["姓名"], "field_mapping": {"姓名": ["姓名"], "电话": ["电话"]}}
    rows = rules.mapping_rows(config)
    assert ("电话", "电话") in rows


def test_build_mapping_falls_back_to_target_name() -> None:
    result = rules.build_mapping([("联系电话", ""), ("", "忽略")])
    assert result["target_headers"] == ["联系电话"]
    assert result["field_mapping"] == {"联系电话": ["联系电话"]}


def test_mapping_round_trip_is_stable() -> None:
    rows = rules.mapping_rows(rules.DEFAULT_MAPPING)
    rebuilt = rules.build_mapping(rows)
    assert rebuilt == rules.DEFAULT_MAPPING
    assert rules.mapping_rows(rebuilt) == rows


# ---------------------------------------------------------------- 归档规则
def test_classify_rows_from_config() -> None:
    rows = rules.classify_rows(rules.DEFAULT_CLASSIFY)
    assert rows[0] == ("报表", "01-报表", ".xlsx, .xls, .csv, .tsv", "")


def test_build_classify_normalizes_extensions() -> None:
    result = rules.build_classify(
        [("报表", "01-报表", "xlsx, CSV", "汇总, 报表")],
        target_root="已整理",
        default_target="其他",
        copy=True,
    )
    assert result["target_root"] == "已整理"
    assert result["copy"] is True
    assert result["rules"] == [
        {
            "name": "报表",
            "target": "01-报表",
            "extensions": [".xlsx", ".csv"],
            "keywords": ["汇总", "报表"],
        }
    ]


def test_build_classify_uses_name_when_target_missing() -> None:
    result = rules.build_classify([("图片", "", "png", "")])
    assert result["rules"][0]["target"] == "图片"
    assert result["rules"][0]["extensions"] == [".png"]
    assert "keywords" not in result["rules"][0]


def test_classify_round_trip_is_stable() -> None:
    rows = rules.classify_rows(rules.DEFAULT_CLASSIFY)
    rebuilt = rules.build_classify(
        rows,
        target_root=rules.DEFAULT_CLASSIFY["target_root"],
        default_target=rules.DEFAULT_CLASSIFY["default_target"],
    )
    assert rebuilt == rules.DEFAULT_CLASSIFY


def test_build_classify_skips_empty_rows() -> None:
    assert rules.build_classify([("", "", "", "")])["rules"] == []


# ---------------------------------------------------------------- 临时文件
def test_temp_rules_dir_is_created() -> None:
    directory = rules.temp_rules_dir()
    assert directory.is_dir()
    assert directory.name == rules.TEMP_SUBDIR


def test_write_and_read_temp_rules() -> None:
    path = rules.write_temp_rules("mapping-test", rules.DEFAULT_MAPPING)
    try:
        assert path.is_file()
        assert path.suffix == ".json"
        assert rules.read_rules(path) == rules.DEFAULT_MAPPING
        assert "target_headers" in json.loads(path.read_text(encoding="utf-8"))
    finally:
        path.unlink(missing_ok=True)


def test_read_rules_missing_file(tmp_path: Path) -> None:
    with pytest.raises(ValueError, match="找不到规则文件"):
        rules.read_rules(tmp_path / "不存在.json")
