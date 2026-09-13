"""文件自动归档测试。"""

from __future__ import annotations

import json
from pathlib import Path

from information_mapper.classify import (
    ClassifyConfig,
    ClassifyRule,
    organize,
    plan_moves,
    summarize,
)


def _config(*, copy: bool = False) -> ClassifyConfig:
    return ClassifyConfig.from_dict(
        {
            "target_root": "归档",
            "default_target": "99-其他",
            "copy": copy,
            "rules": [
                {"name": "报表", "target": "01-报表", "extensions": [".xlsx", ".csv"]},
                {
                    "name": "文档",
                    "target": "02-文档",
                    "extensions": [".docx", ".pdf"],
                    "keywords": ["报告"],
                },
            ],
        }
    )


def test_rule_requires_all_conditions() -> None:
    rule = ClassifyRule.from_dict(
        {"name": "报告", "target": "02-报告", "extensions": [".docx"], "keywords": ["报告"]}
    )
    assert rule.matches("年度报告.docx")
    assert not rule.matches("年度报告.pdf")  # 后缀不符
    assert not rule.matches("会议记录.docx")  # 关键词不符


def test_rule_without_conditions_never_matches() -> None:
    rule = ClassifyRule.from_dict({"name": "空规则", "target": "X"})
    assert not rule.matches("任意文件.docx")


def test_extension_without_dot_is_normalized() -> None:
    rule = ClassifyRule.from_dict({"name": "r", "target": "X", "extensions": ["XLSX"]})
    assert rule.matches("报表.xlsx")


def test_pattern_matching() -> None:
    rule = ClassifyRule.from_dict({"name": "r", "target": "X", "patterns": ["*汇总*"]})
    assert rule.matches("2026年汇总表.docx")
    assert not rule.matches("通知.docx")


def test_config_uses_first_matching_rule() -> None:
    config = ClassifyConfig.from_dict(
        {
            "rules": [
                {"name": "a", "target": "A", "extensions": [".docx"]},
                {"name": "b", "target": "B", "extensions": [".docx"]},
            ]
        }
    )
    assert config.match("x.docx").target == "A"


def test_plan_moves_does_not_touch_files(tmp_path: Path) -> None:
    (tmp_path / "报表.xlsx").write_bytes(b"x")
    plans = plan_moves(tmp_path, _config())
    assert len(plans) == 1
    assert plans[0].target.parent.name == "01-报表"
    assert (tmp_path / "报表.xlsx").exists()


def test_organize_dry_run_keeps_files(tmp_path: Path) -> None:
    (tmp_path / "报表.xlsx").write_bytes(b"x")
    plans = organize(tmp_path, _config(), dry_run=True)
    assert not any(plan.applied for plan in plans)
    assert (tmp_path / "报表.xlsx").exists()
    assert "归档计划" in summarize(plans, dry_run=True)


def test_organize_apply_moves_files(tmp_path: Path) -> None:
    (tmp_path / "报表.xlsx").write_bytes(b"x")
    (tmp_path / "年度报告.docx").write_bytes(b"x")
    (tmp_path / "随手笔记.md").write_text("内容", encoding="utf-8")

    plans = organize(tmp_path, _config(), dry_run=False)

    assert {plan.target.parent.name for plan in plans} == {"01-报表", "02-文档", "99-其他"}
    assert (tmp_path / "归档" / "01-报表" / "报表.xlsx").exists()
    assert (tmp_path / "归档" / "02-文档" / "年度报告.docx").exists()
    assert (tmp_path / "归档" / "99-其他" / "随手笔记.md").exists()
    assert not (tmp_path / "报表.xlsx").exists()


def test_copy_mode_keeps_source(tmp_path: Path) -> None:
    (tmp_path / "报表.xlsx").write_bytes(b"x")
    organize(tmp_path, _config(copy=True), dry_run=False)
    assert (tmp_path / "报表.xlsx").exists()
    assert (tmp_path / "归档" / "01-报表" / "报表.xlsx").exists()


def test_already_archived_files_are_skipped(tmp_path: Path) -> None:
    archive = tmp_path / "归档" / "01-报表"
    archive.mkdir(parents=True)
    (archive / "旧报表.xlsx").write_bytes(b"x")
    assert plan_moves(tmp_path, _config()) == []


def test_name_conflict_gets_suffix(tmp_path: Path) -> None:
    (tmp_path / "报表.xlsx").write_bytes(b"x")
    archive = tmp_path / "归档" / "01-报表"
    archive.mkdir(parents=True)
    (archive / "报表.xlsx").write_bytes(b"y")

    plans = plan_moves(tmp_path, _config())
    assert plans[0].target.name == "报表_1.xlsx"


def test_config_from_file(tmp_path: Path) -> None:
    path = tmp_path / "classify.json"
    path.write_text(
        json.dumps({"target_root": "已整理", "rules": [{"name": "r", "target": "T", "extensions": [".txt"]}]}),
        encoding="utf-8",
    )
    config = ClassifyConfig.from_file(path)
    assert config.target_root == "已整理"
    assert config.rules[0].target == "T"
