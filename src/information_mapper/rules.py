"""规则配置：默认值、界面结构与配置字典的互转、临时文件生成。

界面中的规则编辑器把使用者的设置写成临时 JSON（位于系统临时目录），
执行时按普通配置文件读取。这样无需手工编辑项目内的 JSON 文件。
"""

from __future__ import annotations

import json
import re
import tempfile
from pathlib import Path
from typing import Any, Iterable, Sequence

TEMP_SUBDIR = "info-map-rules"

_LIST_SEPARATOR = re.compile(r"[,，;；\s]+")

# 界面编辑行的形状
MappingRow = tuple[str, str]  # (目标字段, 来源字段候选)
ClassifyRow = tuple[str, str, str, str]  # (规则名, 目标子目录, 扩展名, 文件名关键词)

DEFAULT_MAPPING: dict[str, Any] = {
    "target_headers": ["姓名", "联系电话", "家庭住址"],
    "field_mapping": {
        "姓名": ["姓名", "户主姓名", "人员姓名"],
        "联系电话": ["联系电话", "手机号", "电话"],
        "家庭住址": ["家庭住址", "住址", "地址"],
    },
}

DEFAULT_CLASSIFY: dict[str, Any] = {
    "target_root": "归档",
    "default_target": "99-其他",
    "copy": False,
    "rules": [
        {
            "name": "报表",
            "target": "01-报表",
            "extensions": [".xlsx", ".xls", ".csv", ".tsv"],
        },
        {
            "name": "报告文档",
            "target": "02-报告",
            "extensions": [".docx", ".pdf"],
            "keywords": ["报告", "总结", "汇报"],
        },
        {
            "name": "文本数据",
            "target": "03-文本",
            "extensions": [".txt", ".md", ".json"],
        },
    ],
}


# ---------------------------------------------------------------- 文本与列表
def split_list(text: str) -> list[str]:
    """把 ``a, b；c`` 形式的输入拆成列表。"""
    return [item for item in _LIST_SEPARATOR.split(str(text or "")) if item]


def join_list(items: Iterable[str]) -> str:
    """把列表拼成 ``a, b`` 形式，供界面显示。"""
    return ", ".join(str(item) for item in items)


def normalize_extension(value: str) -> str:
    """``xlsx`` / ``.XLSX`` → ``.xlsx``。"""
    text = str(value).strip().lower()
    return text if text.startswith(".") else f".{text}"


# ---------------------------------------------------------------- 字段映射
def mapping_rows(config: dict[str, Any]) -> list[MappingRow]:
    """映射配置 → 界面行列表。"""
    headers = [str(header) for header in config.get("target_headers") or []]
    mapping = config.get("field_mapping") or {}
    rows: list[MappingRow] = []
    for header in headers:
        candidates = mapping.get(header) or [header]
        rows.append((header, join_list(candidates)))
    for key, value in mapping.items():  # 配置中存在但表头未声明的字段
        if str(key) not in headers:
            rows.append((str(key), join_list(value or [key])))
    return rows


def build_mapping(rows: Sequence[MappingRow]) -> dict[str, Any]:
    """界面行列表 → 映射配置。"""
    headers: list[str] = []
    mapping: dict[str, list[str]] = {}
    for target, candidates in rows:
        name = str(target).strip()
        if not name:
            continue
        headers.append(name)
        mapping[name] = split_list(candidates) or [name]
    return {"target_headers": headers, "field_mapping": mapping}


# ---------------------------------------------------------------- 归档规则
def classify_rows(config: dict[str, Any]) -> list[ClassifyRow]:
    """归档配置 → 界面行列表。"""
    rows: list[ClassifyRow] = []
    for rule in config.get("rules") or []:
        target = str(rule.get("target") or "")
        rows.append(
            (
                str(rule.get("name") or target),
                target,
                join_list(rule.get("extensions") or []),
                join_list(rule.get("keywords") or []),
            )
        )
    return rows


def build_classify(
    rows: Sequence[ClassifyRow],
    *,
    target_root: str = "归档",
    default_target: str = "99-其他",
    copy: bool = False,
) -> dict[str, Any]:
    """界面行列表 → 归档配置。"""
    rules: list[dict[str, Any]] = []
    for name, target, extensions, keywords in rows:
        folder = str(target).strip() or str(name).strip()
        if not folder:
            continue
        rule: dict[str, Any] = {"name": str(name).strip() or folder, "target": folder}
        extensions_list = [normalize_extension(item) for item in split_list(extensions)]
        if extensions_list:
            rule["extensions"] = extensions_list
        keywords_list = split_list(keywords)
        if keywords_list:
            rule["keywords"] = keywords_list
        rules.append(rule)
    return {
        "target_root": str(target_root).strip() or "归档",
        "default_target": str(default_target).strip() or "99-其他",
        "copy": bool(copy),
        "rules": rules,
    }


# ---------------------------------------------------------------- 临时文件
def temp_rules_dir() -> Path:
    """规则临时目录（不存在时创建）。"""
    path = Path(tempfile.gettempdir()) / TEMP_SUBDIR
    path.mkdir(parents=True, exist_ok=True)
    return path


def write_temp_rules(kind: str, payload: dict[str, Any]) -> Path:
    """把规则写成临时 JSON，返回文件路径。"""
    path = temp_rules_dir() / f"{kind}.json"
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def read_rules(path: str | Path) -> dict[str, Any]:
    """读取规则文件。"""
    config_path = Path(path)
    if not config_path.is_file():
        raise ValueError(f"找不到规则文件：{config_path}")
    return json.loads(config_path.read_text(encoding="utf-8"))
