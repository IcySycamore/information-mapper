"""文件自动归档整理。

对应实践方案中的第三个业务场景：**文件资料管理 —— 多版本文件人工归档，流程繁琐**。

按规则把散落的文件移动到（或复制到）分类目录，规则来自 JSON 配置：

.. code-block:: json

    {
      "target_root": "output/归档",
      "default_target": "其他",
      "copy": false,
      "rules": [
        {"name": "报表", "extensions": [".xlsx", ".csv"], "keywords": ["报表", "汇总"],
         "target": "01-报表"},
        {"name": "报告", "extensions": [".docx", ".pdf"], "target": "02-报告"}
      ]
    }

同一个规则内 ``extensions`` / ``keywords`` / ``patterns`` 之间是"与"关系，
规则之间按顺序匹配，第一条命中的规则生效。
"""

from __future__ import annotations

import fnmatch
import json
import shutil
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable

from .formats import normalize_suffix

DEFAULT_TARGET = "其他"


@dataclass
class ClassifyRule:
    """单条归档规则。"""

    name: str
    target: str
    extensions: tuple[str, ...] = ()
    keywords: tuple[str, ...] = ()
    patterns: tuple[str, ...] = ()

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "ClassifyRule":
        target = str(data.get("target") or data.get("name") or DEFAULT_TARGET)
        return cls(
            name=str(data.get("name") or target),
            target=target,
            extensions=tuple(_normalize_extension(item) for item in data.get("extensions", [])),
            keywords=tuple(str(keyword) for keyword in data.get("keywords", [])),
            patterns=tuple(str(pattern) for pattern in data.get("patterns", [])),
        )

    def matches(self, filename: str) -> bool:
        suffix = normalize_suffix(filename)
        if self.extensions and suffix not in self.extensions:
            return False
        if self.keywords and not any(keyword in filename for keyword in self.keywords):
            return False
        if self.patterns and not any(fnmatch.fnmatch(filename, pattern) for pattern in self.patterns):
            return False
        # 三个条件都为空时视为"兜底规则"，不匹配任何文件，避免误吞
        return bool(self.extensions or self.keywords or self.patterns)


@dataclass
class ClassifyConfig:
    """归档配置。"""

    rules: list[ClassifyRule] = field(default_factory=list)
    target_root: str = "归档"
    default_target: str = DEFAULT_TARGET
    copy: bool = False

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "ClassifyConfig":
        return cls(
            rules=[ClassifyRule.from_dict(item) for item in data.get("rules", [])],
            target_root=str(data.get("target_root") or "归档"),
            default_target=str(data.get("default_target") or DEFAULT_TARGET),
            copy=bool(data.get("copy", False)),
        )

    @classmethod
    def from_file(cls, path: str | Path) -> "ClassifyConfig":
        config_path = Path(path)
        return cls.from_dict(json.loads(config_path.read_text(encoding="utf-8")))

    def match(self, filename: str) -> ClassifyRule | None:
        for rule in self.rules:
            if rule.matches(filename):
                return rule
        return None


@dataclass
class MovePlan:
    """一条待执行的归档动作。"""

    source: Path
    target: Path
    rule: str
    applied: bool = False
    error: str = ""

    def summary(self) -> str:
        action = "已归档" if self.applied else "待归档"
        return f"[{action}] {self.source.name} → {self.target.parent.name}/{self.target.name}（规则：{self.rule}）"


def plan_moves(
    source_dir: str | Path,
    config: ClassifyConfig,
    *,
    root: str | Path | None = None,
    recursive: bool = True,
    include_hidden: bool = False,
) -> list[MovePlan]:
    """生成归档计划（不移动任何文件）。"""
    directory = Path(source_dir)
    if not directory.is_dir():
        raise ValueError(f"不是有效目录：{directory}")

    destination_root = Path(root) if root else directory / config.target_root
    candidates = directory.rglob("*") if recursive else directory.glob("*")

    plans: list[MovePlan] = []
    for path in sorted(candidates):
        if not path.is_file():
            continue
        if not include_hidden and path.name.startswith((".", "~$")):
            continue
        # 已归档的文件（位于归档根目录内）跳过，避免反复搬运
        if destination_root == directory or destination_root in path.parents:
            continue
        rule = config.match(path.name)
        folder = rule.target if rule else config.default_target
        target = destination_root / folder / path.name
        plans.append(MovePlan(source=path, target=_unique_path(target), rule=rule.name if rule else "默认"))
    return plans


def apply_moves(plans: Iterable[MovePlan], *, dry_run: bool = True, copy: bool = False) -> list[MovePlan]:
    """执行归档计划；``dry_run=True`` 时只返回计划不落盘。"""
    executed: list[MovePlan] = []
    for plan in plans:
        if dry_run:
            executed.append(plan)
            continue
        try:
            plan.target.parent.mkdir(parents=True, exist_ok=True)
            if copy:
                shutil.copy2(plan.source, plan.target)
            else:
                shutil.move(str(plan.source), str(plan.target))
            plan.applied = True
        except OSError as error:
            plan.error = str(error)
        executed.append(plan)
    return executed


def organize(
    source_dir: str | Path,
    config: ClassifyConfig,
    *,
    dry_run: bool = True,
    root: str | Path | None = None,
    recursive: bool = True,
) -> list[MovePlan]:
    """一步完成"生成计划 + 执行归档"。"""
    plans = plan_moves(source_dir, config, root=root, recursive=recursive)
    return apply_moves(plans, dry_run=dry_run, copy=config.copy)


def summarize(plans: Iterable[MovePlan], *, dry_run: bool = True) -> str:
    """把归档结果整理成可读摘要。"""
    items = list(plans)
    head = f"归档计划 {len(items)} 个文件" if dry_run else f"已归档 {sum(1 for i in items if i.applied)}/{len(items)} 个文件"
    lines = [head]
    lines.extend(f"  {item.summary()}" for item in items)
    errors = [item for item in items if item.error]
    lines.extend(f"  出错：{item.source.name}（{item.error}）" for item in errors)
    return "\n".join(lines)


def _normalize_extension(value: Any) -> str:
    """把 ``docx`` / ``.DOCX`` 统一成 ``.docx``。"""
    text = str(value).strip().lower()
    return text if text.startswith(".") else f".{text}"


def _unique_path(path: Path) -> Path:
    """重名文件自动加序号，避免覆盖。"""
    if not path.exists():
        return path
    index = 1
    while True:
        candidate = path.with_name(f"{path.stem}_{index}{path.suffix}")
        if not candidate.exists():
            return candidate
        index += 1
