from __future__ import annotations

from collections.abc import Iterable
from typing import Any


def map_records(
    records: Iterable[dict[str, Any]],
    target_headers: list[str],
    field_mapping: dict[str, list[str]],
) -> list[dict[str, Any]]:
    """根据候选来源字段，把记录整理成目标字段集合。"""
    mapped_records: list[dict[str, Any]] = []

    for record in records:
        normalized = {str(key).strip(): value for key, value in record.items()}
        mapped = {}
        for target in target_headers:
            candidates = field_mapping.get(target, [target])
            mapped[target] = next(
                (normalized[name] for name in candidates if name in normalized),
                "",
            )
        mapped_records.append(mapped)

    return mapped_records
