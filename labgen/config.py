"""Загрузка конфигурации: значения по умолчанию соответствуют ГОСТ 7.32-2017."""

from __future__ import annotations

import copy
from pathlib import Path
from typing import Any

import yaml

DEFAULTS: dict[str, Any] = {
    "format": {
        "font": "Times New Roman",
        "size": 14,
        "line_spacing": 1.5,
        "first_line_indent": 1.25,
        # поля страницы в миллиметрах
        "margins": {"left": 30, "right": 15, "top": 20, "bottom": 20},
        "page_numbers": True,
        # bottom-center | bottom-right | top-center | top-right
        "page_number_position": "bottom-center",
        "section_page_break": True,
        "number_headings": True,
        "unnumbered_headings": [
            "содержание",
            "реферат",
            "аннотация",
            "введение",
            "заключение",
            "список литературы",
            "список использованных источников",
            "список использованной литературы",
            "приложение",
            "нормативные ссылки",
            "термины и определения",
            "определения, обозначения и сокращения",
        ],
        "code": {"font": "Courier New", "size": 12, "line_spacing": 1.0},
        "table": {"size": 12},
        "captions": {
            "figure": "Рисунок",
            "table": "Таблица",
            "listing": "Листинг",
            "dash": "—",
        },
    },
    "toc": True,
    "title_page": {
        "enabled": True,
        "ministry": "МИНИСТЕРСТВО НАУКИ И ВЫСШЕГО ОБРАЗОВАНИЯ РОССИЙСКОЙ ФЕДЕРАЦИИ",
        "university": "",
        "faculty": "",
        "department": "",
        "work_type": "ОТЧЁТ",
        "work_kind": "",
        "discipline": "",
        "topic": "",
        "author": {"label": "Выполнил:", "name": "", "group": ""},
        "supervisor": {"label": "Проверил:", "name": "", "position": ""},
        "city": "",
        "year": None,
    },
}


def _deep_merge(base: dict[str, Any], override: dict[str, Any]) -> dict[str, Any]:
    """Рекурсивно накладывает override на копию base."""
    result = copy.deepcopy(base)
    for key, value in (override or {}).items():
        if isinstance(value, dict) and isinstance(result.get(key), dict):
            result[key] = _deep_merge(result[key], value)
        else:
            result[key] = value
    return result


class Config:
    """Конфигурация сборки отчёта."""

    def __init__(self, data: dict[str, Any] | None = None) -> None:
        self.data = _deep_merge(DEFAULTS, data or {})

    @classmethod
    def load(cls, path: str | Path | None) -> Config:
        if path is None:
            return cls()
        path = Path(path)
        if not path.exists():
            raise FileNotFoundError(f"Файл конфигурации не найден: {path}")
        with path.open(encoding="utf-8") as handle:
            loaded = yaml.safe_load(handle) or {}
        if not isinstance(loaded, dict):
            raise ValueError(f"{path}: ожидался YAML-объект на верхнем уровне")
        return cls(loaded)

    def merged_with(self, override: dict[str, Any]) -> Config:
        """Новый Config с наложенными поверх значениями (например, из front matter)."""
        return Config(_deep_merge(self.data, override))

    @property
    def fmt(self) -> dict[str, Any]:
        return self.data["format"]

    @property
    def title_page(self) -> dict[str, Any]:
        return self.data["title_page"]

    @property
    def captions(self) -> dict[str, str]:
        return self.fmt["captions"]

    @property
    def toc(self) -> bool:
        return bool(self.data.get("toc"))


def find_config(explicit: str | Path | None, near: Path) -> Path | None:
    """Ищет labgen.yaml: явно указанный путь, рядом с md-файлом, затем в текущей папке."""
    if explicit:
        return Path(explicit)
    for candidate in (near.parent / "labgen.yaml", near.parent / "labgen.yml", Path("labgen.yaml")):
        if candidate.exists():
            return candidate
    return None
