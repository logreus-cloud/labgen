"""Титульный лист."""

from __future__ import annotations

import datetime as _dt
from typing import Any

from docx.document import Document as DocxDocument
from docx.enum.text import WD_LINE_SPACING
from docx.shared import Cm, Pt

from .docxutil import ALIGN, set_font


def _line(document: DocxDocument, text: str, cfg: dict[str, Any], *, align: str = "center",
          bold: bool = False, size: float | None = None, caps: bool = False):
    paragraph = document.add_paragraph()
    pf = paragraph.paragraph_format
    pf.alignment = ALIGN[align]
    pf.first_line_indent = Cm(0)
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
    if text:
        run = paragraph.add_run(text.upper() if caps else text)
        run.bold = bold
        set_font(run, cfg["font"], size or cfg["size"])
    return paragraph


def _blank(document: DocxDocument, cfg: dict[str, Any], count: int = 1) -> None:
    for _ in range(count):
        _line(document, "", cfg)


def _quoted(value: str) -> str:
    value = (value or "").strip()
    if not value:
        return ""
    if value[0] in "«\"'":
        return value
    return f"«{value}»"


def render_title_page(document: DocxDocument, page: dict[str, Any], cfg: dict[str, Any]) -> None:
    """Рисует титульный лист по данным из конфигурации."""
    for key in ("ministry", "university", "faculty", "department"):
        value = (page.get(key) or "").strip()
        if value:
            _line(document, value, cfg, caps=(key == "ministry"))
        if key == "university" and value:
            _blank(document, cfg)

    _blank(document, cfg, 6)

    work_type = (page.get("work_type") or "").strip()
    if work_type:
        _line(document, work_type, cfg, bold=True, size=cfg["size"] + 2, caps=True)
        _blank(document, cfg)

    for text in (
        (page.get("work_kind") or "").strip(),
        f"по дисциплине {_quoted(page.get('discipline'))}" if page.get("discipline") else "",
    ):
        if text:
            _line(document, text, cfg)

    topic = (page.get("topic") or "").strip()
    if topic:
        _line(document, f"на тему {_quoted(topic)}", cfg)

    _blank(document, cfg, 5)

    author = page.get("author") or {}
    supervisor = page.get("supervisor") or {}
    author_lines = [
        author.get("label") or "Выполнил:",
        f"студент группы {author['group']}" if author.get("group") else "",
        author.get("name") or "",
    ]
    supervisor_lines = [
        supervisor.get("label") or "Проверил:",
        supervisor.get("position") or "",
        supervisor.get("name") or "",
    ]
    for text in author_lines:
        if text:
            _line(document, text, cfg, align="right")
    if any(supervisor_lines):
        _blank(document, cfg)
    for text in supervisor_lines:
        if text:
            _line(document, text, cfg, align="right")

    _blank(document, cfg, 4)

    year = page.get("year") or _dt.date.today().year
    city = (page.get("city") or "").strip()
    footer = " ".join(part for part in (city, str(year)) if part)
    if footer:
        _line(document, footer, cfg)
