"""Низкоуровневые операции над OOXML, которых нет в python-docx."""

from __future__ import annotations

from typing import Any

from docx.document import Document as DocxDocument
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Mm, Pt

PAGE_WIDTH_MM = 210
PAGE_HEIGHT_MM = 297

ALIGN = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}


def set_font(target: Any, name: str, size: float | None = None) -> None:
    """Проставляет гарнитуру во всех четырёх слотах w:rFonts — иначе кириллица
    может отрисоваться шрифтом темы, а не заданным."""
    font = target.font
    font.name = name
    if size is not None:
        font.size = Pt(size)
    element = target._element
    rpr = element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    for slot in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rfonts.set(qn(slot), name)


def add_field(paragraph, instr: str, placeholder: str = " ", dirty: bool = True) -> None:
    """Вставляет поле Word (PAGE, TOC, ...).

    dirty=True заставляет Word пересчитать поле при открытии документа."""
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), instr)
    if dirty:
        fld.set(qn("w:dirty"), "true")
    run = OxmlElement("w:r")
    text = OxmlElement("w:t")
    text.text = placeholder
    run.append(text)
    fld.append(run)
    paragraph._p.append(fld)


def set_paragraph_format(paragraph, cfg: dict[str, Any], *, indent: bool = True,
                         alignment: str = "justify", spacing: float | None = None) -> None:
    pf = paragraph.paragraph_format
    pf.alignment = ALIGN[alignment]
    pf.first_line_indent = Cm(cfg["first_line_indent"]) if indent else Cm(0)
    pf.left_indent = Cm(0)
    pf.right_indent = Cm(0)
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.line_spacing = spacing if spacing is not None else cfg["line_spacing"]


def setup_page(document: DocxDocument, cfg: dict[str, Any], *, title_page: bool) -> None:
    """Поля страницы, нумерация, особый первый лист."""
    section = document.sections[0]
    section.page_width = Mm(PAGE_WIDTH_MM)
    section.page_height = Mm(PAGE_HEIGHT_MM)
    margins = cfg["margins"]
    section.left_margin = Mm(margins["left"])
    section.right_margin = Mm(margins["right"])
    section.top_margin = Mm(margins["top"])
    section.bottom_margin = Mm(margins["bottom"])

    if not cfg.get("page_numbers", True):
        return

    # Титульный лист включается в нумерацию, но номер на нём не ставится (ГОСТ 7.32-2017, п. 6.1.5).
    section.different_first_page_header_footer = bool(title_page)

    position = cfg.get("page_number_position", "bottom-center")
    vertical, _, horizontal = position.partition("-")
    container = section.footer if vertical == "bottom" else section.header
    container.is_linked_to_previous = False
    paragraph = container.paragraphs[0] if container.paragraphs else container.add_paragraph()
    paragraph.alignment = ALIGN.get(horizontal or "center", WD_ALIGN_PARAGRAPH.CENTER)
    pf = paragraph.paragraph_format
    pf.first_line_indent = Cm(0)
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    add_field(paragraph, "PAGE", placeholder="1", dirty=False)
    for run in paragraph.runs:
        set_font(run, cfg["font"], cfg["size"])
    # шрифт для текста внутри поля
    for run_element in paragraph._p.findall(".//" + qn("w:r")):
        rpr = run_element.get_or_add_rPr()
        rfonts = rpr.find(qn("w:rFonts"))
        if rfonts is None:
            rfonts = OxmlElement("w:rFonts")
            rpr.append(rfonts)
        for slot in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
            rfonts.set(qn(slot), cfg["font"])
        size = rpr.find(qn("w:sz"))
        if size is None:
            size = OxmlElement("w:sz")
            rpr.append(size)
        size.set(qn("w:val"), str(int(cfg["size"] * 2)))


def setup_styles(document: DocxDocument, cfg: dict[str, Any]) -> None:
    """Базовый стиль и заголовки — по ГОСТ; заголовки остаются Heading N,
    чтобы поле TOC собрало содержание автоматически."""
    normal = document.styles["Normal"]
    set_font(normal, cfg["font"], cfg["size"])
    pf = normal.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf.first_line_indent = Cm(cfg["first_line_indent"])
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.line_spacing = cfg["line_spacing"]

    for level in range(1, 5):
        style = document.styles[f"Heading {level}"]
        set_font(style, cfg["font"], cfg["size"])
        style.font.bold = True
        style.font.italic = False
        style.font.all_caps = False
        style.font.color.rgb = None
        style.font.color.theme_color = None
        hpf = style.paragraph_format
        hpf.alignment = WD_ALIGN_PARAGRAPH.LEFT
        hpf.first_line_indent = Cm(cfg["first_line_indent"])
        hpf.left_indent = Cm(0)
        hpf.keep_with_next = True
        hpf.space_before = Pt(0 if level == 1 else 14)
        hpf.space_after = Pt(14)
        hpf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        hpf.line_spacing = cfg["line_spacing"]


def add_page_break(document: DocxDocument) -> None:
    paragraph = document.add_paragraph()
    paragraph.paragraph_format.first_line_indent = Cm(0)
    paragraph.add_run().add_break(WD_BREAK.PAGE)


def text_width_mm(cfg: dict[str, Any]) -> float:
    margins = cfg["margins"]
    return PAGE_WIDTH_MM - margins["left"] - margins["right"]
