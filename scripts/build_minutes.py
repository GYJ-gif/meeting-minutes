#!/usr/bin/env python3
"""Build a materials-department meeting-minutes DOCX from structured content."""

from __future__ import annotations

import argparse
import json
from datetime import date, datetime
from pathlib import Path
from typing import Any
from zoneinfo import ZoneInfo

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor


RECORDER = "郭源杰"
TABLE_HEADERS = ["序号", "跟进事宜", "负责人", "截止日期", "备注"]
TABLE_GRID_DXA = [817, 4527, 1023, 1295, 1580]
TEMPLATE_PATH = (
    Path(__file__).resolve().parents[1]
    / "assets"
    / "materials-department-minutes-template.docx"
)


def today_shanghai() -> date:
    return datetime.now(ZoneInfo("Asia/Shanghai")).date()


def _validate_content(content: dict[str, Any]) -> None:
    topics = content.get("topics") or []
    if not isinstance(topics, list) or not any(str(item).strip() for item in topics):
        raise ValueError("会议主题不能为空")
    if not isinstance(content.get("sections", []), list):
        raise ValueError("sections 必须是列表")
    if not isinstance(content.get("conclusions", []), list):
        raise ValueError("conclusions 必须是列表")
    if not isinstance(content.get("actions", []), list):
        raise ValueError("actions 必须是列表")


def _clear_body(doc: Document) -> None:
    body = doc._element.body
    for child in list(body):
        if child.tag != qn("w:sectPr"):
            body.remove(child)


def _set_run_font(run, size: float = 12, bold: bool | None = None) -> None:
    run.font.name = "Times New Roman"
    run.font.size = Pt(size)
    run.font.bold = bold
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.get_or_add_rFonts()
    rfonts.set(qn("w:ascii"), "Times New Roman")
    rfonts.set(qn("w:hAnsi"), "Times New Roman")
    rfonts.set(qn("w:eastAsia"), "宋体")


def _format_paragraph(paragraph, *, after: float = 0, line: float = 1.5) -> None:
    fmt = paragraph.paragraph_format
    fmt.space_before = Pt(0)
    fmt.space_after = Pt(after)
    fmt.line_spacing = line


def _add_heading(doc: Document, text: str) -> None:
    paragraph = doc.add_paragraph(style="Heading 2")
    paragraph.paragraph_format.keep_with_next = True
    paragraph.paragraph_format.space_before = Pt(6)
    paragraph.paragraph_format.space_after = Pt(3)
    run = paragraph.add_run(text)
    _set_run_font(run, bold=True)


def _add_subheading(doc: Document, text: str) -> None:
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.keep_with_next = True
    paragraph.paragraph_format.space_before = Pt(3)
    paragraph.paragraph_format.space_after = Pt(2)
    run = paragraph.add_run(text)
    _set_run_font(run, bold=True)


def _add_body(doc: Document, text: str, *, indent: bool = True) -> None:
    paragraph = doc.add_paragraph()
    _format_paragraph(paragraph, after=3)
    if indent:
        paragraph.paragraph_format.first_line_indent = Pt(24)
    run = paragraph.add_run(str(text).strip())
    _set_run_font(run)


def _add_labeled_line(doc: Document, label: str, value: str) -> None:
    paragraph = doc.add_paragraph()
    _format_paragraph(paragraph)
    run = paragraph.add_run(label)
    _set_run_font(run, bold=True)
    run = paragraph.add_run(value)
    _set_run_font(run)


def _remove_cell_fill(cell) -> None:
    tcpr = cell._tc.get_or_add_tcPr()
    for shading in list(tcpr.findall(qn("w:shd"))):
        tcpr.remove(shading)


def _set_repeat_header(row) -> None:
    trpr = row._tr.get_or_add_trPr()
    marker = OxmlElement("w:tblHeader")
    marker.set(qn("w:val"), "true")
    trpr.append(marker)


def _set_table_geometry(table) -> None:
    table.autofit = False
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    tblpr = table._tbl.tblPr
    tblw = tblpr.find(qn("w:tblW"))
    if tblw is None:
        tblw = OxmlElement("w:tblW")
        tblpr.append(tblw)
    tblw.set(qn("w:w"), str(sum(TABLE_GRID_DXA)))
    tblw.set(qn("w:type"), "dxa")

    tblind = OxmlElement("w:tblInd")
    tblind.set(qn("w:w"), "120")
    tblind.set(qn("w:type"), "dxa")
    tblpr.append(tblind)

    grid = table._tbl.tblGrid
    for child in list(grid):
        grid.remove(child)
    for width in TABLE_GRID_DXA:
        node = OxmlElement("w:gridCol")
        node.set(qn("w:w"), str(width))
        grid.append(node)

    for row in table.rows:
        for cell, width in zip(row.cells, TABLE_GRID_DXA):
            tcpr = cell._tc.get_or_add_tcPr()
            tcw = tcpr.find(qn("w:tcW"))
            if tcw is None:
                tcw = OxmlElement("w:tcW")
                tcpr.append(tcw)
            tcw.set(qn("w:w"), str(width))
            tcw.set(qn("w:type"), "dxa")


def _add_actions_table(doc: Document, actions: list[dict[str, Any]]) -> None:
    row_count = max(4, len(actions))
    table = doc.add_table(rows=1, cols=5)
    table.style = "Table Grid"
    for index, header in enumerate(TABLE_HEADERS):
        table.rows[0].cells[index].text = header
    for row_index in range(row_count):
        cells = table.add_row().cells
        action = actions[row_index] if row_index < len(actions) else {}
        values = [
            str(row_index + 1),
            str(action.get("item", "")),
            str(action.get("owner", "")),
            str(action.get("deadline", "")),
            str(action.get("notes", "")),
        ]
        for index, value in enumerate(values):
            cells[index].text = value

    _set_repeat_header(table.rows[0])
    _set_table_geometry(table)
    for row_index, row in enumerate(table.rows):
        for column_index, cell in enumerate(row.cells):
            _remove_cell_fill(cell)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            for paragraph in cell.paragraphs:
                paragraph.alignment = (
                    WD_ALIGN_PARAGRAPH.CENTER
                    if column_index in (0, 2, 3)
                    else WD_ALIGN_PARAGRAPH.LEFT
                )
                _format_paragraph(paragraph, line=1.15)
                for run in paragraph.runs:
                    _set_run_font(run, bold=(row_index == 0))


def _normalize_review_items(content: dict[str, Any]) -> list[str]:
    items = [str(item).strip() for item in content.get("review_items", []) if str(item).strip()]
    defaults = ["参会人员名单", "专业名称及产品名称", "关键数值、负责人和截止日期"]
    for item in defaults:
        if len(items) >= 3:
            break
        if item not in items:
            items.append(item)
    return items[:5]


def build_minutes(
    content: dict[str, Any],
    output_path: str | Path,
    *,
    meeting_date: date | None = None,
) -> Path:
    """Create the final DOCX. Date and recorder fields cannot come from input content."""
    _validate_content(content)
    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(f"缺少模板资产：{TEMPLATE_PATH}")

    meeting_date = meeting_date or today_shanghai()
    output = Path(output_path).resolve()
    output.parent.mkdir(parents=True, exist_ok=True)

    doc = Document(TEMPLATE_PATH)
    _clear_body(doc)

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _format_paragraph(title, after=8, line=1.0)
    _set_run_font(title.add_run("材料部组会会议纪要"), size=16, bold=True)

    _add_heading(doc, "一、会议基本信息")
    _add_labeled_line(
        doc,
        "会议时间：",
        f"{meeting_date.year}年{meeting_date.month}月{meeting_date.day}日",
    )
    _add_labeled_line(doc, "会议地点：", str(content.get("location") or "待核验"))
    topics = [str(item).strip() for item in content["topics"] if str(item).strip()]
    _add_labeled_line(doc, "主要事宜：", "材料部组会（" + "、".join(topics) + "）")
    participants = content.get("participants") or []
    participant_text = "、".join(str(item).strip() for item in participants if str(item).strip())
    _add_labeled_line(doc, "参与人员：", participant_text or "待核验")
    _add_labeled_line(doc, "纪要人员：", RECORDER)

    _add_heading(doc, "二、会议主要内容")
    for index, section in enumerate(content.get("sections", []), start=1):
        section_title = str(section.get("title") or f"议题{index}").strip()
        _add_subheading(doc, f"{index}）{section_title}")
        paragraphs = section.get("paragraphs") or []
        if not paragraphs:
            paragraphs = ["相关内容待核验。"]
        for paragraph in paragraphs:
            _add_body(doc, str(paragraph))

    _add_heading(doc, "三、会议核心结论")
    conclusions = content.get("conclusions") or ["本次会议未形成可确认的核心结论。"]
    for conclusion in conclusions:
        paragraph = doc.add_paragraph(style="List Bullet")
        _format_paragraph(paragraph, after=2, line=1.35)
        _set_run_font(paragraph.add_run(str(conclusion)))

    _add_heading(doc, "四、跟进事宜及节点")
    _add_actions_table(doc, content.get("actions") or [])

    _add_heading(doc, "五、交流图片")
    _add_body(doc, "本次提供的逐字稿未附交流图片。", indent=False)

    _add_heading(doc, "六、资料来源与复核说明")
    source_note = str(content.get("source_note") or "用户提供的会议逐字稿 TXT").strip()
    _add_body(doc, f"资料来源：{source_note}。未进行外部联网检索。", indent=False)
    review_items = _normalize_review_items(content)
    _add_body(doc, "最需要人工复核的信息：" + "；".join(review_items) + "。", indent=False)
    _add_body(
        doc,
        "自动转写可能存在同音词、专业名称和数值识别偏差；无法确认的信息不得写成事实。",
        indent=False,
    )

    doc.core_properties.title = f"{meeting_date:%Y%m%d}材料部组会会议纪要"
    doc.core_properties.subject = "材料部组会"
    doc.core_properties.author = "材料部"
    doc.core_properties.keywords = "会议纪要, 材料部, 逐字稿"
    doc.save(output)
    return output


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--content", required=True, help="Codex 整理后的 UTF-8 JSON 文件")
    parser.add_argument("--transcript", required=True, help="原始会议逐字稿 TXT，用于来源校验")
    parser.add_argument("--topics", action="append", default=[], help="会议主题，可重复传入")
    parser.add_argument("--output", help="输出 DOCX；默认与逐字稿同目录")
    return parser.parse_args()


def main() -> int:
    args = _parse_args()
    transcript = Path(args.transcript).resolve()
    if not transcript.is_file():
        raise FileNotFoundError(f"逐字稿不存在：{transcript}")
    transcript_text = transcript.read_text(encoding="utf-8-sig")
    if not transcript_text.strip():
        raise ValueError("逐字稿 TXT 为空")

    content_path = Path(args.content).resolve()
    content = json.loads(content_path.read_text(encoding="utf-8-sig"))
    if args.topics:
        content["topics"] = [item.strip() for item in args.topics if item.strip()]
    content.setdefault("source_note", f"会议逐字稿《{transcript.name}》")

    meeting_date = today_shanghai()
    output = (
        Path(args.output).resolve()
        if args.output
        else transcript.parent / f"{meeting_date:%Y%m%d}材料部组会会议纪要.docx"
    )
    print(build_minutes(content, output))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

