#!/usr/bin/env python3
"""Validate a generated materials-department meeting-minutes DOCX."""

from __future__ import annotations

import argparse
import zipfile
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

from build_minutes import RECORDER, TABLE_GRID_DXA, TABLE_HEADERS, today_shanghai


def validate_minutes(path: str | Path, topics: list[str] | None = None) -> list[str]:
    docx_path = Path(path).resolve()
    issues: list[str] = []
    if not docx_path.is_file():
        return [f"文件不存在：{docx_path}"]

    try:
        with zipfile.ZipFile(docx_path) as archive:
            bad_member = archive.testzip()
            if bad_member:
                issues.append(f"DOCX 压缩包损坏：{bad_member}")
    except zipfile.BadZipFile:
        return ["文件不是有效的 DOCX/ZIP 包"]

    try:
        doc = Document(docx_path)
    except Exception as exc:
        return [f"python-docx 无法打开文件：{exc}"]

    text = "\n".join(paragraph.text for paragraph in doc.paragraphs)
    current = today_shanghai()
    expected_date = f"会议时间：{current.year}年{current.month}月{current.day}日"
    if expected_date not in text:
        issues.append(f"会议日期不是北京时间当天：应包含“{expected_date}”")
    if f"纪要人员：{RECORDER}" not in text:
        issues.append(f"纪要人员不是固定值“{RECORDER}”")
    for topic in topics or []:
        if topic and topic not in text:
            issues.append(f"未覆盖会议主题：{topic}")

    if len(doc.tables) != 1:
        issues.append(f"跟进事项表数量应为1，实际为{len(doc.tables)}")
        return issues

    table = doc.tables[0]
    headers = [cell.text for cell in table.rows[0].cells]
    if headers != TABLE_HEADERS:
        issues.append(f"表头不匹配：{headers}")
    grid = [int(node.get(qn("w:w"))) for node in table._tbl.tblGrid]
    if grid != TABLE_GRID_DXA:
        issues.append(f"表格列宽不匹配：{grid}")
    for row_index, row in enumerate(table.rows, start=1):
        for column_index, cell in enumerate(row.cells, start=1):
            shading = cell._tc.get_or_add_tcPr().find(qn("w:shd"))
            fill = shading.get(qn("w:fill")) if shading is not None else None
            if fill not in (None, "auto", "FFFFFF"):
                issues.append(
                    f"表格第{row_index}行第{column_index}列存在填充色：{fill}"
                )
    return issues


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("docx", help="待检查的 DOCX 文件")
    parser.add_argument("--topic", action="append", default=[], help="应在文档中出现的主题")
    return parser.parse_args()


def main() -> int:
    args = _parse_args()
    issues = validate_minutes(args.docx, args.topic)
    if issues:
        for issue in issues:
            print(f"FAIL: {issue}")
        return 1
    print("OK: 会议纪要结构、日期、纪要人员和无填色表格均符合要求")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

