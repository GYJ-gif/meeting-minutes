#!/usr/bin/env python3
"""Validate a generated meeting-minutes DOCX."""
from __future__ import annotations
import argparse, zipfile
from pathlib import Path
from docx import Document
from docx.oxml.ns import qn
from build_minutes import TABLE_GRID_DXA, TABLE_HEADERS, build_title, today_shanghai

def validate_minutes(path, topics=None, recorder=None, participants=None):
    path = Path(path).resolve(); issues = []
    if not path.is_file(): return [f"文件不存在：{path}"]
    try:
        with zipfile.ZipFile(path) as archive:
            bad = archive.testzip()
            if bad: issues.append(f"DOCX 压缩包损坏：{bad}")
    except zipfile.BadZipFile: return ["文件不是有效的 DOCX/ZIP 包"]
    try: doc = Document(path)
    except Exception as exc: return [f"python-docx 无法打开文件：{exc}"]
    paragraphs = [p.text for p in doc.paragraphs]; text = "\n".join(paragraphs)
    current = today_shanghai(); expected_date = f"会议时间：{current.year}年{current.month}月{current.day}日"
    if expected_date not in text: issues.append(f"会议日期不是北京时间当天：应包含“{expected_date}”")
    topics = [str(x).strip() for x in topics or [] if str(x).strip()]
    if topics and build_title(topics) not in text: issues.append(f"标题未按主题生成：{build_title(topics)}")
    for topic in topics:
        if topic not in text: issues.append(f"未覆盖会议主题：{topic}")
    recorder_line = next((x for x in paragraphs if x.startswith("纪要人员：")), "")
    if recorder and recorder_line != f"纪要人员：{recorder.strip()}": issues.append(f"纪要人员不匹配：{recorder}")
    if not recorder and recorder_line in ("", "纪要人员：", "纪要人员：待核验", "纪要人员：待确认"): issues.append("纪要人员为空或未确认")
    participant_line = next((x for x in paragraphs if x.startswith("参与人员：")), "")
    expected_people = "、".join(str(x).strip() for x in participants or [] if str(x).strip())
    if expected_people and participant_line != f"参与人员：{expected_people}": issues.append(f"参会人员不匹配：{expected_people}")
    if not expected_people and participant_line in ("", "参与人员：", "参与人员：待核验", "参与人员：待确认"): issues.append("参会人员为空或未确认")
    for item in ("一、会议基本信息", "二、会议主要内容", "三、会议核心结论（如有）", "四、跟进事宜及节点", "下次交流预计时间：", "预计议题：", "交流图片", "如涉及比较好的学习材料和要点可分享给大家共同学习"):
        if item not in text: issues.append(f"缺少模板固定内容：{item}")
    if len(doc.tables) != 1:
        issues.append(f"跟进事项表数量应为1，实际为{len(doc.tables)}"); return issues
    table = doc.tables[0]
    if [c.text for c in table.rows[0].cells] != TABLE_HEADERS: issues.append("跟进事项表表头不匹配")
    grid = [int(n.get(qn("w:w"))) for n in table._tbl.tblGrid]
    if grid != TABLE_GRID_DXA: issues.append(f"表格列宽不匹配：{grid}")
    for ri, row in enumerate(table.rows, 1):
        for ci, cell in enumerate(row.cells, 1):
            shd = cell._tc.get_or_add_tcPr().find(qn("w:shd")); fill = shd.get(qn("w:fill")) if shd is not None else None
            if fill not in (None, "auto", "FFFFFF"): issues.append(f"表格第{ri}行第{ci}列存在填充色：{fill}")
    return issues

def main():
    parser = argparse.ArgumentParser(description=__doc__); parser.add_argument("docx")
    parser.add_argument("--topic", action="append", default=[]); parser.add_argument("--recorder")
    parser.add_argument("--participant", action="append", default=[]); args = parser.parse_args()
    issues = validate_minutes(args.docx, args.topic, args.recorder, args.participant)
    if issues:
        for issue in issues: print(f"FAIL: {issue}")
        return 1
    print("OK: 标题、日期、人员、新模板结构和无填色表格均符合要求"); return 0

if __name__ == "__main__": raise SystemExit(main())

