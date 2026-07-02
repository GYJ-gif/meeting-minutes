---
name: meeting-minutes
description: Use when a user provides a Chinese meeting transcript TXT, meeting topics, recorder, and participants and wants a standardized DOCX based on the 2026 communication-minutes template.
---

# 交流会议纪要

## 概述

将会议逐字稿和指定信息整理为正式 DOCX。`scripts/build_minutes.py` 按生检模板排版，`scripts/validate_minutes.py` 负责校验。

**REQUIRED SUB-SKILL:** 使用 `documents` 完成 DOCX 渲染和逐页检查。

## 必填输入

每次调用必须取得以下四项，缺少任一项时停止并要求补充，不得猜测：

1. 会议逐字稿 `.txt`；
2. 一条或多条会议主题；
3. 纪要人员；
4. 非空参会人员名单。

下次交流预计时间和预计议题为可选输入，未提供时写“待确认”。

## 固定规则

- 日期取 `Asia/Shanghai` 当天，不从文件名或逐字稿推断。
- 标题按主题生成：单主题为“关于{主题}纪要”；2—3个主题全部列出；超过3个主题列出前三项并加“等议题”。
- 纪要人员和参会人员使用本次指令指定值。
- 跟进表沿用模板五列及 DXA 列宽 `[817, 4527, 1023, 1295, 1580]`。
- 表格不得使用填充颜色，包括表头。
- 无法确认的信息标注“待核验”或“待确认”，不得编造。
- 最终列出 3—5 项最需要人工复核的信息。

## 工作流

1. 校验四项必填输入。
2. 读取完整逐字稿，结合主题提取进展、结论和行动项。
3. 保守修正同音转写；证据不足时保留“待核验”。
4. 写入 UTF-8 JSON，不改写用户指定人员。
5. 调用 `scripts/build_minutes.py` 生成 DOCX。
6. 调用 `scripts/validate_minutes.py` 检查标题、日期、人员、模板章节、列宽和无填充颜色。
7. 使用 `documents` Skill 渲染全部页面；不可用时进行结构检查并披露。
8. 只交付最终 DOCX，除非用户要求中间文件。

## 内容 JSON

```json
{
  "recorder": "指定纪要人员",
  "participants": ["参会人员1", "参会人员2"],
  "location": "会议地点或待核验",
  "topics": ["主题1", "主题2"],
  "sections": [{"title": "主题1", "paragraphs": ["审慎整理的纪要段落。"]}],
  "conclusions": ["核心结论。"],
  "actions": [{"item": "跟进事项", "owner": "负责人或待确认", "deadline": "截止日期或待确认", "notes": "备注"}],
  "next_meeting_date": "可选；缺省为待确认",
  "next_topics": ["可选预计议题"],
  "review_items": ["专业名称", "关键数值", "截止日期"],
  "source_note": "会议逐字稿《文件名.txt》"
}
```

## 执行命令

```text
<python> scripts/build_minutes.py --content <content.json> --transcript <transcript.txt> --topics <主题> --recorder <纪要人员> --participant <参会人员> --output <output.docx>
<python> scripts/validate_minutes.py <output.docx> --topic <主题> --recorder <纪要人员> --participant <参会人员>
```

多个主题、参会人员和预计议题分别重复使用 `--topics`、`--participant` 和 `--next-topic`。

## 模板结构

保留会议基本信息、会议主要内容、会议核心结论（如有）、跟进事宜及节点、下次交流预计时间、预计议题、交流图片和学习材料提示语。

## 常见错误

- 缺少纪要人员或参会人员仍生成：必须停止。
- 使用固定人员或从逐字稿猜测人员：禁止。
- 使用固定标题：必须根据主题生成。
- 给表头添加底纹：禁止。
- 未验证就声称完成：修复并重新验证。

## 资源

- `assets/materials-department-minutes-template.docx`：基于2026生检交流纪要模板的脱敏资产，保留页眉品牌图形。
- `scripts/build_minutes.py`：生成 DOCX。
- `scripts/validate_minutes.py`：校验关键不变量。

