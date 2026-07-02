---
name: meeting-minutes
description: Use when a user provides a meeting transcript TXT and meeting topics and wants a standardized materials-department meeting-minutes DOCX, especially for noisy Chinese automatic transcripts that require conservative fact checking.
---

# 材料部会议纪要

## 概述

将用户提供的会议逐字稿 TXT 和会议主题整理为正式 DOCX。语义提炼由 Codex 完成，`scripts/build_minutes.py` 负责确定性排版，`scripts/validate_minutes.py` 负责结构校验。

**REQUIRED SUB-SKILL:** 使用 `documents` 完成 DOCX 生成后的渲染与逐页检查。

## 固定规则

- 纪要人员始终写“郭源杰”，不得从逐字稿、模板或用户内容中替换。
- 会议日期始终取运行时 `Asia/Shanghai` 的当天日期，不从文件名或逐字稿推断。
- 默认文件名为 `YYYYMMDD材料部组会会议纪要.docx`。
- 跟进表必须使用模板的五列、边框和列宽：序号、跟进事宜、负责人、截止日期、备注。
- 表格不得使用填充颜色，包括表头。
- 不确定的人名、产品名、数值、负责人和日期标注“待核验”或“待确认”，不得编造。
- 最终纪要列出 3—5 项最需要人工复核的信息。

## 工作流

1. 确认用户同时提供可读取的会议逐字稿 `.txt` 和一条或多条会议主题。
2. 读取完整逐字稿，不要只读取关键词命中的片段。
3. 结合会议主题提取会议地点、参与人员、分主题进展、核心结论、行动项、负责人、截止日期和复核项。
4. 保守修正常见同音转写。只有上下文足够明确时才规范专业名称；否则保留“待核验”。
5. 将整理结果写为 UTF-8 JSON，字段遵循下方结构。不要在 JSON 中控制纪要人员或会议日期。
6. 调用 `scripts/build_minutes.py` 生成 DOCX。
7. 调用 `scripts/validate_minutes.py`，检查日期、纪要人员、主题覆盖、表格列宽和无填充颜色。
8. 使用 `documents` Skill 渲染 DOCX 并检查全部页面。若 LibreOffice/Word 渲染不可用，进行结构检查并在交付时明确披露。
9. 只向用户交付最终 DOCX；不要交付中间 JSON、渲染图片或临时文件，除非用户明确要求。

## 内容 JSON

```json
{
  "location": "实验楼B515会议室",
  "participants": ["姓名1", "姓名2"],
  "topics": ["同位素材料", "叔丁胺合成"],
  "sections": [{"title": "同位素材料", "paragraphs": ["完整、审慎的纪要段落。"]}],
  "conclusions": ["可执行的核心结论。"],
  "actions": [{"item": "跟进事项", "owner": "负责人或待确认", "deadline": "截止日期或待确认", "notes": "必要备注"}],
  "review_items": ["参会人员名单", "产品名称", "关键数值"],
  "source_note": "会议逐字稿《文件名.txt》"
}
```

## 执行命令

先通过工作区依赖加载器获取捆绑 Python 路径，再执行：

```text
<python> scripts/build_minutes.py --content <content.json> --transcript <transcript.txt> --topics <主题1> --topics <主题2> --output <output.docx>
<python> scripts/validate_minutes.py <output.docx> --topic <主题1> --topic <主题2>
```

若省略 `--output`，文件输出到逐字稿所在目录。

## 质量要求

- “会议主要内容”围绕用户指定主题组织，不机械按发言顺序抄写。
- 删除寒暄、拍照、设备操作等与主题无关的口语噪声。
- 区分“已完成”“计划”“建议”“待确认”，不得把讨论意见写成既定事实。
- 保留有决策价值的实验条件、纯度、丰度、产率和节点；数值含混时标记待核验。
- 资料来源注明为用户提供的逐字稿；未联网时明确写“未进行外部联网检索”。

## 常见错误

- 使用逐字稿日期代替当天日期：禁止，日期由脚本固定。
- 从输入覆盖纪要人员：禁止，纪要人员固定为郭源杰。
- 给表头加蓝色或灰色底纹：禁止，模板表格不使用任何填充色。
- 为了让表格整齐而缩小到难读字号：保持模板字体和可自动扩展的行高。
- 逐字稿存在同音词却直接纠正为某个产品：证据不足时写“待核验”。
- 未运行校验或渲染就声称完成：修复问题并重新验证后再交付。

## 资源

- `assets/materials-department-minutes-template.docx`：脱敏空白模板，不含历史会议内容或图片。
- `scripts/build_minutes.py`：生成 DOCX，固定日期、纪要人员和表格格式。
- `scripts/validate_minutes.py`：检查输出文件的关键不变量。

