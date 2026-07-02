import importlib.util
import tempfile
import unittest
from datetime import date
from pathlib import Path

from docx import Document


ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "scripts" / "build_minutes.py"


def load_module():
    spec = importlib.util.spec_from_file_location("build_minutes_update", SCRIPT)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def base_content():
    return {
        "recorder": "张三",
        "participants": ["张三", "李四"],
        "topics": ["主题甲", "主题乙", "主题丙", "主题丁"],
        "sections": [{"title": "主题甲", "paragraphs": ["会议内容。"]}],
        "conclusions": ["会议结论。"],
        "actions": [],
        "review_items": ["人员", "数值", "日期"],
    }


class TemplateUpdateTests(unittest.TestCase):
    def build_text(self, content):
        module = load_module()
        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / "minutes.docx"
            module.build_minutes(content, output, meeting_date=date(2026, 7, 2))
            return "\n".join(p.text for p in Document(output).paragraphs)

    def test_uses_dynamic_title_and_specified_people(self):
        text = self.build_text(base_content())
        self.assertIn("关于主题甲、主题乙、主题丙等议题纪要", text)
        self.assertIn("纪要人员：张三", text)
        self.assertIn("参与人员：张三、李四", text)

    def test_rejects_missing_recorder(self):
        content = base_content()
        content["recorder"] = ""
        with self.assertRaisesRegex(ValueError, "纪要人员不能为空"):
            self.build_text(content)

    def test_rejects_missing_participants(self):
        content = base_content()
        content["participants"] = []
        with self.assertRaisesRegex(ValueError, "参会人员不能为空"):
            self.build_text(content)

    def test_preserves_new_template_sections_and_optional_next_meeting(self):
        content = base_content()
        content["next_meeting_date"] = "2026年7月9日"
        content["next_topics"] = ["主题戊", "主题己"]
        text = self.build_text(content)
        for required in (
            "三、会议核心结论（如有）",
            "四、跟进事宜及节点",
            "下次交流预计时间：2026年7月9日",
            "预计议题：主题戊、主题己",
            "交流图片",
            "如涉及比较好的学习材料和要点可分享给大家共同学习",
        ):
            self.assertIn(required, text)

    def test_next_meeting_fields_default_to_pending_confirmation(self):
        text = self.build_text(base_content())
        self.assertIn("下次交流预计时间：待确认", text)
        self.assertIn("预计议题：待确认", text)


if __name__ == "__main__":
    unittest.main()

