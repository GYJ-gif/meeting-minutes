import unittest
from pathlib import Path

SKILL = Path(__file__).resolve().parents[1] / "SKILL.md"

class SkillInstructionTests(unittest.TestCase):
    def test_required_workflow_is_explicit(self):
        text = SKILL.read_text(encoding="utf-8")
        self.assertNotIn("TODO", text)
        for required in ("逐字稿", "会议主题", "纪要人员", "参会人员", "缺少任一项", "Asia/Shanghai", "不得使用填充颜色", "build_minutes.py", "validate_minutes.py", "3—5"):
            self.assertIn(required, text)

    def test_template_fidelity_rules_are_explicit(self):
        text = SKILL.read_text(encoding="utf-8")
        for required in ("保留模板骨架", "#2C569A", "五、交流图片", "复制模板数据行"):
            self.assertIn(required, text)

if __name__ == "__main__":
    unittest.main()
