"""Version 3 look: the HTML helpers must escape anything that comes from data.   python -m unittest discover -s tests -v"""
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from app_core import theme  # noqa: E402


class ThemeTests(unittest.TestCase):
    def test_page_header_escapes_markup(self):
        out = theme.page_header_html("<script>alert(1)</script>", "🧾")
        self.assertNotIn("<script>", out)
        self.assertIn("&lt;script&gt;", out)

    def test_kpi_escapes_markup_and_keeps_tone(self):
        out = theme.kpi_html('<img src=x onerror=1>', "5", '"><b>', tone="green")
        self.assertNotIn("<img", out)
        self.assertNotIn("<b>", out)
        self.assertIn(theme.GREEN, out)

    def test_unknown_tone_falls_back_to_blue(self):
        self.assertIn(theme.BLUE, theme.kpi_html("a", "b", tone="nope"))

    def test_css_is_static_and_balanced(self):
        self.assertEqual(theme.CSS.count("{"), theme.CSS.count("}"))
        self.assertNotIn("http", theme.CSS)          # no external resources

    def test_fonts_are_self_hosted_and_declared(self):
        root = Path(__file__).resolve().parent.parent
        cfg = (root / ".streamlit" / "config.toml").read_text(encoding="utf-8")
        self.assertIn("enableStaticServing = true", cfg)
        self.assertNotIn("googleapis", cfg + theme.CSS)
        files = sorted(set(__import__("re").findall(r'url = "app/static/fonts/([^"]+)"', cfg)))
        self.assertEqual(len(files), 3)
        for name in files:
            self.assertTrue((root / "static" / "fonts" / name).is_file(), name)


if __name__ == "__main__":
    unittest.main()
