import importlib.util
import sys
import unittest
from pathlib import Path


MODULE_PATH = Path(__file__).resolve().parents[1] / "3_html" / "mutu24.py"
SPEC = importlib.util.spec_from_file_location("mutu24", MODULE_PATH)
mutu24 = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = mutu24
SPEC.loader.exec_module(mutu24)


class MarketLocationTests(unittest.TestCase):
    def test_accepts_clipped_final_character(self):
        self.assertEqual(
            mutu24.split_market_location("인도ㅣ첸나이무역"),
            ("인도", "첸나이무역관"),
        )

    def test_accepts_regular_office_name(self):
        self.assertEqual(
            mutu24.split_market_location("독일|함부르크무역관"),
            ("독일", "함부르크무역관"),
        )

    def test_rejects_unrelated_delimited_text(self):
        self.assertIsNone(mutu24.split_market_location("합의내용ㅣ경제 협력 강화"))

    def test_best_effort_keeps_five_items_and_other_psd_sections(self):
        def text(index, top, left, name):
            return mutu24.Layer(index, top, left, top + 20, left + 250, name, True)

        layers = [
            text(1, 100, 60, "KOTRA 해외시장뉴스"),
            text(2, 150, 60, "헤드라인 기사"),
            text(3, 220, 70, "인도ㅣ첸나이"),
            text(4, 240, 70, "인도 기사"),
            text(5, 220, 410, "독일ㅣ함부르크무역관"),
            text(6, 240, 410, "독일 기사"),
            text(7, 320, 70, "미국ㅣ뉴욕무역관"),
            text(8, 340, 70, "미국 기사"),
            text(9, 320, 410, "인도네시아ㅣ자카르타무역관"),
            text(10, 340, 410, "인도네시아 기사"),
            text(11, 420, 70, "베트남ㅣ호치민무역관"),
            text(12, 440, 70, "베트남 기사"),
            text(13, 420, 410, "캄보디아ㅣ프놈펜무역관"),
            text(14, 440, 410, "캄보디아 기사"),
            text(15, 600, 60, "KOTRA 사업안내"),
            text(16, 700, 60, "경제외교 소식"),
            text(17, 740, 60, "합의내용ㅣ경제 협력"),
            text(18, 780, 60, "국별정보ㅣ시장 정보"),
            text(19, 30, 200, "2026년 11월 3일"),
        ]
        warnings = []

        headline, market, economy, _, issue_date = mutu24.extract_psd_content(
            layers, True, warnings
        )

        self.assertEqual(headline, "헤드라인 기사")
        self.assertEqual(len(market), 5)
        self.assertEqual([item.map_slot for item in market], [3, 5, 2, 4, 6])
        self.assertEqual(len(economy), 2)
        self.assertEqual(issue_date.isoformat(), "2026-11-03")
        self.assertTrue(any("6개 대신 5개" in warning for warning in warnings))

        template = (MODULE_PATH.parent / "mutu24_template.txt").read_text(encoding="utf-8-sig")
        links = {f"리스트{slot}": f"https://example.com/{slot}" for slot in range(1, 7)}
        output, count = mutu24.replace_market_items(template, market, links)
        self.assertEqual(count, 5)
        self.assertIn("미국 기사", output)
        self.assertIn("캄보디아 기사", output)
        self.assertEqual(output.count('class="newsBox"'), 6)


if __name__ == "__main__":
    unittest.main()
