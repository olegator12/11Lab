import csv
import json
import subprocess
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


class DashboardBuildTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        subprocess.run([sys.executable, str(ROOT / "scripts" / "build_dashboard.py")], cwd=ROOT, check=True)

    def test_input_csv_row_counts(self):
        expected = {
            "character-deaths.csv": 917,
            "character-predictions.csv": 1946,
            "battles.csv": 38,
        }
        for filename, count in expected.items():
            with self.subTest(filename=filename):
                with (ROOT / "data" / "raw" / filename).open(encoding="utf-8-sig", newline="") as stream:
                    rows = list(csv.DictReader(stream))
                self.assertEqual(len(rows), count)

    def test_dashboard_is_standalone_and_embeds_data(self):
        html = (ROOT / "index.html").read_text(encoding="utf-8")
        self.assertIn('<script id="dataset" type="application/json">', html)
        self.assertNotIn("__DATA_JSON__", html)
        self.assertNotIn('<script src=', html.lower())
        self.assertNotIn('<link rel="stylesheet" href=', html.lower())
        self.assertIn("const raw=JSON.parse", html)
        self.assertIn("function drawDeaths()", html)
        self.assertIn("function drawPredictions()", html)
        self.assertIn("function drawBattles()", html)

    def test_quality_report(self):
        report = json.loads((ROOT / "data" / "processed" / "data_quality.json").read_text(encoding="utf-8"))
        self.assertEqual(report["rows_after_exact_duplicate_removal"]["deaths"], 917)
        self.assertEqual(report["rows_after_exact_duplicate_removal"]["predictions"], 1946)
        self.assertEqual(report["rows_after_exact_duplicate_removal"]["battles"], 38)


if __name__ == "__main__":
    unittest.main()
