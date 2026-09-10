import tempfile
import unittest
from pathlib import Path

import nmdc_profiler as profiler


class CrossPlatformOutputTests(unittest.TestCase):
    def test_01_gitattributes_pins_text_outputs_to_lf(self):
        root = Path(__file__).resolve().parents[1]
        text = (root / ".gitattributes").read_text(encoding="utf-8")
        self.assertIn("*.csv text eol=lf", text)
        self.assertIn("*.json text eol=lf", text)
        self.assertIn("*.md text eol=lf", text)

    def test_02_csv_writer_emits_lf_only(self):
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "sample.csv"
            profiler.write_csv(
                path,
                [{"a": "line1\nline2", "b": "value"}],
                ["a", "b"],
            )
            data = path.read_bytes()
            self.assertIn(b"\n", data)
            self.assertNotIn(b"\r\n", data)


if __name__ == "__main__":
    unittest.main()
