from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from PIL import Image, ImageDraw

from pro_docx_gen.figure_preflight import FigureAssetGateError, assert_figure_asset_ready
from pro_docx_gen import generate


class FigurePreflightTests(unittest.TestCase):
    def test_docx_renderer_refuses_bad_text_figure_before_insertion(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            image_path = root / "bad-diagram.png"
            output_path = root / "blocked.docx"
            image = Image.new("RGB", (2000, 1000), "white")
            ImageDraw.Draw(image).rectangle((850, 430, 1150, 570), fill="#184D43")
            image.save(image_path)

            semantic = {
                "meta": {"title": "Blocked Figure"},
                "sections": [
                    {
                        "title": "Figure",
                        "level": 1,
                        "content": [
                            {
                                "type": "figure",
                                "path": str(image_path),
                                "caption": "Must not be inserted",
                                "contains_text": True,
                                "min_text_pt": 8,
                                "source_width_inches": 12,
                                "width_inches": 6,
                            }
                        ],
                    }
                ],
            }

            with self.assertRaises(FigureAssetGateError):
                generate(semantic, str(output_path), theme="premium", lang="en", auto_style=False)
            self.assertFalse(output_path.exists())

    def test_gate_blocks_tiny_text_and_excessive_blank_canvas_with_fix_recipe(self):
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / "bad-chart.png"
            image = Image.new("RGB", (2000, 1000), "white")
            draw = ImageDraw.Draw(image)
            draw.rectangle((820, 410, 1180, 590), fill="#184D43")
            image.save(path)

            with self.assertRaises(FigureAssetGateError) as caught:
                assert_figure_asset_ready(
                    path,
                    display_width_inches=6.0,
                    source_width_inches=12.0,
                    declared_min_font_pt=8.0,
                    contains_text=True,
                )

            report = caught.exception.report
            codes = {issue["code"] for issue in report["issues"]}
            self.assertIn("excessive_canvas_whitespace", codes)
            self.assertIn("effective_text_too_small", codes)
            self.assertFalse(report["passed"])
            self.assertGreaterEqual(len(report["remediation_steps"]), 2)
            self.assertIn("crop", " ".join(report["remediation_steps"]).lower())
            self.assertIn("font", " ".join(report["remediation_steps"]).lower())

    def test_gate_passes_dense_legible_asset(self):
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / "good-chart.png"
            image = Image.new("RGB", (1800, 900), "white")
            draw = ImageDraw.Draw(image)
            draw.rectangle((100, 70, 1700, 830), fill="#F5F0E6", outline="#184D43", width=8)
            image.save(path)

            report = assert_figure_asset_ready(
                path,
                display_width_inches=9.0,
                source_width_inches=9.0,
                declared_min_font_pt=12.0,
                contains_text=True,
            )

            self.assertTrue(report["passed"])
            self.assertEqual(report["issues"], [])

    def test_text_figure_without_font_metadata_is_blocked(self):
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / "unknown-text-size.png"
            Image.new("RGB", (1600, 900), "#184D43").save(path)

            with self.assertRaises(FigureAssetGateError) as caught:
                assert_figure_asset_ready(path, display_width_inches=6.0, contains_text=True)

            codes = {issue["code"] for issue in caught.exception.report["issues"]}
            self.assertIn("missing_text_size_metadata", codes)

    def test_opaque_rgba_asset_still_gets_blank_canvas_detection(self):
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / "opaque-rgba.png"
            image = Image.new("RGBA", (1600, 900), (255, 255, 255, 255))
            ImageDraw.Draw(image).rectangle((720, 390, 880, 510), fill=(24, 77, 67, 255))
            image.save(path)

            with self.assertRaises(FigureAssetGateError) as caught:
                assert_figure_asset_ready(path, display_width_inches=6.0)
            codes = {issue["code"] for issue in caught.exception.report["issues"]}
            self.assertIn("excessive_canvas_whitespace", codes)


if __name__ == "__main__":
    unittest.main()
