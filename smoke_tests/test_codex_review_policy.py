import unittest
import uuid
from pathlib import Path

from PIL import Image

import agent_docx


ROOT = Path(__file__).resolve().parents[1]


class CodexReviewPolicyTests(unittest.TestCase):
    def test_sampling_selects_one_page_from_each_three_page_group(self):
        previews = [(page, f"page-{page:02d}.png") for page in range(1, 14)]

        sampled = agent_docx._select_sample_page_previews(previews)

        self.assertEqual([page for page, _ in sampled], [1, 4, 7, 10, 13])

    def test_multiple_samples_create_four_up_pdf_and_png_sheets(self):
        tmp_path = ROOT / "_work" / "test-codex-review-policy" / uuid.uuid4().hex
        tmp_path.mkdir(parents=True, exist_ok=False)
        sampled = []
        for page in (1, 4, 7, 10, 13):
            preview = tmp_path / f"page-{page:02d}.png"
            Image.new("RGB", (600, 840), (245, 245, 245)).save(preview)
            sampled.append((page, str(preview)))

        bundle = agent_docx._build_four_up_review_bundle(sampled, tmp_path, "sample")

        self.assertTrue(Path(bundle["pdf"]).exists())
        self.assertTrue(Path(bundle["pdf"]).read_bytes().startswith(b"%PDF"))
        self.assertEqual(len(bundle["sheet_pngs"]), 2)
        self.assertEqual(bundle["pages_per_sheet"], 4)
        self.assertEqual(bundle["sampled_page_numbers"], [1, 4, 7, 10, 13])

    def test_skill_root_documents_conditional_render_rules(self):
        skill_text = (ROOT / "SKILL.md").read_text(encoding="utf-8")
        nested_text = (ROOT / "pro_docx_gen" / "SKILL.md").read_text(encoding="utf-8")
        readme_text = (ROOT / "README.md").read_text(encoding="utf-8")

        for text in (skill_text, nested_text, readme_text):
            self.assertIn("one page from every consecutive three-page group", text)
            self.assertIn("every figure page", text)
            self.assertIn("structural", text.lower())
            self.assertIn("cell border", text.lower())

    def test_submission_manifest_marks_full_figure_review(self):
        manifest = agent_docx._submission_manifest(
            target=Path("report.docx"),
            quality_report=Path("quality.json"),
            document_page_count=5,
            sampled_pages=[(page, f"p{page}.png") for page in range(1, 6)],
            review_bundle={
                "sheet_pngs": ["sheet.png"],
                "pdf": "review.pdf",
                "sampled_page_numbers": list(range(1, 6)),
                "pages_per_sheet": 4,
            },
            full_figure_review=True,
        )
        self.assertEqual(manifest["render_scope"], "full_document_for_figure_review")
        self.assertEqual(manifest["visual_review_sampling"]["minimum_total"], 5)


if __name__ == "__main__":
    unittest.main()
