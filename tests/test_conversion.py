"""Integration tests for the conversion core (run from the project folder)."""

from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

from docx import Document
from docx.shared import Pt
import pymupdf

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from conversion import (  # noqa: E402
    ConversionError,
    convert_one,
    find_word_engine,
    prepare_logging,
)


class ConversionCoreTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        prepare_logging()
        cls.engine = find_word_engine()
        if not cls.engine:
            raise unittest.SkipTest("No Word -> PDF engine available")
        cls.temp = tempfile.TemporaryDirectory(prefix="wenshu-test-")
        cls.root = Path(cls.temp.name)
        cls.source = cls.root / "测试 文档.docx"
        cls._make_docx(cls.source)

    @classmethod
    def tearDownClass(cls) -> None:
        cls.temp.cleanup()

    @staticmethod
    def _make_docx(path: Path) -> None:
        document = Document()
        document.add_heading("转换测试文档", level=1)
        document.add_paragraph("这是中文测试段落，包含 English text 和数字 123。")
        document.add_paragraph("第一条项目", style="List Bullet")
        document.add_paragraph("第二条项目", style="List Bullet")
        table = document.add_table(rows=2, cols=2)
        table.style = "Table Grid"
        table.cell(0, 0).text = "名称"
        table.cell(0, 1).text = "结果"
        table.cell(1, 0).text = "测试"
        table.cell(1, 1).text = "通过"
        for style in document.styles:
            if style.type == 1:
                style.font.name = "Microsoft YaHei"
                style.font.size = Pt(10.5)
        document.save(path)

    def test_word_to_pdf_and_unique_names(self) -> None:
        output = self.root / "out-pdf"
        first = convert_one(
            mode="word", source=str(self.source), output_dir=str(output), engine=self.engine
        )
        second = convert_one(
            mode="word", source=str(self.source), output_dir=str(output), engine=self.engine
        )
        self.assertTrue(first["ok"])
        self.assertTrue(second["ok"])
        self.assertTrue(Path(first["output"]).is_file())
        self.assertNotEqual(first["output"], second["output"])
        self.assertTrue(Path(first["output"]).read_bytes().startswith(b"%PDF-"))
        self.assertIn("(1)", Path(second["output"]).name)

    def test_pdf_to_word(self) -> None:
        pdf_dir = self.root / "roundtrip-pdf"
        pdf_result = convert_one(
            mode="word", source=str(self.source), output_dir=str(pdf_dir), engine=self.engine
        )
        docx_dir = self.root / "roundtrip-docx"
        result = convert_one(
            mode="pdf", source=pdf_result["output"], output_dir=str(docx_dir)
        )
        self.assertTrue(result["ok"])
        output = Path(result["output"])
        self.assertTrue(output.is_file())
        document = Document(output)
        text = "\n".join(paragraph.text for paragraph in document.paragraphs)
        self.assertIn("转换测试文档", text)
        self.assertIn("中文测试段落", text)
        self.assertTrue(any("第一条项目" in p.text for p in document.paragraphs))
        self.assertTrue(any("第二条项目" in p.text for p in document.paragraphs))
        table_text = "\n".join(cell.text for table in document.tables for row in table.rows for cell in row.cells)
        self.assertIn("通过", table_text)
        self.assertGreater(result["pages"], 0)

    def test_pdf_layout_mode(self) -> None:
        pdf_dir = self.root / "layout-pdf"
        pdf_result = convert_one(
            mode="word", source=str(self.source), output_dir=str(pdf_dir), engine=self.engine
        )
        output = self.root / "layout-docx"
        result = convert_one(
            mode="pdf",
            source=pdf_result["output"],
            output_dir=str(output),
            pdf_mode="layout",
        )
        self.assertTrue(result["ok"])
        document = Document(result["output"])
        self.assertGreater(len(document.paragraphs), 0)
        self.assertGreater(len(document.tables), 0)

    def test_encrypted_pdf_is_rejected(self) -> None:
        source = pymupdf.open()
        page = source.new_page()
        page.insert_text((72, 72), "secret")
        encrypted = self.root / "encrypted.pdf"
        source.save(
            encrypted,
            encryption=pymupdf.PDF_ENCRYPT_AES_256,
            owner_pw="owner",
            user_pw="user",
        )
        source.close()
        with self.assertRaises(ConversionError) as context:
            convert_one(
                mode="pdf", source=str(encrypted), output_dir=str(self.root / "encrypted-out")
            )
        self.assertIn("加密", str(context.exception))


if __name__ == "__main__":
    unittest.main(verbosity=2)
