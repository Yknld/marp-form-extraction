import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from types import SimpleNamespace
import sort_forms


class SortTests(unittest.TestCase):
    def pdf(self, pages):
        context = unittest.mock.MagicMock()
        context.__enter__.return_value.pages = pages
        return context

    def page(self, text, images=()):
        return SimpleNamespace(extract_text=lambda: text, images=images,
                               curves=[], width=100, height=100)

    def test_pdf_text_and_mixed_scan(self):
        text = self.page('This is a fictional response with enough readable text.')
        scan = self.page('Page 2', [dict(x0=0, x1=100, top=0, bottom=100)])
        with patch('pdfplumber.open', return_value=self.pdf([text])):
            self.assertEqual(sort_forms.classify(Path('x.pdf'))[0], 'Text PDFs')
        with patch('pdfplumber.open', return_value=self.pdf([text, scan])):
            self.assertEqual(sort_forms.classify(Path('x.pdf'))[0], 'Needs OCR')
        with patch('pdfplumber.open', side_effect=ValueError):
            self.assertEqual(sort_forms.classify(Path('x.pdf'))[0], 'Other or unreadable files')

    def test_copies_preserve_paths_and_never_overwrite(self):
        with tempfile.TemporaryDirectory() as td:
            source = Path(td)/'forms'; source.mkdir()
            output = Path(td)/'sorted'
            for folder in ['a', 'b']:
                (source/folder).mkdir()
                (source/folder/'001.docx').write_bytes(b'original')
            (source/'002.png').write_bytes(b'image')
            (source/'003.zip').write_bytes(b'archive')
            sort_forms.run(source, output)
            copied = output/'Word documents/a/001.docx'
            self.assertEqual(copied.read_bytes(), b'original')
            self.assertTrue((output/'Word documents/b/001.docx').exists())
            self.assertTrue((output/'Needs OCR/002.png').exists())
            copied.write_bytes(b'manual edit')
            sort_forms.run(source, output)
            self.assertEqual(copied.read_bytes(), b'manual edit')
            self.assertEqual((source/'a/001.docx').read_bytes(), b'original')
            with self.assertRaises(ValueError):
                sort_forms.run(source, source/'sorted')
