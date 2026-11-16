import importlib.util
from contextlib import nullcontext
from pathlib import Path
import sys
import unittest
from unittest.mock import patch
from uuid import uuid4
import fitz
from openpyxl import load_workbook

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'tools'))
from create_samples import create_samples

class LayoutReviewTest(unittest.TestCase):
    def test_pdf_review_reports_and_marking(self):
        spec=importlib.util.spec_from_file_location('layout',ROOT/'lee.py')
        app=importlib.util.module_from_spec(spec); spec.loader.exec_module(app)
        folder=create_samples(ROOT/'outputs'/('test-'+uuid4().hex)/'inputs')
        outputs=folder.parent/'results'; outputs.mkdir()
        temporary=folder.parent/'temporary'; temporary.mkdir()
        original=(folder/'synthetic_layout.pdf').read_bytes()
        entries=app.extract_instrument_list(str(folder/'synthetic_layout.pdf'))
        self.assertEqual({e['tag'] for e in entries},{f'9001-FT-{i}' for i in range(101,105)})
        with patch.object(app.tempfile,'TemporaryDirectory',return_value=nullcontext(str(temporary))):
            result=app.run_pipeline(str(folder/'instrument_index.xlsx'),[str(folder/'synthetic_layout.pdf')],str(outputs),lambda _:None)
        self.assertEqual((result['old_count'],result['review_count'],result['jb_count']),(1,1,1))
        self.assertEqual(result['failed_files'],[])
        self.assertEqual(result['no_list_files'],[])
        tags=load_workbook(result['tag_report_path']); self.assertEqual(tags.worksheets[0].cell(2,1).value,'9001-FT-102'); self.assertEqual(tags.worksheets[1].cell(2,1).value,'9001-FT-103'); tags.close()
        jbs=load_workbook(result['jb_report_path']); self.assertEqual(jbs.active.cell(2,1).value,'9001-FT-104'); jbs.close()
        marked=Path(result['marked_dir'])/'synthetic_layout.pdf'
        self.assertTrue(marked.is_file())
        with fitz.open(marked) as pdf: self.assertGreaterEqual(len(pdf[0].get_drawings()),2)
        self.assertEqual(original,(folder/'synthetic_layout.pdf').read_bytes())

if __name__=='__main__': unittest.main()
