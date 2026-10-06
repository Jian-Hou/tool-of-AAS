"""Local project storage must not grow with every upload and export."""
import io
from pathlib import Path
import sys
import tempfile
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
from app import KEEP_OUTPUTS, create_app
from test_workbook_export import minimal_workbook


class AppStorageTests(unittest.TestCase):
    def setUp(self):
        tmp = tempfile.TemporaryDirectory(prefix='aas_app_')
        self.addCleanup(tmp.cleanup)
        root = Path(tmp.name)
        self.source = minimal_workbook(root / 'minimal.xlsx').read_bytes()
        self.client = create_app(root / 'project', testing=True).test_client()
        self.state = self.client.get('/api/state').get_json()
        self.folder = next(path for path in (root / 'project' / '.state').iterdir() if path.is_dir())

    def upload(self):
        response = self.client.post('/api/excel/load', content_type='multipart/form-data',
                                    data={'file': (io.BytesIO(self.source), 'minimal.xlsx'), 'revision': str(self.state['revision'])},
                                    headers={'X-CSRF-Token': self.state['csrf']})
        self.assertEqual(response.status_code, 200, response.get_json())
        self.state = response.get_json()

    def export(self):
        response = self.client.post('/api/generate-aasx', json={'revision': self.state['revision'], 'strict': False},
                                    headers={'X-CSRF-Token': self.state['csrf']})
        self.assertEqual(response.status_code, 200, response.get_json())
        return response.get_json()['download_url']

    def test_reupload_replaces_previous_input(self):
        for _ in range(3):
            self.upload()
        self.assertEqual(len(list(self.folder.glob('*.xlsx'))), 1)
        self.assertEqual(self.state['report']['status'], 'draft')

    def test_old_exports_are_pruned(self):
        self.upload()
        urls = [self.export() for _ in range(KEEP_OUTPUTS + 2)]
        self.assertEqual(len(list((self.folder / 'outputs').glob('*.aasx'))), KEEP_OUTPUTS)
        self.assertEqual(len(list((self.folder / 'outputs').glob('*.json'))), KEEP_OUTPUTS)
        response = self.client.get(urls[-1])
        self.assertEqual(response.status_code, 200)
        response.close()


if __name__ == '__main__':
    unittest.main()
