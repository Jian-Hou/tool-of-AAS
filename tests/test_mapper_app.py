"""Mapping page API tests."""
import io
import json
from pathlib import Path
import sys
import tempfile
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
from app import create_app
from test_aas_mapper import existing_package, read_back, template_package, workbook


class MapperAppTests(unittest.TestCase):
    def setUp(self):
        tmp = tempfile.TemporaryDirectory(prefix='aas_mapper_app_')
        self.addCleanup(tmp.cleanup)
        self.root = Path(tmp.name)
        workbook(self.root / 'data.xlsx')
        template_package(self.root / 'nameplate.aasx')
        existing_package(self.root / 'pump.aasx')
        self.client = create_app(self.root / 'project', testing=True).test_client()
        self.state = self.client.get('/api/mapper/state').get_json()

    def post(self, url, expect=200, **body):
        response = self.client.post(url, json={**body, 'revision': self.state['revision']}, headers={'X-CSRF-Token': self.state['csrf']})
        self.assertEqual(response.status_code, expect, response.get_json())
        if expect == 200:
            self.state = response.get_json()
        return response.get_json()

    def upload(self, url, name, revision=True, expect=200):
        data = {'file': (io.BytesIO((self.root / name).read_bytes()), name)}
        if revision:
            data['revision'] = str(self.state['revision'])
        response = self.client.post(url, data=data, content_type='multipart/form-data', headers={'X-CSRF-Token': self.state['csrf']})
        self.assertEqual(response.status_code, expect, response.get_json())
        if expect == 200 and revision:
            self.state = response.get_json()
        return response.get_json()

    def node(self, *labels):
        return next(n['target'] for n in self.state['aas']['tree'] if [n['target']['submodel_idShort'], *n['target']['path']] == list(labels))

    def test_new_aas_from_template_with_own_elements(self):
        self.upload('/api/mapper/excel', 'data.xlsx')
        self.assertEqual(self.state['excel']['sheets']['parts']['row_count'], 3)
        library = self.upload('/api/mapper/templates', 'nameplate.aasx', revision=False)
        self.assertEqual(library['templates'][0]['submodels'][0]['idShort'], 'Nameplate')
        self.post('/api/mapper/aas/new', idShort='Pump01', globalAssetId='urn:example:asset:pump01')
        self.post('/api/mapper/aas/edit', action='template', file=library['added'], submodel='urn:template:nameplate')
        self.post('/api/mapper/aas/edit', action='element', parent=self.node('Nameplate'), modelType='Property', idShort='Color', valueType='xs:string')
        self.post('/api/mapper/aas/edit', action='submodel', idShort='Parts')
        rules = [
            {'kind': 'value', 'sheet': 'info', 'column': 'value', 'match': {'column': 'key', 'value': 'manufacturer'}, 'target': self.node('Nameplate', 'ManufacturerName')},
            {'kind': 'value', 'sheet': 'parts', 'column': 'material', 'row': 2, 'target': self.node('Nameplate', 'Color')},
            {'kind': 'rows', 'sheet': 'parts', 'key_column': 'part', 'target': self.node('Parts'),
             'columns': [{'column': 'mass_kg', 'path': ['MassKg'], 'valueType': 'xs:double'}]},
        ]
        self.post('/api/mapper/rules', rules=rules)
        checked = self.post('/api/mapper/check')['result']
        self.assertEqual(checked['report']['errors'], [])
        self.assertTrue(any(n['changed'] and n['value'] == 'ACME GmbH' for n in checked['tree']))
        imported = self.post('/api/mapper/import')['result']
        download = self.client.get(imported['download_url'])
        self.assertEqual(download.status_code, 200)
        self.assertIn('Pump01.aasx', download.headers['Content-Disposition'])
        out = self.root / 'out.aasx'
        out.write_bytes(download.data)
        download.close()
        store, _ = read_back(out)
        nameplate = next(x for x in store if getattr(x, 'id_short', None) == 'Nameplate')
        self.assertEqual((nameplate.get_referable('ManufacturerName').value['en'], nameplate.get_referable('Color').value), ('ACME GmbH', 'steel'))
        mapping = self.client.get('/api/mapper/rules.json')
        self.assertEqual(json.loads(mapping.data)['rules'], self.state['rules'])

    def test_open_existing_package_and_reject_bad_inputs(self):
        self.upload('/api/mapper/excel', 'data.xlsx')
        self.upload('/api/mapper/aas/open', 'nameplate.aasx', expect=422)
        self.upload('/api/mapper/aas/open', 'pump.aasx')
        self.assertEqual(self.state['aas']['shell'], 'urn:example:aas:pump01')
        self.assertEqual(self.state['aas']['format'], 'xml')
        self.post('/api/mapper/check', expect=422)
        rule = {'kind': 'value', 'sheet': 'parts', 'column': 'part', 'row': 2, 'target': self.node('Documentation', 'Revision')}
        self.post('/api/mapper/rules', rules=[rule])
        stale = self.client.post('/api/mapper/check', json={'revision': self.state['revision'] - 1}, headers={'X-CSRF-Token': self.state['csrf']})
        self.assertEqual(stale.status_code, 409)
        self.assertEqual(self.client.post('/api/mapper/check', json={'revision': self.state['revision']}).status_code, 403)
        self.post('/api/mapper/aas/edit', action='element', parent=self.node('Documentation'), modelType='Property', idShort='Revision', valueType='xs:string', expect=422)
        self.post('/api/mapper/rules', rules=[dict(rule, target=dict(rule['target'], path=['Missing']))])
        report = self.post('/api/mapper/import', expect=422)['report']
        self.assertIn('Target not found', report['errors'][0]['message'])


if __name__ == '__main__':
    unittest.main()
