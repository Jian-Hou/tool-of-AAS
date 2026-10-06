"""Checks against official IDTA submodel templates in aas_templates/.

Template files are local data and are not committed, so each test is skipped when its file is absent.
See aas_templates/README.md for the source of the files."""
import datetime
from pathlib import Path
import sys
import tempfile
import unittest

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
import aas_mapper as m  # noqa: E402
from conversion import ValidationError  # noqa: E402
from test_aas_mapper import read_back  # noqa: E402

LIBRARY = Path(__file__).resolve().parents[1] / 'aas_templates'
NAMEPLATE = LIBRARY / 'IDTA_02006-3-0-2_DigitalNameplate.aasx'
TECHNICAL_DATA = LIBRARY / 'IDTA_02003-2-0-2_TechnicalData.aasx'
BOM = LIBRARY / 'IDTA_02011-1-1-2_HierarchicalStructures.aasx'
CONTACT_V31 = LIBRARY / 'IDTA_02002-1-0-2_ContactInformation_V3-1.aasx'


def needs(path):
    return unittest.skipUnless(path.exists(), f'{path.name} is not in the template library')


class OfficialTemplateTests(unittest.TestCase):
    def setUp(self):
        tmp = tempfile.TemporaryDirectory(prefix='aas_official_')
        self.addCleanup(tmp.cleanup)
        self.dir = Path(tmp.name)

    def instance(self, path):
        template = m.load_source(path)
        env = m.new_environment('Asset01', 'urn:example:asset:01')
        shell_id = env['assetAdministrationShells'][0]['id']
        for submodel in template.env['submodels']:
            m.add_template_submodel(env, shell_id, template.env, submodel['id'])
        return env, shell_id

    def target(self, env, shell_id, *path):
        submodel = m.submodels_of(env, shell_id)[0]
        return {'submodel': submodel['id'], 'submodel_idShort': submodel['idShort'], 'path': list(path)}

    def workbook(self, sheets):
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        for name, rows in sheets.items():
            sheet = wb.create_sheet(name)
            for row in rows:
                sheet.append(row)
        wb.save(self.dir / 'data.xlsx')
        return m.read_tables(self.dir / 'data.xlsx')

    def export(self, final):
        self.assertEqual(m.verify(final), [])
        out = self.dir / 'out.aasx'
        m.write_package(out, final)
        store, _ = read_back(out)
        return next(x for x in store if hasattr(x, 'submodel_element'))

    def test_every_library_template_becomes_a_valid_instance(self):
        paths = sorted(LIBRARY.glob('*.aasx'))
        if not paths:
            self.skipTest('the template library is empty')
        for path in paths:
            with self.subTest(template=path.name):
                try:
                    env, _ = self.instance(path)
                except ValidationError as exc:
                    self.assertIn('V3.1', str(exc))
                    continue
                for drop in (True, False):
                    self.export(m.finalize(env, drop)[0])

    @needs(CONTACT_V31)
    def test_v3_1_files_get_a_clear_message(self):
        with self.assertRaisesRegex(ValidationError, 'V3.1 or later; the tool supports V3.0'):
            m.load_source(CONTACT_V31)

    @needs(NAMEPLATE)
    def test_digital_nameplate(self):
        env, shell_id = self.instance(NAMEPLATE)
        t = lambda *p: self.target(env, shell_id, *p)
        tables = self.workbook({
            'info': [('key', 'value'), ('uri', 'https://example.com/products/fm-1'), ('maker', 'Example GmbH'),
                     ('designation', 'Flow meter'), ('order', 'FM-1-ORDER'), ('built', datetime.datetime(2026, 5, 4))],
            'markings': [('name', 'issued'), ('CE', datetime.datetime(2025, 1, 1)), ('UKCA', datetime.datetime(2025, 6, 1))]})
        value = lambda key, *path, **kw: {'kind': 'value', 'sheet': 'info', 'column': 'value', 'match': {'column': 'key', 'value': key}, 'target': t(*path), **kw}
        rules = [value('uri', 'URIOfTheProduct'), value('maker', 'ManufacturerName', language='de'),
                 value('designation', 'ManufacturerProductDesignation'), value('order', 'OrderCodeOfManufacturer'), value('built', 'DateOfManufacture'),
                 {'kind': 'rows', 'sheet': 'markings', 'prototype': 0, 'target': t('Markings'),
                  'columns': [{'column': 'name', 'path': ['MarkingName']}, {'column': 'issued', 'path': ['IssueDate']}]}]
        result, report = m.apply_rules(env, shell_id, tables, rules)
        self.assertEqual(report['errors'], [])
        final, info = m.finalize(result)
        self.assertGreater(info['removed'], 10)
        self.assertEqual(info['warnings'], ['Mandatory template element Nameplate/AddressInformation (One) is empty.',
                                            'Mandatory template element Nameplate/Markings/[*]/MarkingFile (One) has no value in 2 list entries.'])
        nameplate = self.export(final)
        self.assertEqual(nameplate.get_referable('ManufacturerName').value['de'], 'Example GmbH')
        self.assertEqual(str(nameplate.get_referable('DateOfManufacture').value), '2026-05-04')
        self.assertEqual([x.get_referable('MarkingName').value for x in nameplate.get_referable('Markings').value], ['CE', 'UKCA'])
        self.assertNotIn('HardwareVersion', {e.id_short for e in nameplate.submodel_element})

    @needs(TECHNICAL_DATA)
    def test_technical_data(self):
        env, shell_id = self.instance(TECHNICAL_DATA)
        t = lambda *p: self.target(env, shell_id, *p)
        tables = self.workbook({
            'general': [('field', 'value'), ('ManufacturerName', 'Example GmbH'), ('ManufacturerProductDesignation', 'Pipe assembly'),
                        ('ManufacturerArticleNumber', 'A-100'), ('ManufacturerOrderCode', 'O-100')],
            'classes': [('system', 'id', 'name'), ('ECLASS', '36-41-01-01', 'Pipe'), ('ETIM', 'EC000123', 'Elbow')]})
        rules = [{'kind': 'value', 'sheet': 'general', 'column': 'value', 'match': {'column': 'field', 'value': field},
                  'target': t('GeneralInformation', field)} for field in ('ManufacturerName', 'ManufacturerProductDesignation',
                                                                         'ManufacturerArticleNumber', 'ManufacturerOrderCode')]
        rules.append({'kind': 'rows', 'sheet': 'classes', 'prototype': 0, 'target': t('ProductClassifications'),
                      'columns': [{'column': 'system', 'path': ['ClassificationSystem']}, {'column': 'id', 'path': ['ProductClassId']},
                                  {'column': 'name', 'path': ['ProductClassCodedName']}]})
        result, report = m.apply_rules(env, shell_id, tables, rules)
        self.assertEqual(report['errors'], [])
        final, info = m.finalize(result)
        self.assertEqual(info['warnings'], [])
        data = self.export(final)
        self.assertEqual([e.id_short for e in data.submodel_element], ['GeneralInformation', 'ProductClassifications'])
        self.assertEqual([x.get_referable('ProductClassId').value for x in data.get_referable('ProductClassifications').value], ['36-41-01-01', 'EC000123'])

    @needs(BOM)
    def test_bill_of_materials_from_rows(self):
        env, shell_id = self.instance(BOM)
        t = lambda *p: self.target(env, shell_id, *p)
        tables = self.workbook({'bom': [('part', 'asset', 'qty'), ('pipe-1', 'urn:example:asset:pipe-1', 2),
                                        ('elbow 2', 'urn:example:asset:elbow-2', 1), ('tank_3', 'urn:example:asset:tank-3', 1)],
                                'meta': [('key', 'value'), ('root', 'urn:example:asset:01'), ('archetype', 'OneDown')]})
        rules = [{'kind': 'value', 'sheet': 'meta', 'column': 'value', 'row': 2, 'target': t('EntryNode')},
                 {'kind': 'value', 'sheet': 'meta', 'column': 'value', 'row': 3, 'target': t('ArcheType')},
                 {'kind': 'rows', 'sheet': 'bom', 'key_column': 'part', 'prototype': 'Node', 'target': t('EntryNode'),
                  'columns': [{'column': 'asset', 'path': ['#globalAssetId']}, {'column': 'qty', 'path': ['BulkCount']}]}]
        result, report = m.apply_rules(env, shell_id, tables, rules)
        self.assertEqual(report['errors'], [])
        final, info = m.finalize(result)
        self.assertEqual(info['warnings'], [])
        bom = self.export(final)
        entry = bom.get_referable('EntryNode')
        self.assertEqual(entry.global_asset_id, 'urn:example:asset:01')
        self.assertEqual(sorted((e.id_short, e.global_asset_id, e.get_referable('BulkCount').value) for e in entry.statement),
                         sorted([(m.safe_id('pipe-1'), 'urn:example:asset:pipe-1', 2), (m.safe_id('elbow 2'), 'urn:example:asset:elbow-2', 1),
                                 ('tank_3', 'urn:example:asset:tank-3', 1)]))


if __name__ == '__main__':
    unittest.main()
