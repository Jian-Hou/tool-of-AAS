"""Audit regressions using synthetic data only; no private workbook is required."""
import copy
import hashlib
from pathlib import Path
import sys
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
from conversion import WorkbookData, ValidationError, analyze, build_model
from qa import audit
from schema import default_settings


def children(node):
    return {item['idShort']: item for item in node.get('value', [])}


def assembly(model):
    submodel = next(item for item in model['submodels'] if item['idShort'] == 'AssemblyDefinition')
    return {item['idShort']: item for item in submodel['submodelElements']}


def fixture():
    data = {'components': [], 'pipeline_instances': [], 'elbow_instances': []}
    for index, (label, family) in enumerate([('pipe_1', 'pipeline'), ('pipe_2', 'pipeline'), ('elbow_1', 'elbow')], 1):
        data['components'].append(dict(assembly_id=index, label=label, tag='', type=family,
                                       shape_type='P01', coord='(0,0,0)', placement='(0,0,0,1)|0,0,0,1,1,1'))
        data[family + '_instances'].append(dict(assembly_id=index, label=label, shape_type='P01', params='length_mm=1'))
    for family, count in [('pipeline', 2), ('elbow', 1)]:
        data[family + '_types'] = [dict(shape_type='P01', geometric_description='Synthetic geometry', params_needed='length_mm', count=count)]
    data['joint_instances'] = [
        dict(joint_id=1, joint_type='GroundedJoint', side1_id='Assembly001', side1_sub='', side2_id='', side2_sub=''),
        dict(joint_id=2, joint_type='Fixed', side1_id='pipe_1', side1_sub='End1', side2_id='pipe_2', side2_sub='End2'),
    ]
    data['joint_types'] = [dict(joint_type='GroundedJoint', description='Synthetic root joint', count=1),
                           dict(joint_type='Fixed', description='Synthetic fixed joint', count=1)]
    headers = {sheet: list(rows[0]) for sheet, rows in data.items()}
    for rows in data.values():
        for index, row in enumerate(rows, 2):
            row['_row'] = index
    payload = b'Synthetic in-memory audit fixture; not an Excel file.'
    workbook = WorkbookData(data, headers, 'synthetic.xlsx', hashlib.sha256(payload).hexdigest(), payload)
    settings = default_settings()
    settings.update(manufacturer_name='Synthetic manufacturer', product_designation='Synthetic fixture',
                    article_number='TEST-1', order_code='TEST-1')
    settings['coordinate_system'].update(LengthUnit='mm', OriginDescription='Synthetic origin', Handedness='RightHanded',
                                         AxisConvention='X right, Y forward, Z up', RotationDirection='LocalToAssembly',
                                         BoundingBoxCoordinateSystem='ComponentLocal', Confirmed=True, Source='Synthetic test fixture')
    return workbook, settings


class AuditRegressionTests(unittest.TestCase):
    def setUp(self):
        self.workbook, self.settings = fixture()
        self.model, self.report = build_model(self.workbook, self.settings, strict=True)

    def assert_rejected(self, model):
        result = audit(self.workbook, model, self.settings)
        self.assertFalse(result['data_match'], result)
        self.assertEqual(result['status'], 'failed')

    def test_unmodified_complete_model_passes(self):
        self.assertEqual(self.report['status'], 'complete')
        self.assertTrue(audit(self.workbook, self.model, self.settings)['data_match'])

    def test_joint_count_mismatch_warns_and_blocks_strict_export(self):
        self.workbook.sheets['joint_types'][0]['count'] = 999
        report = analyze(self.workbook, self.settings)[0]
        self.assertEqual(report['status'], 'draft')
        self.assertTrue(any(item['code'] == 'count_mismatch' and item['sheet'] == 'joint_types' for item in report['warnings']))
        self.assertEqual(build_model(self.workbook, self.settings)[1]['status'], 'draft')
        with self.assertRaises(ValidationError):
            build_model(self.workbook, self.settings, strict=True)

    def test_invalid_joint_count_is_reported_before_export(self):
        for count in ('invalid', 1.5, True):
            with self.subTest(count=count):
                self.workbook.sheets['joint_types'][0]['count'] = count
                report = analyze(self.workbook, self.settings)[0]
                self.assertTrue(any(item['code'] == 'invalid_count' and item['sheet'] == 'joint_types' for item in report['errors']))

    def test_optional_joint_count_and_unused_type(self):
        self.workbook.sheets['joint_types'][0]['count'] = ''
        self.workbook.sheets['joint_types'].append(dict(joint_type='Unused', description='Unused type', count=0, _row=4))
        model, _ = build_model(self.workbook, self.settings, strict=True)
        self.assertTrue(audit(self.workbook, model, self.settings)['data_match'])
        self.workbook.sheets['joint_types'][-1]['count'] = 1
        self.assertEqual(analyze(self.workbook, self.settings)[0]['status'], 'draft')

    def test_joint_catalog_values_and_membership_are_audited(self):
        for key, value in [('TypeCode', 'Wrong'), ('GeometricDescription', 'Wrong'), ('SourceCount', '999')]:
            with self.subTest(field=key):
                changed = copy.deepcopy(self.model)
                definition = children(children(assembly(changed)['TypeCatalog'])['JointTypes'])['Fixed']
                children(definition)[key]['value'] = value
                self.assert_rejected(changed)
        changed = copy.deepcopy(self.model)
        definitions = children(assembly(changed)['TypeCatalog'])['JointTypes']['value']
        definitions.pop()
        self.assert_rejected(changed)
        changed = copy.deepcopy(self.model)
        definitions = children(assembly(changed)['TypeCatalog'])['JointTypes']['value']
        extra = copy.deepcopy(definitions[0])
        extra['idShort'] = 'Unexpected'
        definitions.append(extra)
        self.assert_rejected(changed)

    def test_joint_references_must_point_to_the_correct_target(self):
        for joint, field in [('Joint_1', 'Side1Reference'), ('Joint_2', 'Side1Reference'), ('Joint_2', 'Side2Reference')]:
            with self.subTest(joint=joint, field=field):
                changed = copy.deepcopy(self.model)
                nodes = assembly(changed)
                wrong = children(children(nodes['Components'])['elbow_1'])['TechnicalParameters']['value']
                children(children(nodes['Joints'])[joint])[field]['value'] = copy.deepcopy(wrong)
                self.assert_rejected(changed)
        changed = copy.deepcopy(self.model)
        nodes = assembly(changed)
        joint = children(nodes['Joints'])['Joint_1']
        extra = copy.deepcopy(children(joint)['Side1Reference'])
        extra['idShort'] = 'Side2Reference'
        joint['value'].append(extra)
        self.assert_rejected(changed)

    def test_component_references_must_point_to_the_correct_target(self):
        for field, wrong_component in [('TypeDefinition', 'elbow_1'), ('TechnicalParameters', 'pipe_2')]:
            with self.subTest(field=field):
                changed = copy.deepcopy(self.model)
                components = children(assembly(changed)['Components'])
                wrong = children(components[wrong_component])[field]['value']
                children(components['pipe_1'])[field]['value'] = copy.deepcopy(wrong)
                self.assert_rejected(changed)


if __name__ == '__main__':
    unittest.main()
