"""Shared field mapping and configuration contract for CLI and browser."""
from copy import deepcopy

FAMILIES = ('pipeline', 'elbow', 'blackbox', 'tank')
CATALOG_NAMES = dict(zip(FAMILIES, ('PipelineTypes', 'ElbowTypes', 'BlackboxTypes', 'TankTypes')))
FIELDS = []

def _add(sheet, column, target, label, required=True):
    FIELDS.append(dict(excel_sheet=sheet, excel_column=column, aas_path=target, label=label, required=required))

for column, target, label, required in [
    ('assembly_id', 'Identity.ComponentId', 'Component ID', True),
    ('label', 'Identity.Label', 'Component label / reference key', True),
    ('tag', 'Identity.AssetTag', 'Asset tag', False),
    ('type', 'Identity.ComponentType', 'Component family', True),
    ('shape_type', 'Identity.ShapeTypeCode', 'Original type code', True),
    ('coord', 'Position', 'Position (x, y, z)', True),
    ('placement', 'OrientationAndBoundingBox', 'Orientation and bounding box (4 + 6 values)', True),
]:
    _add('components', column, 'AssemblyDefinition.Components[].' + target, label, required)

for family in FAMILIES:
    for col, label in [('assembly_id', 'Component ID'), ('label', 'Component label'), ('shape_type', 'Geometry type'), ('params', 'Parameters')]:
        _add(family + '_instances', col, f'TechnicalData.TechnicalProperties[{family}].{col}', label)
    for col, label, required in [('shape_type', 'Type code', True), ('geometric_description', 'Geometry description', False),
                                 ('params_needed', 'Required parameters', True), ('count', 'Source count', False)]:
        _add(family + '_types', col, f'AssemblyDefinition.TypeCatalog.{CATALOG_NAMES[family]}[].{col}', label, required)

for col, target in [('joint_id', 'JointId'), ('joint_type', 'JointType'), ('side1_id', 'Side1Component'),
                    ('side1_sub', 'Side1Feature'), ('side2_id', 'Side2Component'), ('side2_sub', 'Side2Feature')]:
    _add('joint_instances', col, 'AssemblyDefinition.Joints[].' + target, target)
for col in ('joint_type', 'description', 'count'):
    _add('joint_types', col, 'AssemblyDefinition.TypeCatalog.JointTypes[].' + col, col, col == 'joint_type')

FIELD_BY_PATH = {field['aas_path']: field for field in FIELDS}
DEFAULT_SETTINGS = {
    'name': 'Assembly001', 'asset_id': '',
    'manufacturer_name': '', 'product_designation': '', 'article_number': '', 'order_code': '',
    'coordinate_system': {
        'LengthUnit': '', 'OriginDescription': '', 'Handedness': '', 'AxisConvention': '',
        'QuaternionOrder': 'XYZW', 'RotationDirection': '', 'BoundingBoxCoordinateSystem': '',
        'Confirmed': False, 'Source': '',
    },
    'parameter_units': {},
}

def default_settings():
    return deepcopy(DEFAULT_SETTINGS)

def default_bindings():
    return [{k: field[k] for k in ('aas_path', 'excel_sheet', 'excel_column')} for field in FIELDS]

def mapping_contract():
    return deepcopy(FIELDS)
