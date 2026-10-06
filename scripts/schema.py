"""Field mapping."""
from copy import deepcopy

FAMILIES = ('pipeline', 'elbow', 'blackbox', 'tank')
CATALOG_NAMES = dict(zip(FAMILIES, ('PipelineTypes', 'ElbowTypes', 'BlackboxTypes', 'TankTypes')))
FIELDS = []

# aas_path: stable key for saved settings.
# output: where the field goes in the AAS.
def _add(sheet, column, target, label, required=True, output=(), note=''):
    FIELDS.append(dict(excel_sheet=sheet, excel_column=column, aas_path=target, label=label, required=required,
                       output=list(output), note=note))

COMPONENT = 'AssemblyDefinition/Components/<label>/'
for column, target, label, required, output, note in [
    ('assembly_id', 'Identity.ComponentId', 'Component ID', True, [COMPONENT + 'Identity/ComponentId'],
     'Must match assembly_id in the instance sheet.'),
    ('label', 'Identity.Label', 'Component label / reference key', True,
     [COMPONENT + 'Identity/Label', 'AssemblyDefinition/Components/<label> (collection name)', 'TechnicalData/TechnicalProperties/<label> (collection name)'],
     'Links the instance sheets and joints. Characters not allowed in names are replaced and a hash is appended.'),
    ('tag', 'Identity.AssetTag', 'Asset tag', False, [COMPONENT + 'Identity/AssetTag'], ''),
    ('type', 'Identity.ComponentType', 'Component family', True,
     [COMPONENT + 'Identity/ComponentType', 'TechnicalData/ProductClassifications/<family>_<code>/ProductClassId (family part)'],
     'Selects the instance sheet and type catalog.'),
    ('shape_type', 'Identity.ShapeTypeCode', 'Original type code', True,
     [COMPONENT + 'Identity/ShapeTypeCode', 'TechnicalData/ProductClassifications/<family>_<code>/ProductClassId (code part)'], ''),
    ('coord', 'Position', 'Position (x, y, z)', True, [COMPONENT + 'Position (X, Y, Z)', COMPONENT + 'SourceCoord (original text)'], ''),
    ('placement', 'OrientationAndBoundingBox', 'Orientation and bounding box (4 + 6 values)', True,
     [COMPONENT + 'Orientation (Qx, Qy, Qz, Qw)', COMPONENT + 'BoundingBox (XMin to ZMax)', COMPONENT + 'SourcePlacement (original text)'], ''),
]:
    _add('components', column, 'AssemblyDefinition.Components[].' + target, label, required, output, note)

for family in FAMILIES:
    catalog = f'AssemblyDefinition/TypeCatalog/{CATALOG_NAMES[family]}/<code>/'
    for col, label, output, note in [
        ('assembly_id', 'Component ID', [], 'Not written to the AAS; must match the component\'s assembly_id.'),
        ('label', 'Component label', ['TechnicalData/TechnicalProperties/<label> (collection name)'], 'Links this record to the component with the same label.'),
        ('shape_type', 'Geometry type', [COMPONENT + 'Identity/GeometryTypeCode', COMPONENT + f'TypeDefinition (reference to {catalog.rstrip("/")})'], ''),
        ('params', 'Parameters', ['TechnicalData/TechnicalProperties/<label>/<parameter> (one property per value)', COMPONENT + 'SourceParameters (original text)'],
         'Each parameter also gets a concept description with its unit.'),
    ]:
        _add(family + '_instances', col, f'TechnicalData.TechnicalProperties[{family}].{col}', label, True, output, note)
    for col, label, required, output, note in [
        ('shape_type', 'Type code', True, [catalog + 'TypeCode'], ''),
        ('geometric_description', 'Geometry description', False, [catalog + 'GeometricDescription'], ''),
        ('params_needed', 'Required parameters', True, [catalog + 'RequiredParameters'], 'Also used to report missing parameters.'),
        ('count', 'Source count', False, [catalog + 'SourceCount'], 'Compared with the number of instances.'),
    ]:
        _add(family + '_types', col, f'AssemblyDefinition.TypeCatalog.{CATALOG_NAMES[family]}[].{col}', label, required, output, note)

JOINT = 'AssemblyDefinition/Joints/Joint_<id>/'
for col, target, output in [
    ('joint_id', 'JointId', [JOINT + 'JointId']), ('joint_type', 'JointType', [JOINT + 'JointType']),
    ('side1_id', 'Side1Component', [JOINT + 'Side1Component', JOINT + 'Side1Reference (reference to the component or the AAS)']),
    ('side1_sub', 'Side1Feature', [JOINT + 'Side1Feature']),
    ('side2_id', 'Side2Component', [JOINT + 'Side2Component', JOINT + 'Side2Reference (reference to the component or the AAS)']),
    ('side2_sub', 'Side2Feature', [JOINT + 'Side2Feature']),
]:
    _add('joint_instances', col, 'AssemblyDefinition.Joints[].' + target, target, True, output)
for col, element, note in [('joint_type', 'TypeCode', ''), ('description', 'GeometricDescription', ''), ('count', 'SourceCount', 'Compared with the number of joints of this type.')]:
    _add('joint_types', col, 'AssemblyDefinition.TypeCatalog.JointTypes[].' + col, col, col == 'joint_type',
         ['AssemblyDefinition/TypeCatalog/JointTypes/<type>/' + element], note)

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
