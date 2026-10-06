# Excel to AASX

A local web tool that turns Excel data into Asset Administration Shell (AAS) V3.0 packages. Everything runs on your machine.

## Features

The tool has two pages.

**1. Assembly converter** (start page): converts an assembly workbook in the fixed format below into a new AAS.
- Checks the workbook before export. Errors block the export; missing information is listed as warnings.
- Exports either a draft that records missing information, or a strict version that requires complete data.
- The AAS contains the submodels `TechnicalData` (IDTA Technical Data semantics) and `AssemblyDefinition` (components, positions, type catalogs, joints, data quality), and embeds the original workbook.
- The field-mapping tab lets you choose the source column for each field and shows where it is written in the AAS.
- Command line: `scripts/generate.py` converts; `scripts/qa.py` compares an export field by field with the workbook.

**2. Excel to AAS mapping** (`/mapper`, button **Map Excel to Any AAS**): imports any workbook into an AAS you choose.
- **Excel:** any sheet whose first row holds column names.
- **Target:** open an existing AASX (AAS V3.0, XML or JSON) or create a new AAS.
- **Structure:** add submodels from official IDTA templates or your own template files, and add your own submodels and elements where a template is not enough. The structure view shows each template element's cardinality.
- **Mapping:**
  - *Single value*: one cell, chosen by row number or by a key column, goes into a property, a multi-language property or an entity's global asset ID.
  - *Table rows*: each row becomes a collection or an entity. Rows can follow a template's structure, for example list entries or the `Node` entities of a bill of materials.
  - Mappings can be saved and loaded again for the next workbook.
- **Check:**
  - Every cell is converted to its target type; invalid cells are reported by sheet, row and column, and block the import.
  - Mandatory template fields that are still empty are listed as warnings.
- **Export:**
  - A new AASX download. Empty optional template elements are left out by default.
  - When an existing package is updated, attachments, thumbnails and other submodels are kept unchanged.

## Quick start

1. Install **Python 3.12** on Windows and run `setup.bat`.
2. Run `scripts/start.bat`; the browser opens the start page. If port 5000 is busy, the terminal shows the address used. Press **Ctrl+C** there to stop.
3. Assembly data: load the `.xlsx`, review the checks, then click **Export Draft** or **Export with Strict Validation**.
4. Other data: click **Map Excel to Any AAS**, load the Excel file, open or create the AAS, add mappings, then **Check** and **Import into AAS**.

## Assembly input format

Use `.xlsx` with column names in the first row. Sheet names are fixed, and column names can be remapped on the field-mapping tab. Other sheets, such as notes, are ignored. Limits per sheet: 25,000 rows and 128 columns; the file may be up to 12 MB.

| Sheet | Columns |
|---|---|
| `components` (required) | `assembly_id` (unique integer), `label` (unique; links all sheets), `tag` (optional), `type` (`pipeline`, `elbow`, `blackbox` or `tank`), `shape_type`, `coord` `(x,y,z)`, `placement` `(qx,qy,qz,qw)\|xmin,ymin,zmin,xmax,ymax,zmax` |
| `<type>_instances` | `assembly_id`, `label`, `shape_type`, `params` such as `length_mm=100;dims_mm=1,2,3` |
| `<type>_types` (optional) | `shape_type`, `geometric_description`, `params_needed` (names separated by `;`), `count` |
| `joint_instances` (required, may be empty) | `joint_id`, `joint_type`, `side1_id`, `side1_sub`, `side2_id`, `side2_sub` |
| `joint_types` (optional) | `joint_type`, `description`, `count` |

Rules:
- **Values:** numbers in text cells use plain decimal notation, such as `12.5` or `1.2e-3`; formulas are not allowed in these sheets.
- **Placement:** the quaternion must have length 1, and each bounding-box minimum must not exceed its maximum.
- **Parameters:** names ending in `_mm` are millimetres; enter other units in the tool.
- **Joints:** endpoints are component labels or the assembly root. `GroundedJoint` may leave side 2 empty; other joints need both endpoints and features.
- **Counts and conventions:** counts are compared with the number of instances. Coordinate conventions (unit, axes, quaternion order) are entered in the tool and recorded as unconfirmed until you confirm them.

## Official templates

Put IDTA submodel templates into the folder `aas_templates/`, or add them on the mapping page with **Add template file**. Template files are not committed. Download them from <https://github.com/admin-shell-io/submodel-templates> (folder `published`) and use the AAS V3.0 files, not those ending in `forAASMetamodelV3.1`. Tested versions:
- Digital Nameplate 3.0.2
- Technical Data 2.0.2
- Hierarchical Structures enabling Bills of Material 1.1.2
- Contact Information 1.0.2

## Testing

Run `.\.venv\Scripts\python.exe -B -m unittest discover -s tests`. The 40 tests use synthetic data; the five tests for official templates are skipped when the template files are missing.

Also tested during development (October 2026):
- **Real project workbook (301 components):**
  - The assembly export matched the previous version exactly, and the audit matched all 9,404 fields.
  - On the mapping page, the workbook was imported into a custom submodel, the Technical Data template and a 301-node bill of materials, with no cell mismatches.
- **Randomized tests:**
  - Several thousand random workbooks, AAS packages, templates and mappings, with every imported value read back through the BaSyx SDK;
  - fault injection;
  - hostile uploads (XML entity attacks, path traversal, broken packages);
  - 20,000 rows.

## Limitations

- **AAS versions:** only V3.0 is supported; save V1/V2 models as V3.0 in AASX Package Explorer first.
- **Mapping targets:**
  - The mapping writes properties, multi-language properties and entity asset IDs.
  - Files, ranges, references and relationships are not filled. For example, a bill of materials gets no `HasPart` links, and the Nameplate's `MarkingFile` stays empty.
- **Cell values:**
  - Cells are imported as they are, so values such as asset IDs must exist as columns.
  - Values with more than about 17 significant digits are rounded to the nearest double.
- **Cardinality:** an exported AAS no longer carries template cardinality, so the mandatory-field check only covers templates added in the current mapping project.
- **Structure view:** only the first 25 entries of each container are shown.
- **Assembly AAS ID:** it is derived from the workbook content and changes whenever the file changes; enter a global asset ID for a stable identity.
- **Assembly audit:** it compares data from the workbook only, not values from settings such as manufacturer data.
- **Local use:** the tool is for local, single-user use and listens on `127.0.0.1` only. Old session folders in `.state/` are not cleaned up automatically, and output paths over 260 characters need Windows long paths enabled.
