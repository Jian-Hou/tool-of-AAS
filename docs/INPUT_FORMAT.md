# Excel Input Format

Use the first row as column headers. Do not use merged headers or formulas. Sheet names are fixed; source columns can be mapped. Instances join by `label`, and their `assembly_id` values must agree.

Only `.xlsx` is supported. Limits: 12 MB compressed, 64 MB uncompressed, 32 sheets, 25,000 rows per sheet, and 128 columns per sheet.

## Components: components

| Column | Meaning / example |
|---|---|
| assembly_id | Unique integer ID, such as `1` |
| label | Unique label, such as `demo_pipe_001` |
| tag | Optional asset tag |
| type | `pipeline`, `elbow`, `blackbox`, or `tank` |
| shape_type | Original classification code, such as `P01` |
| coord | Three position values, such as `(0,0,0)` |
| placement | Quaternion and bounding box, such as `(0,0,0,1)\|0,0,0,100,10,10` |

The quaternion uses the selected XYZW or WXYZ order and is exported as XYZW. The bounding box uses XMin, YMin, ZMin, XMax, YMax, ZMax. Quaternion norm must be 1 within 0.001; each minimum must not exceed its maximum.

Default XYZW is an initial interpretation, not source confirmation. Units, origin, axes, handedness, rotation direction, and reference frame must come from reliable documentation. Unconfirmed conventions are recorded in drafts.

## Instance parameters: *_instances

Provide `pipeline_instances`, `elbow_instances`, `blackbox_instances`, or `tank_instances` for each family present.

| Column | Meaning / example |
|---|---|
| assembly_id | Same integer ID as the component |
| label | Same label as the component |
| shape_type | Instance geometry code |
| params | `length_mm=100;outer_diameter_mm=10;inner_diameter_mm=8` |

Separate parameters with semicolons and names from values with `=`. Multiple values use commas. Invalid numbers, duplicates, and normalized name collisions are rejected. Parameters must not be empty.

Names ending in `_mm` use mm. Enter other units in the interface; leave unknown units blank. Original parameter text is preserved alongside expanded values.

Blackbox classification and geometry codes may differ. Both are retained; type references use the instance geometry code.

## Type catalogs: *_types

| Column | Meaning |
|---|---|
| shape_type | Geometry code |
| geometric_description | Optional description |
| params_needed | Required names separated by semicolons |
| count | Optional integer source count, checked against instances |

Without a formal definition, the tool records observed types and parameters and reports incomplete information. It does not infer required parameters. Strict export requires complete definitions.

## Joints: joint_instances

| Column | Meaning |
|---|---|
| joint_id | Unique integer ID |
| joint_type | Type matched to its catalog |
| side1_id | First component label, or root `Assembly001` / current assembly name |
| side1_sub | First joint feature |
| side2_id | Second component label or root name |
| side2_sub | Second joint feature |

Endpoints use labels, not numeric IDs. `GroundedJoint` permits an empty second endpoint. Other joints require both endpoint labels and features.

`components` and `joint_instances` are mandatory. With no joints, retain an empty sheet with all joint headers; missing connections produce a warning.

`joint_types` uses `joint_type`, `description`, and `count`; only `joint_type` is mandatory.

## Minimal manual example

This is a format example, not a real product. Create five sheets with the following names, headers, and rows. Save as `.xlsx` and try draft export.

**components**

| assembly_id | label | tag | type | shape_type | coord | placement |
|---|---|---|---|---|---|---|
| 1 | demo_pipe_001 | DEMO | pipeline | P01 | (0,0,0) | (0,0,0,1)\|0,0,0,100,10,10 |

**pipeline_instances**

| assembly_id | label | shape_type | params |
|---|---|---|---|
| 1 | demo_pipe_001 | P01 | length_mm=100;outer_diameter_mm=10;inner_diameter_mm=8 |

**pipeline_types**

| shape_type | geometric_description | params_needed | count |
|---|---|---|---|
| P01 | Synthetic demonstration pipe | length_mm;outer_diameter_mm;inner_diameter_mm | 1 |

**joint_instances**

| joint_id | joint_type | side1_id | side1_sub | side2_id | side2_sub |
|---|---|---|---|---|---|
| 1 | GroundedJoint | demo_pipe_001 | | | |

**joint_types**

| joint_type | description | count |
|---|---|---|
| GroundedJoint | Synthetic grounding example | 1 |

Enter the actual `|` character in the placement cell without the Markdown escape backslash. Manufacturer information and confirmed coordinate conventions are not supplied, so the expected status is `draft`.
