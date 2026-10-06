# Template library

Submodel templates in this folder (`.aasx` or AAS V3.0 `.json`) appear under **Add submodel from template** on the mapping page (`/mapper`). Files added with **Add template file** are stored here too.

Official templates are published by the IDTA at <https://github.com/admin-shell-io/submodel-templates> (folder `published`). Use the AAS V3.0 files, not those ending in `_forAASMetamodelV3.1`. These versions were downloaded and tested with the tool (see `tests/test_official_templates.py`):

| Template | Source file in `published/` | Local file name |
|---|---|---|
| Digital Nameplate 3.0.2 | `Digital nameplate/3/0/2/IDTA 02006-3-0-2_Template_Digital Nameplate.aasx` | `IDTA_02006-3-0-2_DigitalNameplate.aasx` |
| Technical Data 2.0.2 | `Technical_Data/2/0/2/IDTA 02003_2-0-2_Template_TechnicalData.aasx` | `IDTA_02003-2-0-2_TechnicalData.aasx` |
| Hierarchical Structures enabling Bills of Material 1.1.2 | `Hierarchical Structures enabling Bills of Material/1/1/2/IDTA 02011-1-1-2_Template_HSEBoM.aasx` | `IDTA_02011-1-1-2_HierarchicalStructures.aasx` |
| Contact Information 1.0.2 | `Contact Information/1/0/2/IDTA 02002-1-0-2_Template_ContactInformation.aasx` | `IDTA_02002-1-0-2_ContactInformation.aasx` |

When a template is added to an AAS, the tool creates an instance:

- The submodel gets a new ID and kind `Instance`.
- Example values are removed unless you choose to keep them.
- Cardinality information is used for the checks and removed from the export.

Where a template lacks fields, add your own elements on the mapping page.

Template files are local data and are not committed (`*.aasx` and everything except this README are ignored). The IDTA publishes the templates under its own license terms; keep the source and version with any copies you share.
