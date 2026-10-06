# Template library

Submodel templates in this folder (`.aasx` or AAS V3.0 `.json`) appear under **Add submodel from template** on the mapping page (`/mapper`). Files added with **Add template file** are stored here too.

Official templates are published by the IDTA at <https://github.com/admin-shell-io/submodel-templates> (folder `published`), for example Digital Nameplate, Technical Data, Hierarchical Structures (Bill of Material) and Contact Information. Record the template version you use.

When a template is added to an AAS, the tool creates an instance: new submodel ID, kind `Instance`, example values removed (unless kept explicitly) and template-only qualifiers such as `SMT/Cardinality` removed. Where a template lacks fields, add your own elements on the mapping page.

Template files are local data and are not committed (`*.aasx` is ignored).
