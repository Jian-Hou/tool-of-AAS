# Tests

After running `setup.bat`, run these checks from the project root:

```powershell
.\.venv\Scripts\python.exe -m pip check
.\.venv\Scripts\python.exe -B -m unittest discover -s tests -v
```

All tests use synthetic data; no private workbook is required.

- `test_audit_regressions.py` checks joint counts, catalog fields, reference targets, whitespace handling, self-joints, number and name limits, and that every destination shown in the field-mapping tab exists in the output.
- `test_workbook_export.py` reads a generated `.xlsx`, ignores unrelated worksheets, and checks that the exported AASX reads back with the BaSyx SDK.
- `test_app_storage.py` checks that a new upload replaces the stored input and that old exports are pruned.
- `test_aas_mapper.py` checks the generic mapping: templates, own elements, single values, table rows, existing XML packages patched in place, value conversion errors and pre-existing metamodel problems.
- `test_mapper_app.py` runs the mapping page's API from Excel upload to AASX download.
