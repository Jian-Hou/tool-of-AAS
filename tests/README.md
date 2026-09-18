# Tests

After running `setup.bat`, run these checks from the project root:

```powershell
.\.venv\Scripts\python.exe -m pip check
.\.venv\Scripts\python.exe -B -m unittest discover -s tests -v
.\.venv\Scripts\python.exe -B scripts\test_serialize.py
```

The seven synthetic audit tests check joint counts, catalog fields, and reference targets. The serialization test checks AASX property read-back. No private workbook is required.
