# Excel to AASX

A local web tool that turns Excel data into Asset Administration Shell (AAS) V3.0 packages. Everything runs on your machine; nothing is uploaded to a server.

Status: October 2026. This page describes what the tool can do today and what is still missing.

## What it can do

The tool has two pages.

### 1. Assembly converter (start page)

Converts workbooks in the fixed [assembly input format](docs/INPUT_FORMAT.md) into a new assembly AAS.

- Validates the workbook before export: required sheets and columns, unique IDs and labels, numbers, quaternions, bounding boxes, joint endpoints, type catalogs and counts.
- Unrelated worksheets, such as notes, are ignored and listed instead of rejecting the file.
- The field-mapping tab lets you pick the source column for each field and shows where the field is written in the AAS. An automated test keeps these paths in line with the real output.
- **Draft export** keeps missing information visible in the output; **strict export** requires complete data.
- Output: one AAS with the submodels `TechnicalData` (IDTA Technical Data semantics) and `AssemblyDefinition` (project-specific), parameter concept descriptions, and the original workbook embedded for traceability.
- An independent audit (`scripts/qa.py`) compares the exported AASX field by field with the workbook.
- Command line: `scripts/generate.py` (convert) and `scripts/qa.py` (audit).

### 2. Excel to AAS mapping (`/mapper`)

Imports data from any workbook into an AAS you choose. See the [mapping guide](docs/MAPPER.md).

- **Excel:** any sheet whose first row holds column names. Column types are guessed; formula cells use the result last saved by Excel.
- **Target AAS:** open an existing AASX (AAS V3.0, XML or JSON; packages with several shells are supported) or create a new AAS.
- **Official and own structures:** add submodels from the [template library](aas_templates/README.md), for example IDTA submodel templates, and add your own submodels and elements where a template is not enough.
- **Mappings:**
  - *Single value*: one cell, selected by row number or by a key column, goes into one property.
  - *Table rows*: each Excel row becomes one collection, either with new properties or following a template's row structure, such as the entries of a list.
- **Check before import:** every cell is converted to its target type. Errors name the mapping, sheet, row and column, and the import is refused while errors remain.
- **Safe output:** the result is a new AASX download. When importing into an existing package, only its AAS model part is rewritten, in its original format. Attachments, thumbnails and other submodels are copied unchanged.
- **Reuse:** mappings can be downloaded and loaded again for the next workbook with the same layout.

## Quick start

1. Install **Python 3.12** on Windows.
2. Run `setup.bat` to install dependencies into `.venv`.
3. Run `scripts/start.bat`. The browser opens the start page. If port 5000 is busy, another free port is used and printed in the terminal.
4. Assembly data: load the `.xlsx`, review the checks, then click **Export Draft** or **Export with Strict Validation**.
5. Any other data: click **Map Excel to Any AAS**, then load Excel, open or create the AAS, add mappings, **Check**, and **Import into AAS**.
6. Press **Ctrl+C** in the terminal to stop the tool.

## Testing status

Run the automated tests from the project root after `setup.bat`:

```powershell
.\.venv\Scripts\python.exe -B -m unittest discover -s tests -v
```

The 33 tests use synthetic data only; see [tests/README.md](tests/README.md). In addition, randomized tests were run during development but are not part of the repository:

- **Assembly converter:** hundreds of random workbooks, single-fault injection, audit tamper detection, boundary values, and 20,000 components.
- **Mapper:**
  - 1,600 random AAS packages, templates, workbooks and mappings, with every imported value read back through the BaSyx SDK;
  - fault injection;
  - random structure edits;
  - malformed and hostile requests (XML entity expansion, external entities, path traversal, broken packages) without a server error;
  - 20,000 rows.

## Known limitations and open issues

### Mapping page

- **Template cardinality is not enforced.** Template-only qualifiers such as `SMT/Cardinality` must be removed from instances, so the tool does not warn when a mandatory template field stays empty.
- **Value targets** are properties and multi-language properties. Files, ranges and references cannot be filled from Excel.
- **Own elements** can be properties, multi-language properties, collections and lists of collections.
- **Supported packages:** AAS V3.0 only. V1/V2 models must first be saved as V3.0 in AASX Package Explorer. A package must contain exactly one AAS model part.
- **Large containers:** the structure view shows the first 25 children of each container, and deeper entries cannot be selected as single-value targets.
- **Import button:** it stays enabled after a failed check. The server refuses the import, and nothing is written.
- **Template library:** it starts empty. Official IDTA templates must be added from <https://github.com/admin-shell-io/submodel-templates>.
- **Not yet tested with real files:** neither official IDTA template files nor real project workbooks have been tested.

### Assembly converter

- **One AAS per file version.** The AAS ID is derived from the workbook's SHA-256 hash and the assembly name, so changing any cell produces a different AAS ID. The optional global asset ID stays stable.
- **Project-specific submodel.** `AssemblyDefinition` has no standard IDTA semantics.
- **Limited audit scope.** The audit checks source-derived fields only. Values from settings or derived counts, such as manufacturer data, coordinate conventions and instance counts, are not compared.
- **Number formats.** Numbers in text cells must use plain decimal notation. Values with more than about 17 significant digits are rounded to the nearest double.

### General

- **State storage.** Each browser session keeps its files under `.state/`. Old session folders are not deleted automatically; each session keeps only its latest input and its last 5 exports.
- **Long paths.** Output paths longer than 260 characters fail unless Windows long paths are enabled.
- **Local use only.** The built-in server is for local use and binds to `127.0.0.1` only.

## Project layout

| Path | Content |
|---|---|
| `scripts/app.py` | Web application and assembly converter routes |
| `scripts/conversion.py` | Assembly workbook reading, validation and AASX export |
| `scripts/qa.py` | Independent audit of assembly exports |
| `scripts/aas_mapper.py` | Generic workbook reading, AAS packages, structure edits and mapping engine |
| `scripts/mapper_routes.py` | Web routes of the mapping page |
| `scripts/templates/`, `scripts/static/` | Pages, scripts and styles |
| `aas_templates/` | Local template library (template files are not committed) |
| `docs/` | Input format and mapping guide |
| `tests/` | Automated tests |
