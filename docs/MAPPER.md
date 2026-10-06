# Excel to AAS Mapping

The mapping page (`/mapper`, linked as **Map Excel to Any AAS**) imports data from any Excel workbook into an existing AAS, an official submodel template, or an AAS built on the page. The assembly converter on the start page is unchanged.

## Workflow

1. **Excel data.** Load an `.xlsx` file. Every worksheet whose first row holds unique column names can be used; sheets with blank or duplicate headers are listed as not usable. Formula cells use the result last saved by Excel; formulas without a saved result cannot be imported.
2. **Target AAS.**
   - **Open AASX** opens an existing package (AAS V3.0, XML or JSON), for example one made in AASX Package Explorer, received from a supplier, or exported earlier. If it holds several shells, select one.
   - **New AAS** creates an empty shell with a generated ID and an optional global asset ID.
   - **Add submodel from template** copies a submodel from the template library as an instance: new ID, kind `Instance`, example values removed and concept descriptions copied. See [aas_templates/README.md](../aas_templates/README.md) for official IDTA templates.
   - **Add own submodel** and **Add own element** fill gaps that no template covers. Elements can be deleted from the structure view.
   - The structure view shows each template element's cardinality: `One`, `OneToMany`, `ZeroToOne` or `ZeroToMany`.
3. **Mapping.**
   - **Single value**: one cell, chosen by Excel row number or by a value in another column (for key/value sheets), is written to a property, a multi-language property, or the global asset ID of an entity.
   - **Table rows**: each Excel row becomes one collection or entity inside the target container. Rows are named by a column (names that are not valid idShorts get a hash suffix) or by row number.
     - With **Row structure**, each row copies an existing collection or entity, such as the first entry of a template list or the `Node` entity of the Hierarchical Structures template. Columns are written to its properties; for entities, a column can also set the **entity global asset ID**.
     - Without a row structure, columns become new properties with the chosen value types.
     - **Update by name** keeps existing rows and updates matching ones; **Replace all** rebuilds the container. Lists are always rebuilt.
     - Two Excel rows that map to the same name are an error.
     - A worksheet without data rows leaves the target unchanged, so an empty sheet never wipes a template's row structure.
   - Mappings can be downloaded as JSON and loaded again for the next workbook of the same layout.
4. **Check and import.**
   - **Check** converts every cell to the target value type and reports errors with worksheet, row and column. Import is refused while errors remain.
   - **Leave empty optional template elements out of the export** (on by default) removes template elements with cardinality `ZeroToOne` or `ZeroToMany` that received no data, including placeholders such as `ArbitraryProperty` and unused relationships. A repeating template element that served as the row structure is also left out once filled entries of the same kind exist.
   - Check lists **mandatory template elements** (`One`, `OneToMany`) that are still empty, and entities that still use a template placeholder asset ID (`https://admin-shell.io/...`). Repeated findings in lists are counted in one line. These are warnings; they do not block the import.
   - The result is a new AASX download; the opened file is never changed.

The structure view shows the first 25 children of each container followed by a line such as "… 35 more"; counts in the check report always cover all rows.

## What is preserved

When an existing package is imported into, only its AAS model part is rewritten, in its original format (XML or JSON). Thumbnails, attachments and all other package parts are copied unchanged. Metamodel problems that the opened AAS already had are listed but do not block the import; problems introduced by the import do.

Template qualifiers such as `SMT/Cardinality` are kept in the working copy for the checks above and removed from every export, because the metamodel allows them only in submodel templates (constraint AASd-129). The working copy also keeps empty template elements, so a mapping can be imported again later.

## Value conversion

Numbers in text cells must use plain decimal notation. Booleans accept `true`, `false`, `1` and `0`. Excel dates are written as `xs:date`, `xs:dateTime` or `xs:time`. Every value is checked against its XSD type before export.

## Limits

- Only AAS V3.0 models are supported. V1/V2 models must be saved as V3.0 in AASX Package Explorer first; for IDTA templates, use the files without `forAASMetamodelV3.1`.
- Values can be written to properties, multi-language properties and entity global asset IDs. Files, ranges, references and relationships cannot be filled from Excel; for example, the `HasPart`/`IsPartOf` relationships of Hierarchical Structures are not generated.
- Cells are written as they are. Values cannot be computed or combined, so columns such as asset IDs must exist in the workbook.
- Cardinality checks only apply to templates added on this page. An exported AAS no longer carries template qualifiers, so its cardinalities are not known when it is opened again.
