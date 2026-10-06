# Excel to AAS Mapping

The mapping page (`/mapper`, linked as **Map Excel to Any AAS**) imports data from any Excel workbook into an existing AAS, an official submodel template, or an AAS built on the page. The assembly converter on the start page is unchanged.

## Workflow

1. **Excel data.** Load an `.xlsx` file. Every worksheet whose first row holds unique column names can be used; sheets with blank or duplicate headers are listed as not usable. Formula cells use the result last saved by Excel; formulas without a saved result cannot be imported.
2. **Target AAS.**
   - **Open AASX** opens an existing package (AAS V3.0, XML or JSON), for example one made in AASX Package Explorer, received from a supplier, or exported earlier. If it holds several shells, select one.
   - **New AAS** creates an empty shell with a generated ID and an optional global asset ID.
   - **Add submodel from template** copies a submodel from the template library as an instance: new ID, kind `Instance`, example values and template-only qualifiers removed, concept descriptions copied. See [aas_templates/README.md](../aas_templates/README.md) for official IDTA templates.
   - **Add own submodel** and **Add own element** fill gaps that no template covers. Elements can be deleted from the structure view.
3. **Mapping.**
   - **Single value**: one cell, chosen by Excel row number or by a value in another column (for key/value sheets), is written to a property or multi-language property.
   - **Table rows**: each Excel row becomes one collection inside the target container. Rows are named by a column (names that are not valid idShorts get a hash suffix) or by row number. With **Row structure**, each row copies an existing collection, such as the first entry of a template list, and columns are written to its properties. Otherwise columns become new properties with the chosen value types. **Update by name** keeps existing rows and updates matching ones; **Replace all** rebuilds the container. Lists are always rebuilt. Two Excel rows that map to the same name are an error. A worksheet without data rows leaves the target unchanged, so an empty sheet never wipes a template's row structure.
   - Mappings can be downloaded as JSON and loaded again for the next workbook of the same layout.
4. **Check and import.** Check converts every cell to the target value type and reports errors with worksheet, row and column. Import is refused while errors remain. The result is a new AASX download; the opened file is never changed.

The structure view shows the first 25 children of each container followed by a line such as "… 35 more"; counts in the check report always cover all rows. AAS V1/V2 packages must first be saved as V3.0 in AASX Package Explorer.

## What is preserved

When an existing package is imported into, only its AAS model part is rewritten, in its original format (XML or JSON). Thumbnails, attachments and all other package parts are copied unchanged. Metamodel problems that the opened AAS already had are listed but do not block the import; problems introduced by the import do.

## Value conversion

Numbers in text cells must use plain decimal notation. Booleans accept `true`, `false`, `1` and `0`. Excel dates are written as `xs:date`, `xs:dateTime` or `xs:time`. Every value is checked against its XSD type before export.
